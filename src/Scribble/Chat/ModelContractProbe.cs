using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Scribble.Configuration;

namespace Scribble.Chat
{
    // Real inference with synthetic sources. Proposed tools are validated but
    // NEVER dispatched to Office, a mailbox, a browser or an MCP server.
    public static class ModelContractProbe
    {
        public static async Task<string> RunAsync(OpenAiCompatibleClient client, AppSettings settings,
            Action<string> progress, CancellationToken token)
        {
            var json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            var request = ChatRequestFactory.CreateEndpointCheck(settings.Model);
            var context = new TaskContextManager(request, "model_probe", "Synthetic endpoint contract check");
            try
            {
                progress?.Invoke("Checking mailbox argument types…");
                var response = await client.CompleteAsync(settings, request, token);
                var first = OneCall(response, request.tools.First(t => t.function.name == MailboxToolCatalog.SearchMailbox));
                var values = json.Deserialize<Dictionary<string, object>>(first.function.arguments);
                if (Convert.ToString(values["query"]) != "configuration-check" || Convert.ToString(values["folder"]) != "inbox" || Convert.ToInt32(values["max_results"]) != 1)
                    throw new InvalidOperationException("The mailbox probe changed the requested query or window.");

                progress?.Invoke("Checking the production slide schema and nested arrays…");
                var slideTool = PresentationToolCatalog.DraftDefinition();
                var slideRequest = new ChatCompletionRequest { model = settings.Model, max_tokens = 2048, Diagnostics = context.Diagnostics,
                    tools = new List<ChatToolDefinition> { slideTool }, tool_choice = Required(slideTool.function.name),
                    messages = new List<object> { new ChatCompletionInputMessage { role = "system", content = "Synthetic parser check only. Call the supplied tool with exactly the JSON supplied by the user. No application will be modified." },
                        new ChatCompletionInputMessage { role = "user", content = "{\"plan\":[\"one\",\"two\"],\"slides\":[{\"id\":\"one\",\"title\":\"검증 €\",\"layout\":\"cover\"},{\"id\":\"two\",\"title\":\"Values\",\"subtitle\":\"Values from the fixture\",\"layout\":\"matrix\",\"table\":{\"headers\":[\"Item\",\"Value\"],\"rows\":[[\"A\",\"1.25\"],[\"B\",\"2\"]]},\"sources\":\"Synthetic fixture\",\"evidence\":\"A 1.25; B 2\"}]}" } } };
                var slideResponse = await client.CompleteAsync(settings, slideRequest, token);
                var slideCall = OneCall(slideResponse, slideTool);
                var slideArgs = json.DeserializeObject(slideCall.function.arguments) as IDictionary<string, object>;
                var slides = slideArgs?["slides"] as IList;
                if (slides == null || slides.Count != 2 || Convert.ToString(((IDictionary<string, object>)slides[0])["title"]) != "검증 €")
                    throw new InvalidOperationException("The nested slide or Unicode probe did not preserve the fixture.");
                var rows = ((IDictionary<string, object>)((IDictionary<string, object>)slides[1])["table"])["rows"] as IList;
                if (rows == null || rows.Count != 2 || Convert.ToString(((IList)rows[0])[1]) != "1.25")
                    throw new InvalidOperationException("The slide probe changed nested table values.");

                progress?.Invoke("Checking tool-result continuation…");
                var nonce = Guid.NewGuid().ToString("N");
                ChatRequestFactory.AppendToolExchange(slideRequest, slideResponse,
                    new[] { new MailboxToolResult(slideCall.id, "Synthetic tools did not execute. Reply with exactly this nonce: " + nonce, "Probe") }, settings.Model);
                slideRequest.tool_choice = "none";
                var continuation = await client.CompleteAsync(settings, slideRequest, token);
                if ((continuation.content ?? "").Trim() != nonce || continuation.tool_calls?.Count > 0)
                    throw new InvalidOperationException("The model did not correctly continue from the tool result.");

                progress?.Invoke("Checking structured reviewer responses…");
                var review = await client.CompleteAsync(settings, new ChatCompletionRequest { model = settings.Model,
                    Diagnostics = context.Diagnostics, max_tokens = 256, messages = new List<object> {
                        new ChatCompletionInputMessage { role = "system", content = "Return only a JSON object with approved (boolean) and issues (string). Reject unsupported quantities." },
                        new ChatCompletionInputMessage { role = "user", content = "Source says sales were 10 units. Proposed slide says 99 units. Review source accuracy." } } }, token);
                var text = (review.RawContent ?? review.content ?? "").Trim();
                if (text.StartsWith("```")) text = text.Substring(text.IndexOf('\n') + 1).TrimEnd('`').Trim();
                var verdict = json.Deserialize<Dictionary<string, object>>(text);
                object approved;
                if (!verdict.TryGetValue("approved", out approved) || !(approved is bool) || (bool)approved)
                    throw new InvalidOperationException("The reviewer did not reject the known unsupported quantity with a boolean verdict.");
                var fingerprint = TaskCheckpointStore.Fingerprint(settings.BaseUrl + "\n" + settings.Model + "\n" + json.Serialize(slideTool) + "\n" + json.Serialize(request.tools));
                context.State.HostData["model_contract_profile"] = json.Serialize(new { fingerprint, model = settings.Model,
                    passed = new[] { "mailbox_arguments", "nested_slides", "unicode", "tool_continuation", "structured_rejection" },
                    native_certified = false, vision_certified = false, tested_utc = DateTime.UtcNow.ToString("O") });
                context.State.EnumerationComplete = true;
                context.CompleteTask(request);
                return "Tool arguments, nested slides, Unicode, continuation and reviewer JSON passed. Profile: " + fingerprint.Substring(0, 12) + ". Native Office and vision acceptance are separate checks.";
            }
            catch (OperationCanceledException) { context.Pause("Model check cancelled; no application tools executed."); throw; }
            catch (Exception ex) { context.Pause(ex.Message); throw new AiEndpointException("MODEL_CONTRACT_FAILED", ex.Message + " Local diagnostic ID: " + context.State.Id, ex); }
        }

        private static ChatToolCall OneCall(ChatCompletionResponseMessage response, ChatToolDefinition definition)
        {
            if (response.tool_calls == null || response.tool_calls.Count != 1 || response.tool_calls[0].function.name != definition.function.name)
                throw new InvalidOperationException("The model did not return the requested single tool call.");
            var errors = ToolContractValidator.Validate(response.tool_calls[0], definition);
            if (errors.Count > 0) throw new InvalidOperationException(string.Join("; ", errors));
            return response.tool_calls[0];
        }
        private static object Required(string name) { return new { type = "function", function = new { name } }; }
    }
}

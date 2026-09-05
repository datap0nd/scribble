using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Scribble.Configuration;
using Scribble.Outlook;
using Scribble.Security;

namespace Scribble.Chat
{
    public sealed class BrowserTopicInfo
    {
        public string Id { get; set; }

        public string Name { get; set; }

        public string Binding { get; set; }

        public bool Available { get; set; }
    }

    public sealed class BrowserChatResult
    {
        public BrowserChatResult(
            string content,
            string model,
            bool screenshotUsed)
            : this(
                content,
                model,
                screenshotUsed,
                null,
                null,
                null)
        {
        }

        public BrowserChatResult(
            string content,
            string model,
            bool screenshotUsed,
            string pendingAssistantContent,
            IReadOnlyList<ChatToolCall> pendingCalls,
            IReadOnlyList<BrowserExchangeResult> hostResults)
        {
            Content = content ?? string.Empty;
            Model = model ?? string.Empty;
            ScreenshotUsed = screenshotUsed;
            PendingAssistantContent =
                pendingAssistantContent ?? string.Empty;
            PendingCalls = pendingCalls ?? new ChatToolCall[0];
            HostResults = hostResults ??
                new BrowserExchangeResult[0];
        }

        public string Content { get; }

        public string Model { get; }

        public bool ScreenshotUsed { get; }

        // The assistant turn and calls the extension must complete
        // (navigation or page reads) before the loop can continue.
        public string PendingAssistantContent { get; }

        public IReadOnlyList<ChatToolCall> PendingCalls { get; }

        // Results for calls in PendingCalls the host already
        // executed itself (MCP tools, the unsent Outlook draft).
        public IReadOnlyList<BrowserExchangeResult> HostResults
        {
            get;
        }

        public bool HasPendingCalls
        {
            get { return PendingCalls.Count > 0; }
        }
    }

    // Runs the browser host's model loop. The extension executes
    // navigation and page reads in the user's own visible tab and
    // replays the results; the host executes user-configured MCP
    // tools, the unsent-Outlook-draft tool, and the
    // unsaved-workbook tool (each once per request). No
    // click, form, credential, or page-mutation capability exists
    // anywhere in this process, and it can never send email.
    public sealed class BrowserChatService : IDisposable
    {
        public const int MaxBrowserStagnantCalls = 20;
        // An emergency cost/safety fuse. Normal completion is governed by
        // observed progress rather than a small fixed action allowance.
        // Kept for binary compatibility; no task-ending round fuse.
        public const int MaxBrowserEmergencyRounds = int.MaxValue;
        public const int MaxBrowserToolCallsPerRound = 4;
        public const int MaxStateChangingBrowserCallsPerRound = 1;

        private readonly AppSettings _settings;
        private readonly OpenAiCompatibleClient _client;
        private readonly McpToolHost _mcpTools;

        public BrowserChatService()
            : this(new SettingsStore().Load())
        {
        }

        public BrowserChatService(AppSettings settings)
        {
            _settings = settings ?? new AppSettings();
            _settings.ApplyLimits();
            TopicToolHost.CleanupExpiredPersistentSessions();
            _client = new OpenAiCompatibleClient();
            _mcpTools = new McpToolHost(
                _settings.McpServers,
                true);
        }

        public bool IsConfigured
        {
            get { return _settings.IsConfigured; }
        }

        public string Model
        {
            get { return _settings.Model ?? string.Empty; }
        }

        public bool SupportsVision
        {
            get
            {
                if (ModelCatalog.IsVisionCapable(Model))
                {
                    return true;
                }

                if (!_settings.SwitchToVisionModelForImages)
                {
                    return false;
                }

                return ModelCatalog.FindBestVisionModel(
                    _settings.DiscoveredModels).Length > 0;
            }
        }

        public IReadOnlyList<BrowserTopicInfo> Topics
        {
            get
            {
                return _settings.Topics.Select(topic =>
                    new BrowserTopicInfo
                    {
                        Id = topic.Id,
                        Name = topic.Name,
                        Binding = TopicBinding(topic),
                        Available = Directory.Exists(topic.FolderPath)
                    }).ToArray();
            }
        }

        public async Task<BrowserChatResult> CompleteAsync(
            IReadOnlyList<ChatTurn> history,
            string prompt,
            string title,
            string url,
            string selection,
            string pageText,
            string links,
            string screenshotDataUrl,
            IReadOnlyList<BrowserExchangeTurn> exchange,
            string chatId,
            string turnId,
            string topicId,
            string topicBinding,
            CancellationToken cancellationToken)
        {
            if (!IsConfigured)
            {
                throw new AiEndpointException(
                    "CONFIGURATION_INCOMPLETE",
                    "I need you to open Scribble Settings and configure the endpoint, model, and API key first.");
            }

            var safePrompt = TextBoundary.PlainText(
                prompt,
                TextBoundary.MaxUserPromptCharacters);
            if (safePrompt.Length == 0)
            {
                throw new AiEndpointException(
                    "PROMPT_EMPTY",
                    "Type a message first.");
            }

            var safeScreenshot =
                BrowserChatRequestFactory.NormalizeScreenshot(
                    screenshotDataUrl);
            if (!string.IsNullOrWhiteSpace(screenshotDataUrl) &&
                safeScreenshot.Length == 0)
            {
                throw new AiEndpointException(
                    "SCREENSHOT_INVALID",
                    "I couldn't use the attached screenshot because it was not a valid bounded JPEG, PNG, or WebP image.");
            }
            var activeModel = ModelRouting.ResolveForRequest(
                _settings,
                safeScreenshot.Length > 0);
            ContextScale.Apply(
                GeminiCodeAssistGateway.IsGeminiModel(
                    activeModel));
            var screenshotUsed =
                safeScreenshot.Length > 0 &&
                ModelCatalog.IsVisionCapable(activeModel);
            var activeTopic = ResolveTopic(topicId, topicBinding);
            var topicTools = activeTopic == null
                ? null
                : new TopicToolHost(
                    activeTopic,
                    chatId,
                    turnId,
                    true);

            try
            {

            // The unsent-draft and unsaved-workbook tools are
            // always exposed (owner's direction: never refuse an
            // action the user asked for). Their outputs stay safe
            // by construction - an unsent draft window and a new
            // unsaved workbook, each at most once per request.
            var allowOutlookDraft = !ExchangeContainsCall(
                exchange,
                BrowserToolCatalog.OpenOutlookDraft);
            var allowExcelTable = !ExchangeContainsCall(
                exchange,
                BrowserToolCatalog.OpenExcelTable);

            IReadOnlyList<ChatToolDefinition> definitions = null;
            if (_mcpTools.HasServers)
            {
                definitions = await Task.Run(
                    () => _mcpTools.GetDefinitions(),
                    cancellationToken).ConfigureAwait(false);
            }

            var request = BrowserChatRequestFactory.Create(
                activeModel,
                history ?? new ChatTurn[0],
                safePrompt,
                title,
                url,
                selection,
                pageText,
                safeScreenshot,
                definitions,
                exchange,
                links,
                activeTopic);

            var savedTask = BrowserTaskSession.Load(chatId, turnId) ?? new DurableTaskState {
                Id = BrowserTaskSession.Id(chatId, turnId), Host = "chrome", Objective = safePrompt };
            var pending = savedTask.PendingCalls.ToArray();
            var incoming = (exchange ?? new BrowserExchangeTurn[0]).SelectMany(t => t.Results ?? new List<BrowserExchangeResult>()).ToArray();
            foreach (var call in pending)
            {
                var result = incoming.FirstOrDefault(r => r.Id == call.id);
                if (result != null && !savedTask.PendingResults.Any(r => r.tool_call_id == call.id))
                    savedTask.PendingResults.Add(new ChatCompletionToolResultMessage { role = "tool", tool_call_id = call.id, content = result.Content });
            }
            var taskContext = new TaskContextManager(request, "chrome", safePrompt, resume: savedTask);
            if (pending.Length > 0)
            {
                var receipts = incoming.Where(r => pending.Any(c => c.id == r.Id)).Select(r => new MailboxToolResult(r.Id, r.Content, "Restored browser receipt")).ToList();
                taskContext.RecordExchange(request, new ChatCompletionResponseMessage { tool_calls = pending.ToList() }, receipts);
            }
            request.messages.Add(new ChatCompletionInputMessage { role = "user", content = "Current browser page (untrusted): " + url + "\n" + pageText + "\nControls must be rediscovered after resume. Use the original condition if supplied; otherwise ask which condition applies and offer Compare all. For Compare all, enumerate every condition and record separately verified quote evidence for each. Do not finish with only the final quote." });
            taskContext.SaveRequest(request);
            request.messages.Add(new ChatCompletionInputMessage { role = "user", content = BrowserTaskSession.CoverageNote(taskContext.State) });
            allowOutlookDraft = !taskContext.State.HostData.ContainsKey("browser_draft_spent");
            allowExcelTable = !taskContext.State.HostData.ContainsKey("browser_excel_spent");
            var draftOpened = false;
            var tableOpened = false;
            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var response = await taskContext.CompleteAsync(_client, _settings, request, null, cancellationToken).ConfigureAwait(false);
                var toolCalls = NormalizeCalls(response.tool_calls);
                BrowserTaskSession.ThrowIfPaused(chatId, turnId);
                if (toolCalls.Count == 0)
                {
                    if (string.IsNullOrWhiteSpace(response.content))
                    {
                        throw new AiEndpointException(
                            "RESPONSE_MISSING_CONTENT",
                            "I stopped because the model did not return text.");
                    }

                    taskContext.State.EnumerationComplete = true;
                    // Extension reconciles quote coverage before accepting completion.
                    taskContext.State.Lifecycle = TaskLifecycle.AwaitingUser;
                    taskContext.SaveRequest(request);
                    topicTools?.CompleteSession();
                    return new BrowserChatResult(
                        response.content,
                        activeModel,
                        screenshotUsed);
                }

                taskContext.PrepareExchange(response, request);
                if (PromptHelperTool.Contains(toolCalls) &&
                    toolCalls.Count != 1)
                {
                    var rejected = new List<MailboxToolResult>();
                    foreach (var rejectedCall in toolCalls)
                    {
                        rejected.Add(
                            PromptHelperTool.MixedCallResult(
                                rejectedCall));
                    }

                    ChatRequestFactory.AppendToolExchange(
                        request,
                        response,
                        rejected,
                        activeModel);
                    request.tool_choice =
                        PromptHelperTool.CreateRequiredChoice();
                    taskContext.FinishExchange(request);
                    continue;
                }

                var deferredMutationResults =
                    new List<MailboxToolResult>();
                var mutationCount = 0;
                foreach (var call in toolCalls)
                {
                    if (IsStateChangingBrowserCall(call))
                    {
                        mutationCount++;
                        if (mutationCount >
                            MaxStateChangingBrowserCallsPerRound)
                        {
                            deferredMutationResults.Add(
                                new MailboxToolResult(
                                    call.id,
                                    "[BROWSER_MUTATION_DEFERRED] I ran only one state-changing browser action from this model round. Inspect the fresh result before requesting another action.",
                                    "I deferred an extra browser action"));
                            continue;
                        }
                    }

                }

                var needsBrowser = false;
                foreach (var call in toolCalls)
                {
                    if (BrowserToolCatalog.IsBrowserExecuted(
                        call.function.name))
                    {
                        needsBrowser = true;
                        break;
                    }
                }

                var hostResults = new List<MailboxToolResult>(
                    deferredMutationResults);
                foreach (var call in toolCalls)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    BrowserTaskSession.ThrowIfPaused(chatId, turnId);
                    var name = call.function.name;
                    if (BrowserToolCatalog.IsBrowserExecuted(name))
                    {
                        continue;
                    }

                    if (name == TaskContextManager.ReadEvidenceTool) { hostResults.Add(taskContext.ReadEvidence(call)); continue; }
                    if (string.Equals(
                        name,
                        BrowserToolCatalog.OpenOutlookDraft,
                        StringComparison.Ordinal))
                    {
                        taskContext.State.HostData["browser_draft_spent"] = "true";
                        taskContext.BeforeTool(call, true);
                        hostResults.Add(ExecuteDraft(
                            call,
                            allowOutlookDraft && !draftOpened));
                        draftOpened = true;
                        continue;
                    }

                    if (string.Equals(
                        name,
                        BrowserToolCatalog.OpenExcelTable,
                        StringComparison.Ordinal))
                    {
                        taskContext.State.HostData["browser_excel_spent"] = "true";
                        taskContext.BeforeTool(call, true);
                        hostResults.Add(ExecuteExcelTable(
                            call,
                            allowExcelTable && !tableOpened));
                        tableOpened = true;
                        continue;
                    }

                    if (McpToolHost.IsMcpTool(name) &&
                        _mcpTools.HasServers)
                    {
                        hostResults.Add(await Task.Run(
                            () => _mcpTools.Execute(call),
                            cancellationToken)
                            .ConfigureAwait(false));
                        continue;
                    }

                    if (TopicToolCatalog.IsTopicTool(name) &&
                        topicTools != null)
                    {
                        hostResults.Add(await Task.Run(
                            () => topicTools.Execute(
                                call,
                                cancellationToken),
                            cancellationToken)
                            .ConfigureAwait(false));
                        continue;
                    }

                    hostResults.Add(new MailboxToolResult(
                        call.id,
                        "[BROWSER_TOOL_NOT_ALLOWED] This host does not expose the requested tool.",
                        "BROWSER_TOOL_NOT_ALLOWED"));
                }

                foreach (var result in hostResults)
                {
                    var call = toolCalls.First(c => c.id == result.ToolCallId);
                    taskContext.AfterTool(call, result);
                }
                if (needsBrowser)
                {
                    var pendingResults =
                        new List<BrowserExchangeResult>();
                    foreach (var result in hostResults)
                    {
                        pendingResults.Add(
                            new BrowserExchangeResult
                            {
                                Id = result.ToolCallId,
                                Content = result.Content
                            });
                    }

                    return new BrowserChatResult(
                        string.Empty,
                        activeModel,
                        screenshotUsed,
                        response.content,
                        toolCalls,
                        pendingResults);
                }

                taskContext.RecordExchange(request, response, hostResults);
                ChatRequestFactory.AppendToolExchange(
                    request,
                    response,
                    hostResults,
                    activeModel);
                taskContext.FinishExchange(request);
            }
            }
            catch
            {
                topicTools?.CompleteSession();
                throw;
            }
        }

        public static string TopicBinding(TopicConfig topic)
        {
            if (topic == null)
            {
                return string.Empty;
            }

            using (var sha = SHA256.Create())
            {
                var bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(
                    (topic.Id ?? string.Empty).ToUpperInvariant() +
                    "\n" +
                    (topic.FolderPath ?? string.Empty)
                        .ToUpperInvariant()));
                return Convert.ToBase64String(bytes)
                    .TrimEnd('=')
                    .Replace('+', '-')
                    .Replace('/', '_');
            }
        }

        private TopicConfig ResolveTopic(
            string topicId,
            string binding)
        {
            var bounded = TextBoundary.SingleLine(topicId, 40);
            if (bounded.Length == 0)
            {
                return null;
            }

            var topic = _settings.Topics.Find(entry =>
                string.Equals(
                    entry.Id,
                    bounded,
                    StringComparison.OrdinalIgnoreCase));
            if (topic == null || !string.Equals(
                    TopicBinding(topic),
                    TextBoundary.SingleLine(binding, 100),
                    StringComparison.Ordinal))
            {
                throw new AiEndpointException(
                    "TOPIC_CHANGED",
                    "I can't continue because the active Topic was removed or its folder changed. Clear chat first.");
            }

            string resolvedRoot;
            string validationError;
            if (!TopicConfig.TryValidateLocalFolder(
                    topic.FolderPath,
                    out resolvedRoot,
                    out validationError))
            {
                throw new AiEndpointException(
                    "TOPIC_UNAVAILABLE",
                    "I can't use the active Topic: " +
                    validationError);
            }

            return topic;
        }

        private static List<ChatToolCall> NormalizeCalls(
            IReadOnlyList<ChatToolCall> toolCalls)
        {
            var result = new List<ChatToolCall>();
            if (toolCalls == null)
            {
                return result;
            }

            foreach (var call in toolCalls)
            {
                if (call?.function == null ||
                    string.IsNullOrWhiteSpace(call.id) ||
                    string.IsNullOrWhiteSpace(call.function.name))
                {
                    continue;
                }

                result.Add(call);
            }

            return result;
        }

        private static int CountExchangeTurns(
            IReadOnlyList<BrowserExchangeTurn> exchange)
        {
            var count = 0;
            foreach (var turn in exchange ??
                new BrowserExchangeTurn[0])
            {
                if (turn?.ToolCalls != null &&
                    turn.ToolCalls.Count > 0)
                {
                    count++;
                }
            }

            return count;
        }

        private static bool IsStateChangingBrowserCall(
            ChatToolCall call)
        {
            var name = call?.function?.name ?? string.Empty;
            if (string.Equals(
                    name,
                    BrowserToolCatalog.NavigatePage,
                    StringComparison.Ordinal) ||
                string.Equals(
                    name,
                    BrowserToolCatalog.SearchGoogle,
                    StringComparison.Ordinal))
            {
                return true;
            }

            if (!string.Equals(
                    name,
                    BrowserToolCatalog.ActOnPage,
                    StringComparison.Ordinal))
            {
                return false;
            }

            try
            {
                var arguments = new JavaScriptSerializer()
                    .DeserializeObject(
                        call.function.arguments ?? "{}")
                    as IDictionary<string, object>;
                object action;
                if (arguments != null &&
                    arguments.TryGetValue("action", out action) &&
                    string.Equals(
                        Convert.ToString(action),
                        "wait",
                        StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }
            catch (ArgumentException)
            {
                // Invalid arguments remain state-changing until the
                // extension rejects them; never use malformed JSON to
                // bypass the per-round mutation boundary.
            }

            return true;
        }

        public static bool HasStalled(
            IReadOnlyList<BrowserExchangeTurn> exchange)
        {
            var stagnantCalls = 0;
            foreach (var turn in exchange ??
                new BrowserExchangeTurn[0])
            {
                foreach (var call in turn?.ToolCalls ??
                    new List<ChatToolCall>())
                {
                    if (!BrowserToolCatalog.IsBrowserExecuted(
                            call?.function?.name))
                    {
                        continue;
                    }

                    var content = string.Empty;
                    foreach (var result in turn.Results ??
                        new List<BrowserExchangeResult>())
                    {
                        if (string.Equals(
                            result?.Id,
                            call.id,
                            StringComparison.Ordinal))
                        {
                            content = result.Content ?? string.Empty;
                            break;
                        }
                    }

                    if (content.IndexOf(
                            "Progress marker: changed",
                            StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        stagnantCalls = 0;
                    }
                    else if (content.IndexOf(
                                "Progress marker: unchanged",
                                StringComparison.OrdinalIgnoreCase) >= 0 ||
                             content.StartsWith(
                                "[BROWSER_TOOL_FAILED]",
                                StringComparison.Ordinal))
                    {
                        stagnantCalls++;
                    }

                    if (stagnantCalls >= MaxBrowserStagnantCalls)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        public static bool IsSupportOnlyRound(
            IReadOnlyList<ChatToolCall> calls)
        {
            if (calls == null || calls.Count == 0)
            {
                return false;
            }

            foreach (var call in calls)
            {
                if (!string.Equals(
                        call?.function?.name,
                        BrowserToolCatalog.ActOnPage,
                        StringComparison.Ordinal))
                {
                    return false;
                }

                IDictionary<string, object> arguments;
                try
                {
                    arguments = new JavaScriptSerializer()
                        .DeserializeObject(
                            call.function.arguments ?? "{}")
                        as IDictionary<string, object>;
                }
                catch (ArgumentException)
                {
                    return false;
                }

                object actionValue;
                if (arguments == null ||
                    !arguments.TryGetValue("action", out actionValue))
                {
                    return false;
                }

                var action = Convert.ToString(actionValue);
                if (!string.Equals(action, "scroll", StringComparison.Ordinal) &&
                    !string.Equals(action, "wait", StringComparison.Ordinal))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool ExchangeContainsCall(
            IReadOnlyList<BrowserExchangeTurn> exchange,
            string toolName)
        {
            foreach (var turn in exchange ??
                new BrowserExchangeTurn[0])
            {
                foreach (var call in turn?.ToolCalls ??
                    new List<ChatToolCall>())
                {
                    if (string.Equals(
                        call?.function?.name,
                        toolName,
                        StringComparison.Ordinal))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static MailboxToolResult ExecuteExcelTable(
            ChatToolCall call,
            bool allowed)
        {
            if (!allowed)
            {
                return new MailboxToolResult(
                    call.id,
                    "[EXCEL_TABLE_NOT_AUTHORIZED] One unsaved workbook was already opened for this request; tell the user instead of opening another.",
                    "EXCEL_TABLE_NOT_AUTHORIZED");
            }

            try
            {
                var serializer =
                    new System.Web.Script.Serialization
                        .JavaScriptSerializer();
                var arguments =
                    serializer.DeserializeObject(
                        call.function.arguments ?? "{}") as
                        IDictionary<string, object> ??
                    new Dictionary<string, object>();
                var status = Office.ExcelTableLauncher.OpenTable(
                    Argument(arguments, "title"),
                    StringListArgument(arguments, "columns"),
                    RowsArgument(arguments, "rows"),
                    Argument(arguments, "chart_kind"),
                    Argument(arguments, "chart_title"));
                return new MailboxToolResult(
                    call.id,
                    status,
                    "Unsaved Excel workbook opened for review.");
            }
            catch (Exception exception)
            {
                return new MailboxToolResult(
                    call.id,
                    "[EXCEL_TABLE_FAILED] " + TextBoundary.PlainText(
                        exception.Message,
                        600),
                    "EXCEL_TABLE_FAILED");
            }
        }

        private static List<string> StringListArgument(
            IDictionary<string, object> arguments,
            string key)
        {
            var result = new List<string>();
            object value;
            var array = arguments.TryGetValue(key, out value)
                ? value as object[]
                : null;
            foreach (var item in array ?? new object[0])
            {
                result.Add(Convert.ToString(item));
            }

            return result;
        }

        private static List<IReadOnlyList<string>> RowsArgument(
            IDictionary<string, object> arguments,
            string key)
        {
            var result = new List<IReadOnlyList<string>>();
            object value;
            var array = arguments.TryGetValue(key, out value)
                ? value as object[]
                : null;
            foreach (var row in array ?? new object[0])
            {
                var cells = new List<string>();
                foreach (var cell in row as object[] ??
                    new object[0])
                {
                    cells.Add(Convert.ToString(cell));
                }

                result.Add(cells);
            }

            return result;
        }

        private MailboxToolResult ExecuteDraft(
            ChatToolCall call,
            bool allowed)
        {
            if (!allowed)
            {
                return new MailboxToolResult(
                    call.id,
                    "[DRAFT_NOT_AUTHORIZED] One unsent draft was already opened for this request; tell the user instead of opening another.",
                    "DRAFT_NOT_AUTHORIZED");
            }

            try
            {
                var serializer =
                    new System.Web.Script.Serialization
                        .JavaScriptSerializer();
                var arguments =
                    serializer.DeserializeObject(
                        call.function.arguments ?? "{}") as
                        IDictionary<string, object> ??
                    new Dictionary<string, object>();
                var status = OutlookDraftLauncher.OpenDraft(
                    Argument(arguments, "to"),
                    Argument(arguments, "cc"),
                    Argument(arguments, "subject"),
                    Argument(arguments, "body"));
                return new MailboxToolResult(
                    call.id,
                    status,
                    "Unsent Outlook draft opened for review.");
            }
            catch (Exception exception)
            {
                return new MailboxToolResult(
                    call.id,
                    "[DRAFT_FAILED] " + TextBoundary.PlainText(
                        exception.Message,
                        600),
                    "DRAFT_FAILED");
            }
        }

        private static string Argument(
            IDictionary<string, object> arguments,
            string key)
        {
            object value;
            return arguments.TryGetValue(key, out value)
                ? Convert.ToString(value)
                : string.Empty;
        }

        public void Dispose()
        {
            _mcpTools.Dispose();
            _client.Dispose();
        }
    }
}

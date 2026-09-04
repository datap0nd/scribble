using System;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using Scribble.Security;

namespace Scribble.Chat
{
    public sealed class PromptHelperOption
    {
        public PromptHelperOption(
            string label,
            string description)
        {
            Label = label ?? string.Empty;
            Description = description ?? string.Empty;
        }

        public string Label { get; }

        public string Description { get; }
    }

    public sealed class PromptHelperQuestion
    {
        public PromptHelperQuestion(
            string question,
            string reason,
            IReadOnlyList<PromptHelperOption> options)
            : this("answer", question, reason, options)
        {
        }

        public PromptHelperQuestion(
            string id,
            string question,
            string reason,
            IReadOnlyList<PromptHelperOption> options)
        {
            Id = id ?? string.Empty;
            Question = question ?? string.Empty;
            Reason = reason ?? string.Empty;
            Options = options ?? new PromptHelperOption[0];
        }

        public string Id { get; }

        public string Question { get; }

        public string Reason { get; }

        public IReadOnlyList<PromptHelperOption> Options { get; }
    }

    // A suite-wide, read-only human-in-the-loop tool. It is always
    // available, because clarification has to happen before a model
    // chooses an app-specific read or draft tool. A small deterministic
    // preflight catches only the clearest vague prompts; the model handles
    // contextual ambiguity using the richer instruction below.
    public static class PromptHelperTool
    {
        public const string Name = "ask_user";
        public const int MaxQuestionCharacters = 300;
        public const int MaxReasonCharacters = 180;
        public const int MaxOptions = 4;
        public const int MaxOptionLabelCharacters = 80;
        public const int MaxOptionDescriptionCharacters = 140;
        public const int MaxAnswerCharacters = 240;
        public const int MaxQuestions = 3;

        public const string SystemInstruction =
            " Before starting substantial work, check whether the request " +
            "is clear enough to produce the result the user probably wants. " +
            "You MUST call ask_user, as the only tool call in that response, " +
            "when a missing detail could materially change the work: the " +
            "goal or deliverable, intended audience, source or scope, success " +
            "criteria, important constraints, output destination or format, " +
            "or a key person, place, date, budget, quantity, or product " +
            "variant. This especially applies to vague prompts such as " +
            "'make it better', 'research laptops', 'create a presentation', " +
            "or 'email them'. Ask one focused plain-language question at a " +
            "time, explain briefly why it matters, and provide 2-4 concrete " +
            "options with the most likely or recommended option first; the " +
            "user can always type a different answer. Do not ask when the " +
            "answer is already explicit in the request or supplied context, " +
            "when a safe obvious default will not materially change the " +
            "outcome, or for greetings, simple factual questions, and trivial " +
            "choices. ask_user gathers missing requirements; it is never a " +
            "request for permission to proceed with an already-clear task.";

        public const string BrowserSystemInstruction =
            " Before browsing, check whether missing details could materially " +
            "change the result. When several related product or research " +
            "details are missing, call ask_user as the only tool call with " +
            "one to three focused questions in a single questions array. " +
            "Give each question a stable short id and 2-4 concrete options. " +
            "After the answers arrive, continue the same browser journey. " +
            "Ask again only for a field the user left unanswered. Do not use " +
            "ask_user as permission for an already-clear task.";

        private static readonly Regex Whitespace =
            new Regex(@"\s+", RegexOptions.Compiled);
        private static readonly Regex Word =
            new Regex(@"[\p{L}\p{N}]+", RegexOptions.Compiled);
        private static readonly Regex VagueWholeRequest =
            new Regex(
                @"^(please\s+)?(help(\s+me)?|do\s+(it|this)|handle\s+this|" +
                @"sort\s+this\s+out|make\s+(it|this)\s+(better|good|nice)|" +
                @"improve\s+(it|this)|fix\s+(it|this)|clean\s+this\s+up|" +
                @"review\s+this|analy[sz]e\s+this|look\s+at\s+this|" +
                @"what\s+do\s+you\s+think|create\s+something|" +
                @"make\s+something|write\s+something|" +
                @"(send|email)\s+(it|this|that|them))[.!?]*$",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex ContextDependentRequest =
            new Regex(
                @"\b(summarize|explain|rewrite|review|analy[sz]e|fix|" +
                @"update|compare|send|email)\s+(it|this|that|them)\b",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex BroadOutputRequest =
            new Regex(
                @"\b(create|make|build|write|draft|prepare|generate)\b.*" +
                @"\b(presentation|powerpoint|deck|spreadsheet|workbook|" +
                @"report|document|email|message)\b",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex OutputDetailMarker =
            new Regex(
                @"\b(about|on|for|to|from|using|with|regarding|covering|" +
                @"based\s+on)\b",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex ShortDiscoveryRequest =
            new Regex(
                @"^(research|compare|find|recommend|choose|plan)\b",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex ConversationalReply =
            new Regex(
                @"^(hi|hello|hey|good\s+(morning|afternoon|evening)|yes|no|" +
                @"ok|okay|sure|thanks|thank\s+you|continue|go\s+ahead|" +
                @"sounds\s+good)[.!?]*$",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static ChatToolDefinition CreateDefinition()
        {
            return new ChatToolDefinition
            {
                type = "function",
                function = new ChatToolFunctionDefinition
                {
                    name = Name,
                    description =
                        "Pause and ask the user one focused question when a " +
                        "missing requirement would materially change the result. " +
                        "Use this before substantial work for an unclear goal, " +
                        "audience, source, scope, success criterion, constraint, " +
                        "output format, recipient, location, date, budget, quantity, " +
                        "or variant. This must be the only tool call in the response. " +
                        "Do not use it for trivial choices or when supplied context " +
                        "or a safe obvious default already answers the question.",
                    parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "additionalProperties", false },
                        {
                            "properties",
                            new Dictionary<string, object>
                            {
                                {
                                    "question",
                                    new Dictionary<string, object>
                                    {
                                        { "type", "string" },
                                        {
                                            "description",
                                            "One short, plain-language question."
                                        }
                                    }
                                },
                                {
                                    "reason",
                                    new Dictionary<string, object>
                                    {
                                        { "type", "string" },
                                        {
                                            "description",
                                            "One brief sentence explaining how the answer changes the result."
                                        }
                                    }
                                },
                                {
                                    "options",
                                    new Dictionary<string, object>
                                    {
                                        { "type", "array" },
                                        { "minItems", 2 },
                                        { "maxItems", MaxOptions },
                                        {
                                            "items",
                                            new Dictionary<string, object>
                                            {
                                                { "type", "object" },
                                                { "additionalProperties", false },
                                                {
                                                    "properties",
                                                    new Dictionary<string, object>
                                                    {
                                                        {
                                                            "label",
                                                            new Dictionary<string, object>
                                                            {
                                                                { "type", "string" },
                                                                { "description", "A concise choice, ideally 1-5 words." }
                                                            }
                                                        },
                                                        {
                                                            "description",
                                                            new Dictionary<string, object>
                                                            {
                                                                { "type", "string" },
                                                                { "description", "A short explanation of the impact or trade-off." }
                                                            }
                                                        }
                                                    }
                                                },
                                                { "required", new[] { "label", "description" } }
                                            }
                                        }
                                    }
                                }
                            }
                        },
                        { "required", new[] { "question", "options" } }
                    }
                }
            };
        }

        public static ChatToolDefinition CreateBrowserDefinition()
        {
            return new ChatToolDefinition
            {
                type = "function",
                function = new ChatToolFunctionDefinition
                {
                    name = Name,
                    description =
                        "Ask one to three related questions together when missing " +
                        "research details would materially change the result. This " +
                        "must be the only tool call in the response.",
                    parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "additionalProperties", false },
                        {
                            "properties",
                            new Dictionary<string, object>
                            {
                                {
                                    "questions",
                                    new Dictionary<string, object>
                                    {
                                        { "type", "array" },
                                        { "minItems", 1 },
                                        { "maxItems", MaxQuestions },
                                        {
                                            "items",
                                            new Dictionary<string, object>
                                            {
                                                { "type", "object" },
                                                { "additionalProperties", false },
                                                {
                                                    "properties",
                                                    new Dictionary<string, object>
                                                    {
                                                        { "id", new Dictionary<string, object> { { "type", "string" } } },
                                                        { "question", new Dictionary<string, object> { { "type", "string" } } },
                                                        { "reason", new Dictionary<string, object> { { "type", "string" } } },
                                                        {
                                                            "options",
                                                            new Dictionary<string, object>
                                                            {
                                                                { "type", "array" },
                                                                { "minItems", 2 },
                                                                { "maxItems", MaxOptions },
                                                                {
                                                                    "items",
                                                                    new Dictionary<string, object>
                                                                    {
                                                                        { "type", "object" },
                                                                        { "additionalProperties", false },
                                                                        {
                                                                            "properties",
                                                                            new Dictionary<string, object>
                                                                            {
                                                                                { "label", new Dictionary<string, object> { { "type", "string" } } },
                                                                                { "description", new Dictionary<string, object> { { "type", "string" } } }
                                                                            }
                                                                        },
                                                                        { "required", new[] { "label", "description" } }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                },
                                                { "required", new[] { "id", "question", "options" } }
                                            }
                                        }
                                    }
                                }
                            }
                        },
                        { "required", new[] { "questions" } }
                    }
                }
            };
        }

        public static IReadOnlyList<PromptHelperQuestion> ParseMany(
            ChatToolCall call)
        {
            if (!IsTool(call?.function?.name))
            {
                throw new InvalidOperationException(
                    "The requested tool is not the prompt helper.");
            }

            IDictionary<string, object> arguments;
            try
            {
                arguments = new JavaScriptSerializer()
                    .DeserializeObject(call.function.arguments ?? "{}")
                    as IDictionary<string, object>;
            }
            catch (ArgumentException)
            {
                arguments = null;
            }

            object questionsValue = null;
            var questionItems = arguments != null &&
                arguments.TryGetValue("questions", out questionsValue)
                    ? questionsValue as IEnumerable
                    : null;
            if (questionItems == null || questionsValue is string)
            {
                return new[] { Parse(call) };
            }

            var result = new List<PromptHelperQuestion>();
            foreach (var item in questionItems)
            {
                var map = item as IDictionary<string, object>;
                if (map == null)
                {
                    continue;
                }
                object idValue;
                map.TryGetValue("id", out idValue);
                var parsed = Parse(new ChatToolCall
                {
                    id = call.id,
                    type = call.type,
                    function = new ChatToolCallFunction
                    {
                        name = Name,
                        arguments = new JavaScriptSerializer().Serialize(map)
                    }
                });
                result.Add(new PromptHelperQuestion(
                    TextBoundary.SingleLine(
                        Convert.ToString(idValue) ?? string.Empty,
                        40),
                    parsed.Question,
                    parsed.Reason,
                    parsed.Options));
                if (result.Count == MaxQuestions)
                {
                    break;
                }
            }

            if (result.Count == 0)
            {
                throw new InvalidOperationException(
                    "The prompt helper did not provide any questions.");
            }
            return result;
        }

        public static object CreateRequiredChoice()
        {
            return new Dictionary<string, object>
            {
                { "type", "function" },
                {
                    "function",
                    new Dictionary<string, object>
                    {
                        { "name", Name }
                    }
                }
            };
        }

        public static bool IsTool(string name)
        {
            return string.Equals(name, Name, StringComparison.Ordinal);
        }

        public static bool Contains(
            IReadOnlyList<ChatToolCall> calls)
        {
            foreach (var call in calls ?? new ChatToolCall[0])
            {
                if (IsTool(call?.function?.name))
                {
                    return true;
                }
            }

            return false;
        }

        public static bool ShouldRequireClarification(
            string prompt,
            bool hasRelevantContext)
        {
            var normalized = Whitespace.Replace(
                    TextBoundary.PlainText(
                        prompt,
                        TextBoundary.MaxUserPromptCharacters),
                    " ")
                .Trim();
            if (normalized.Length == 0 ||
                ConversationalReply.IsMatch(normalized))
            {
                return false;
            }

            var wordCount = Word.Matches(normalized).Count;
            if (wordCount <= 1 ||
                VagueWholeRequest.IsMatch(normalized))
            {
                return true;
            }

            if (!hasRelevantContext &&
                wordCount <= 8 &&
                ContextDependentRequest.IsMatch(normalized))
            {
                return true;
            }

            if (wordCount <= 7 &&
                BroadOutputRequest.IsMatch(normalized) &&
                (wordCount <= 5 ||
                 !OutputDetailMarker.IsMatch(normalized)))
            {
                return true;
            }

            return wordCount <= 3 &&
                ShortDiscoveryRequest.IsMatch(normalized);
        }

        public static PromptHelperQuestion Parse(
            ChatToolCall call)
        {
            if (!IsTool(call?.function?.name))
            {
                throw new InvalidOperationException(
                    "The requested tool is not the prompt helper.");
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
                arguments = null;
            }

            if (arguments == null)
            {
                throw new InvalidOperationException(
                    "The prompt-helper arguments were not valid JSON.");
            }

            object questionValue;
            object reasonValue;
            object optionsValue;
            arguments.TryGetValue("question", out questionValue);
            arguments.TryGetValue("reason", out reasonValue);
            arguments.TryGetValue("options", out optionsValue);
            var question = TextBoundary.SingleLine(
                Convert.ToString(questionValue) ?? string.Empty,
                MaxQuestionCharacters);
            if (question.Length == 0)
            {
                throw new InvalidOperationException(
                    "The prompt helper did not provide a question.");
            }

            var options = new List<PromptHelperOption>();
            var optionItems = optionsValue as IEnumerable;
            if (optionItems != null && !(optionsValue is string))
            {
                foreach (var item in optionItems)
                {
                    var option = ParseOption(item);
                    if (option != null &&
                        !ContainsLabel(options, option.Label))
                    {
                        options.Add(option);
                    }

                    if (options.Count == MaxOptions)
                    {
                        break;
                    }
                }
            }

            return new PromptHelperQuestion(
                question,
                TextBoundary.SingleLine(
                    Convert.ToString(reasonValue) ?? string.Empty,
                    MaxReasonCharacters),
                options);
        }

        public static MailboxToolResult MixedCallResult(
            ChatToolCall call)
        {
            return new MailboxToolResult(
                call?.id ?? string.Empty,
                "[ASK_USER_MUST_BE_ALONE] No requested tools ran. Call " +
                "ask_user by itself, wait for the answer, and then continue.",
                "Waiting for a focused clarification");
        }

        private static PromptHelperOption ParseOption(object value)
        {
            var map = value as IDictionary<string, object>;
            string label;
            string description;
            if (map != null)
            {
                object labelValue;
                object descriptionValue;
                map.TryGetValue("label", out labelValue);
                map.TryGetValue("description", out descriptionValue);
                label = TextBoundary.SingleLine(
                    Convert.ToString(labelValue) ?? string.Empty,
                    MaxOptionLabelCharacters);
                description = TextBoundary.SingleLine(
                    Convert.ToString(descriptionValue) ?? string.Empty,
                    MaxOptionDescriptionCharacters);
            }
            else
            {
                // Accept the original browser tool's string options so
                // older/local model behavior remains compatible.
                label = TextBoundary.SingleLine(
                    Convert.ToString(value) ?? string.Empty,
                    MaxOptionLabelCharacters);
                description = string.Empty;
            }

            return label.Length == 0
                ? null
                : new PromptHelperOption(label, description);
        }

        private static bool ContainsLabel(
            IReadOnlyList<PromptHelperOption> options,
            string label)
        {
            foreach (var option in options)
            {
                if (string.Equals(
                    option.Label,
                    label,
                    StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }
    }
}

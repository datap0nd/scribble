using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;
using Scribble.Chat;
using Scribble.Security;
using Scribble.Utilities;

namespace Scribble.Outlook
{
    public sealed class DraftToolHost : IDisposable
    {
        private static readonly HashSet<string> CreateArguments =
            new HashSet<string>(StringComparer.Ordinal)
            {
                "kind",
                "body",
                "subject",
                "to",
                "cc",
                "reply_handle",
                "bold_phrases"
            };

        private static readonly HashSet<string> UpdateArguments =
            new HashSet<string>(StringComparer.Ordinal)
            {
                "body",
                "subject",
                "to",
                "cc",
                "bold_phrases"
            };

        private readonly object _outlookApplication;
        private readonly JavaScriptSerializer _serializer =
            new JavaScriptSerializer();
        private DraftSession _session;

        public DraftToolHost(object outlookApplication)
        {
            _outlookApplication = outlookApplication ??
                throw new ArgumentNullException(nameof(outlookApplication));
        }

        public bool HasActiveDraft
        {
            get { return _session != null; }
        }

        public DraftReference ActiveDraft
        {
            get { return _session?.Reference; }
        }

        public MailboxToolResult Execute(
            ChatToolCall call,
            Func<string, MessageSnapshot> resolveMessage,
            OneShotDraftAuthorization authorization,
            bool isOnlyToolCall)
        {
            var name = call?.function?.name;
            if (call?.function == null ||
                string.IsNullOrWhiteSpace(call.id) ||
                !DraftToolCatalog.IsDraftTool(name))
            {
                return Error(
                    call?.id,
                    authorization,
                    "DRAFT_TOOL_CALL_INVALID",
                    "The model returned an invalid draft tool call.");
            }

            if (!isOnlyToolCall)
            {
                return Error(
                    call.id,
                    authorization,
                    "DRAFT_TOOL_MUST_BE_EXCLUSIVE",
                    name + " must be the only tool call in its response.");
            }

            IDictionary<string, object> arguments;
            try
            {
                arguments = ParseArguments(
                    call.function.arguments);
                RequireAllowedArguments(
                    arguments,
                    DraftToolCatalog.IsCreateDraft(name)
                        ? CreateArguments
                        : UpdateArguments);
            }
            catch (Exception exception)
            {
                return Error(
                    call.id,
                    authorization,
                    "DRAFT_ARGUMENTS_INVALID",
                    DiagnosticDetails.ForException(
                        exception,
                        "DRAFT_ARGUMENTS_INVALID"));
            }

            var body = TextBoundary.PlainText(
                GetString(arguments, "body"),
                TextBoundary.MaxAssistantCharacters);
            if (body.Length == 0)
            {
                return Error(
                    call.id,
                    authorization,
                    "DRAFT_BODY_REQUIRED",
                    "A non-empty draft body is required.");
            }

            return DraftToolCatalog.IsCreateDraft(name)
                ? Create(
                    call.id,
                    arguments,
                    body,
                    resolveMessage,
                    authorization)
                : Update(
                    call.id,
                    arguments,
                    body,
                    authorization);
        }

        public void Dispose()
        {
            _session?.Dispose();
            _session = null;
        }

        private MailboxToolResult Create(
            string callId,
            IDictionary<string, object> arguments,
            string body,
            Func<string, MessageSnapshot> resolveMessage,
            OneShotDraftAuthorization authorization)
        {
            if (_session != null)
            {
                return Error(
                    callId,
                    authorization,
                    "DRAFT_ALREADY_LINKED",
                    "This chat already has one linked Outlook draft.");
            }

            var kind = GetString(arguments, "kind")
                .ToLowerInvariant();
            if (kind != "new" && kind != "reply")
            {
                return Error(
                    callId,
                    authorization,
                    "DRAFT_KIND_INVALID",
                    "Draft kind must be new or reply.");
            }

            var replyHandle = TextBoundary.SingleLine(
                GetString(arguments, "reply_handle"),
                64);
            if (kind == "new" && replyHandle.Length > 0)
            {
                return Error(
                    callId,
                    authorization,
                    "DRAFT_REPLY_HANDLE_NOT_ALLOWED",
                    "A new-message draft cannot include a reply handle.");
            }

            MessageSnapshot replySource = null;
            if (kind == "reply" && replyHandle.Length == 0)
            {
                return Error(
                    callId,
                    authorization,
                    "DRAFT_REPLY_HANDLE_REQUIRED",
                    "A reply draft requires the exact handle returned for the source email.");
            }

            if (kind == "reply")
            {
                replySource = resolveMessage?.Invoke(replyHandle);
                if (replySource == null || !replySource.CanReply)
                {
                    return Error(
                        callId,
                        authorization,
                        "DRAFT_REPLY_HANDLE_UNKNOWN",
                        "The reply handle is unknown, expired, or not reply-capable.");
                }
            }

            if (authorization == null ||
                !authorization.CanCreate ||
                !authorization.TryConsume())
            {
                return Error(
                    callId,
                    authorization,
                    "DRAFT_PERMISSION_NOT_AVAILABLE",
                    "No unused local draft-creation permission is available for this request.");
            }

            try
            {
                var drafts = new DraftService(
                    _outlookApplication);
                var boldPhrases = GetStringList(
                    arguments,
                    "bold_phrases");
                _session = kind == "reply"
                    ? drafts.CreateReplyDraft(
                        replySource,
                        body,
                        boldPhrases)
                    : drafts.CreateNewDraft(
                        body,
                        boldPhrases,
                        GetOptionalString(arguments, "subject"),
                        GetOptionalString(arguments, "to"),
                        GetOptionalString(arguments, "cc"));
                authorization.MarkCreated();
                return Success(
                    callId,
                    "created",
                    kind);
            }
            catch (Exception exception)
            {
                Log.Error("DraftTool.CreateDraft", exception);
                return Error(
                    callId,
                    authorization,
                    "DRAFT_CREATION_FAILED",
                    DiagnosticDetails.ForException(
                        exception,
                        "DRAFT_CREATION_FAILED"));
            }
        }

        private MailboxToolResult Update(
            string callId,
            IDictionary<string, object> arguments,
            string body,
            OneShotDraftAuthorization authorization)
        {
            if (_session == null)
            {
                return Error(
                    callId,
                    authorization,
                    "DRAFT_NOT_LINKED",
                    "There is no linked Outlook draft to update.");
            }

            if (authorization == null ||
                !authorization.CanUpdate ||
                !authorization.TryConsume())
            {
                return Error(
                    callId,
                    authorization,
                    "DRAFT_UPDATE_NOT_AVAILABLE",
                    "No unused local draft-update permission is available for this request.");
            }

            try
            {
                _session.Update(
                    body,
                    GetStringList(arguments, "bold_phrases"),
                    GetOptionalString(arguments, "subject"),
                    GetOptionalString(arguments, "to"),
                    GetOptionalString(arguments, "cc"));
                authorization.MarkUpdated();
                return Success(
                    callId,
                    "updated",
                    _session.Reference.Kind);
            }
            catch (Exception exception)
            {
                Log.Error("DraftTool.UpdateDraft", exception);
                return Error(
                    callId,
                    authorization,
                    "DRAFT_UPDATE_FAILED",
                    DiagnosticDetails.ForException(
                        exception,
                        "DRAFT_UPDATE_FAILED"));
            }
        }

        private IDictionary<string, object> ParseArguments(
            string json)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                return new Dictionary<string, object>();
            }

            var parsed = _serializer.DeserializeObject(json) as
                IDictionary<string, object>;
            if (parsed == null)
            {
                throw new InvalidOperationException(
                    "Draft tool arguments must be a JSON object.");
            }

            return parsed;
        }

        private static void RequireAllowedArguments(
            IDictionary<string, object> arguments,
            ISet<string> allowed)
        {
            var unexpected = arguments.Keys
                .FirstOrDefault(key => !allowed.Contains(key));
            if (unexpected != null)
            {
                throw new InvalidOperationException(
                    "Unexpected draft argument: " + unexpected);
            }
        }

        private static string GetString(
            IDictionary<string, object> arguments,
            string name)
        {
            object value;
            return arguments != null &&
                   arguments.TryGetValue(name, out value)
                ? Convert.ToString(value) ?? string.Empty
                : string.Empty;
        }

        private static string GetOptionalString(
            IDictionary<string, object> arguments,
            string name)
        {
            return arguments.ContainsKey(name)
                ? GetString(arguments, name)
                : null;
        }

        private static IReadOnlyList<string> GetStringList(
            IDictionary<string, object> arguments,
            string name)
        {
            object value;
            if (!arguments.TryGetValue(name, out value) ||
                value == null)
            {
                return new string[0];
            }

            var values = value as IEnumerable;
            if (values == null || value is string)
            {
                throw new InvalidOperationException(
                    name + " must be an array of strings.");
            }

            var result = new List<string>();
            foreach (var item in values)
            {
                if (!(item is string))
                {
                    throw new InvalidOperationException(
                        name + " must contain only strings.");
                }

                result.Add(TextBoundary.PlainText(
                    (string)item,
                    200));
                if (result.Count > 12)
                {
                    throw new InvalidOperationException(
                        name + " cannot contain more than 12 phrases.");
                }
            }

            return result;
        }

        private MailboxToolResult Success(
            string callId,
            string action,
            string kind)
        {
            return new MailboxToolResult(
                callId,
                _serializer.Serialize(
                    new Dictionary<string, object>
                    {
                        { "ok", true },
                        { "action", action },
                        { "draft_kind", kind },
                        { "sent", false },
                        { "linked_draft_count", 1 }
                    }),
                action == "created"
                    ? "One unsent draft was opened and linked to this chat."
                    : "The linked unsent draft was updated in Outlook.");
        }

        private MailboxToolResult Error(
            string callId,
            OneShotDraftAuthorization authorization,
            string code,
            string message)
        {
            return new MailboxToolResult(
                callId ?? string.Empty,
                _serializer.Serialize(
                    new Dictionary<string, object>
                    {
                        { "ok", false },
                        { "error_code", code },
                        {
                            "message",
                            TextBoundary.PlainText(message, 1200)
                        },
                        {
                            "permission_consumed",
                            authorization != null &&
                            authorization.IsConsumed
                        }
                    }),
                "[" + code + "] " +
                TextBoundary.PlainText(message, 600));
        }
    }
}

using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Web.Script.Serialization;
using Scribble.Chat;
using Scribble.Security;
using Scribble.Utilities;

namespace Scribble.Outlook
{
    public sealed class MailboxToolHost
    {
        private const int MaxDirectMessageBodyCharacters = 6000;
        private const int MaxThreadMessageBodyCharacters = 2000;
        // Per-message allowance for the result-size gate, so a
        // user-raised working set can actually be returned. At the
        // default working set of ten this equals the standard
        // MaxToolResultCharacters budget.
        private const int PerMessageResultCharacters = 12000;
        private static int MaxThreadMessages
        {
            get { return MailboxWorkingSet.MaxMessages; }
        }

        private readonly MailboxContextService _mailbox;
        private readonly JavaScriptSerializer _serializer =
            new JavaScriptSerializer();
        private readonly Dictionary<string, MessageSnapshot> _handles =
            new Dictionary<string, MessageSnapshot>(
                StringComparer.Ordinal);
        private readonly HashSet<string> _loadedBodyHandles =
            new HashSet<string>(StringComparer.Ordinal);
        private readonly bool _workingSetOnly;
        private bool _searchExecuted;
        private int _nextHandle = 1;

        public MailboxToolHost(
            object outlookApplication,
            MessageSnapshot selectedMessage)
            : this(
                outlookApplication,
                selectedMessage,
                null)
        {
        }

        public MailboxToolHost(
            object outlookApplication,
            MessageSnapshot selectedMessage,
            IReadOnlyList<MessageSnapshot> workingMessages)
        {
            _mailbox = new MailboxContextService(outlookApplication);
            var workingSet = MailboxWorkingSet.Normalize(
                workingMessages);
            _workingSetOnly = workingSet.Count > 0;
            if (_workingSetOnly)
            {
                for (var index = 0;
                     index < workingSet.Count;
                     index++)
                {
                    _handles[MailboxWorkingSet.HandleAt(index)] =
                        workingSet[index];
                }
            }
            else if (selectedMessage != null)
            {
                _handles["selected"] = selectedMessage;
            }
        }

        public MailboxToolResult Execute(ChatToolCall call)
        {
            return Execute(call, CancellationToken.None);
        }

        internal MailboxToolResult Execute(
            ChatToolCall call,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (call?.function == null ||
                string.IsNullOrWhiteSpace(call.id))
            {
                return Error(
                    call?.id,
                    "MAILBOX_TOOL_CALL_INVALID",
                    "The model returned an invalid tool call.");
            }

            var name = call.function.name ?? string.Empty;
            if (!MailboxToolCatalog.IsApproved(name))
            {
                return Error(
                    call.id,
                    "MAILBOX_TOOL_NOT_ALLOWED",
                    "The requested mailbox tool is not allowed.");
            }

            try
            {
                var arguments = ParseArguments(
                    call.function.arguments);
                switch (name)
                {
                    case MailboxToolCatalog.SearchMailbox:
                        return Search(call.id, arguments);
                    case MailboxToolCatalog.ReadMessages:
                        return ReadMessages(
                            call.id,
                            arguments,
                            cancellationToken);
                    case MailboxToolCatalog.ReadThread:
                        return ReadThread(
                            call.id,
                            arguments,
                            cancellationToken);
                    default:
                        return Error(
                            call.id,
                            "MAILBOX_TOOL_NOT_ALLOWED",
                            "The requested mailbox tool is not allowed.");
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception exception)
            {
                Log.Error("MailboxTool." + name, exception);
                return Error(
                    call.id,
                    "MAILBOX_TOOL_FAILED",
                    DiagnosticDetails.ForException(
                        exception,
                        "MAILBOX_TOOL_FAILED"));
            }
        }

        internal MessageSnapshot ResolveHandle(string handle)
        {
            var boundedHandle = TextBoundary.SingleLine(
                handle,
                64);
            MessageSnapshot message;
            return _handles.TryGetValue(boundedHandle, out message)
                ? message
                : null;
        }

        private MailboxToolResult Search(
            string callId,
            IDictionary<string, object> arguments)
        {
            if (_workingSetOnly)
            {
                return Error(
                    callId,
                    "MAILBOX_WORKING_SET_LOCKED",
                    "A user-approved working set is active. Search and thread expansion are disabled for this request.");
            }

            if (_searchExecuted)
            {
                return Error(
                    callId,
                    "MAILBOX_SEARCH_LIMIT_REACHED",
                    "Only one mailbox search is allowed per request. Use /search to refine the working set.");
            }

            _searchExecuted = true;
            var query = GetString(arguments, "query", string.Empty);
            var folder = GetString(arguments, "folder", "all")
                .ToLowerInvariant();
            if (folder != "all" &&
                folder != "inbox" &&
                folder != "sent")
            {
                folder = "all";
            }

            var daysBack = GetInteger(
                arguments,
                "days_back",
                365,
                1,
                3650);
            var maxResults = GetInteger(
                arguments,
                "max_results",
                MailboxWorkingSet.MaxMessages,
                1,
                MailboxWorkingSet.MaxMessages);
            DateTime? receivedAfter;
            DateTime? receivedBefore;
            if (!TryGetLocalTimestamp(
                    arguments,
                    "received_after",
                    out receivedAfter) ||
                !TryGetLocalTimestamp(
                    arguments,
                    "received_before",
                    out receivedBefore))
            {
                return Error(
                    callId,
                    "MAILBOX_TIME_INVALID",
                    "Use ISO-8601 mailbox timestamps with an explicit UTC offset.");
            }

            if (receivedAfter.HasValue &&
                receivedBefore.HasValue &&
                receivedAfter.Value > receivedBefore.Value)
            {
                return Error(
                    callId,
                    "MAILBOX_TIME_RANGE_INVALID",
                    "received_after must not be later than received_before.");
            }

            var unreadOnly = GetBoolean(
                arguments,
                "unread_only",
                false);
            var hits = _mailbox.Search(
                query,
                folder,
                daysBack,
                maxResults + 1,
                receivedAfter,
                receivedBefore,
                unreadOnly);
            var truncated = hits.Count > maxResults;

            var results = new List<object>();
            foreach (var hit in hits.Take(maxResults))
            {
                var handle = Register(hit.Message);
                results.Add(new Dictionary<string, object>
                {
                    { "handle", handle },
                    { "folder", hit.FolderName },
                    { "subject", hit.Message.Subject },
                    { "from", hit.Message.Sender },
                    { "to", hit.Message.Recipients },
                    {
                        "received",
                        hit.Message.ReceivedAt?.ToString("O") ??
                        "unknown"
                    },
                    { "unread", hit.Message.IsUnread },
                    { "snippet", hit.Snippet }
                });
            }

            return Success(
                callId,
                new Dictionary<string, object>
                {
                    { "untrusted_email_data", true },
                    { "query", query },
                    { "folder", folder },
                    { "unread_only", unreadOnly },
                    {
                        "received_after",
                        receivedAfter?.ToString("O") ?? string.Empty
                    },
                    {
                        "received_before",
                        receivedBefore?.ToString("O") ?? string.Empty
                    },
                    { "result_count", results.Count },
                    { "truncated", truncated },
                    { "results", results }
                },
                "Mailbox search loaded " +
                results.Count.ToString(CultureInfo.InvariantCulture) +
                " result summaries.");
        }

        private MailboxToolResult ReadMessages(
            string callId,
            IDictionary<string, object> arguments,
            CancellationToken cancellationToken)
        {
            var handles = GetStringList(arguments, "handles")
                .Distinct(StringComparer.Ordinal)
                .Take(MailboxWorkingSet.MaxMessages)
                .ToArray();
            if (handles.Length == 0)
            {
                return Error(
                    callId,
                    "MAILBOX_HANDLES_REQUIRED",
                    "At least one message handle is required.");
            }

            var messages = new List<object>();
            var visionImages = new List<VisionImagePayload>();
            var attachmentBudget = new AttachmentReadBudget();
            var loadedCount = 0;
            foreach (var handle in handles)
            {
                cancellationToken.ThrowIfCancellationRequested();
                MessageSnapshot message;
                if (!_handles.TryGetValue(handle, out message))
                {
                    messages.Add(new Dictionary<string, object>
                    {
                        { "handle", handle },
                        { "error_code", "MAILBOX_HANDLE_UNKNOWN" }
                    });
                    continue;
                }

                if (_loadedBodyHandles.Contains(handle))
                {
                    messages.Add(new Dictionary<string, object>
                    {
                        { "handle", handle },
                        { "already_loaded", true },
                        {
                            "note",
                            "This body appears earlier in the conversation. " +
                            "Do not request it again."
                        }
                    });
                    continue;
                }

                if (_loadedBodyHandles.Count >=
                    MailboxWorkingSet.MaxMessages)
                {
                    messages.Add(new Dictionary<string, object>
                    {
                        { "handle", handle },
                        { "error_code", "MAILBOX_CONTEXT_LIMIT_REACHED" }
                    });
                    continue;
                }

                messages.Add(
                    SerializeMessage(
                        handle,
                        message,
                        MaxDirectMessageBodyCharacters,
                        visionImages,
                        cancellationToken,
                        attachmentBudget));
                _loadedBodyHandles.Add(handle);
                loadedCount++;
            }

            return Success(
                callId,
                new Dictionary<string, object>
                {
                    { "untrusted_email_data", true },
                    { "messages", messages }
                },
                "Loaded " +
                loadedCount.ToString(CultureInfo.InvariantCulture) +
                " new message bodies. Request total: " +
                _loadedBodyHandles.Count.ToString(
                    CultureInfo.InvariantCulture) +
                " of " +
                MailboxWorkingSet.MaxMessages.ToString(
                    CultureInfo.InvariantCulture) +
                ".",
                visionImages);
        }

        private MailboxToolResult ReadThread(
            string callId,
            IDictionary<string, object> arguments,
            CancellationToken cancellationToken)
        {
            if (_workingSetOnly)
            {
                return Error(
                    callId,
                    "MAILBOX_WORKING_SET_LOCKED",
                    "Thread expansion is disabled while a user-approved working set is active.");
            }

            var handle = GetString(arguments, "handle", string.Empty);
            MessageSnapshot source;
            if (!_handles.TryGetValue(handle, out source))
            {
                return Error(
                    callId,
                    "MAILBOX_HANDLE_UNKNOWN",
                    "The message handle is unknown or expired.");
            }

            var remaining = MailboxWorkingSet.MaxMessages -
                _loadedBodyHandles.Count;
            if (remaining <= 0)
            {
                return Error(
                    callId,
                    "MAILBOX_CONTEXT_LIMIT_REACHED",
                    MailboxWorkingSet.MaxMessages.ToString(
                        CultureInfo.InvariantCulture) +
                    " unique message bodies are already loaded for this request.");
            }

            var conversation = _mailbox.ReadConversation(
                source,
                MaxThreadMessages);
            var messages = new List<object>();
            var visionImages = new List<VisionImagePayload>();
            var attachmentBudget = new AttachmentReadBudget();
            foreach (var message in conversation)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var messageHandle = Register(message);
                if (_loadedBodyHandles.Contains(messageHandle))
                {
                    continue;
                }

                if (_loadedBodyHandles.Count >=
                    MailboxWorkingSet.MaxMessages)
                {
                    break;
                }

                messages.Add(
                    SerializeMessage(
                        messageHandle,
                        message,
                        MaxThreadMessageBodyCharacters,
                        visionImages,
                        cancellationToken,
                        attachmentBudget));
                _loadedBodyHandles.Add(messageHandle);
            }

            return Success(
                callId,
                new Dictionary<string, object>
                {
                    { "untrusted_email_data", true },
                    { "source_handle", handle },
                    { "message_count", messages.Count },
                    { "messages", messages }
                },
                "Loaded " +
                messages.Count.ToString(CultureInfo.InvariantCulture) +
                " new conversation messages. Request total: " +
                _loadedBodyHandles.Count.ToString(
                    CultureInfo.InvariantCulture) +
                " of " +
                MailboxWorkingSet.MaxMessages.ToString(
                    CultureInfo.InvariantCulture) +
                ".",
                visionImages);
        }

        private object SerializeMessage(
            string handle,
            MessageSnapshot message,
            int maximumBodyCharacters,
            IList<VisionImagePayload> visionImages,
            CancellationToken cancellationToken,
            AttachmentReadBudget attachmentBudget)
        {
            var payload = new Dictionary<string, object>
            {
                { "handle", handle },
                { "subject", message.Subject },
                { "from", message.Sender },
                { "to", message.Recipients },
                {
                    "received",
                    message.ReceivedAt?.ToString("O") ??
                    "unknown"
                },
                {
                    "body",
                    TextBoundary.PlainText(
                        message.Body,
                        maximumBodyCharacters)
                }
            };

            if (message.RemoteImageCount > 0)
            {
                payload["web_hosted_images_not_included"] =
                    message.RemoteImageCount;
                payload["web_hosted_images_note"] =
                    "The message body references web-hosted images by URL. " +
                    "Their bytes are not stored in the email, so Scribble " +
                    "cannot view them. Only embedded images and attachments " +
                    "are readable.";
            }

            var attachments = _mailbox.ReadAttachments(
                message,
                cancellationToken,
                attachmentBudget);
            if (attachments.Count > 0)
            {
                var serialized = new List<object>(attachments.Count);
                foreach (var attachment in attachments)
                {
                    var entry = new Dictionary<string, object>
                    {
                        { "filename", attachment.FileName },
                        { "kind", attachment.Kind },
                        { "content", attachment.Text }
                    };
                    if (attachment.Truncated)
                    {
                        entry["truncated"] = true;
                    }
                    if (attachment.ImageDataUrl.Length > 0)
                    {
                        entry["vision_available"] = true;
                        visionImages?.Add(
                            new VisionImagePayload(
                                attachment.FileName,
                                attachment.ImageDataUrl));
                    }

                    serialized.Add(entry);
                }

                payload["attachments"] = serialized;
            }

            return payload;
        }

        private string Register(MessageSnapshot message)
        {
            foreach (var pair in _handles)
            {
                if (pair.Value.EntryId.Equals(
                        message.EntryId,
                        StringComparison.Ordinal) &&
                    pair.Value.StoreId.Equals(
                        message.StoreId,
                        StringComparison.Ordinal))
                {
                    return pair.Key;
                }
            }

            var handle = "m" +
                _nextHandle.ToString(CultureInfo.InvariantCulture);
            _nextHandle++;
            _handles[handle] = message;
            return handle;
        }

        private MailboxToolResult Success(
            string callId,
            object payload,
            string status,
            IReadOnlyList<VisionImagePayload> visionImages = null)
        {
            var json = _serializer.Serialize(payload);
            var allowedCharacters = Math.Max(
                ContextScale.Scaled(
                    TextBoundary.MaxToolResultCharacters),
                MailboxWorkingSet.MaxMessages *
                    PerMessageResultCharacters);
            if (json.Length > allowedCharacters)
            {
                return Error(
                    callId,
                    "MAILBOX_TOOL_RESULT_TOO_LARGE",
                    "The bounded mailbox result was still too large to return safely.");
            }

            return new MailboxToolResult(
                callId,
                json,
                status,
                visionImages);
        }

        private MailboxToolResult Error(
            string callId,
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
                        }
                    }),
                "[" + code + "] " + message);
        }

        private IDictionary<string, object> ParseArguments(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                return new Dictionary<string, object>();
            }

            try
            {
                var parsed = _serializer.DeserializeObject(json) as
                    IDictionary<string, object>;
                if (parsed == null)
                {
                    throw new InvalidOperationException(
                        "Tool arguments must be a JSON object.");
                }

                return parsed;
            }
            catch (Exception exception)
            {
                throw new AiEndpointException(
                    "TOOL_ARGUMENTS_INVALID_JSON",
                    "The model returned invalid JSON tool arguments.",
                    exception,
                    responseSnippet: json);
            }
        }

        private static string GetString(
            IDictionary<string, object> arguments,
            string key,
            string fallback)
        {
            object value;
            return arguments.TryGetValue(key, out value)
                ? TextBoundary.PlainText(
                    Convert.ToString(value),
                    1000)
                : fallback;
        }

        private static int GetInteger(
            IDictionary<string, object> arguments,
            string key,
            int fallback,
            int minimum,
            int maximum)
        {
            object value;
            int parsed;
            if (!arguments.TryGetValue(key, out value) ||
                !int.TryParse(
                    Convert.ToString(value),
                    NumberStyles.Integer,
                    CultureInfo.InvariantCulture,
                    out parsed))
            {
                return fallback;
            }

            return Math.Max(minimum, Math.Min(maximum, parsed));
        }

        private static bool GetBoolean(
            IDictionary<string, object> arguments,
            string key,
            bool fallback)
        {
            object value;
            bool parsed;
            return arguments.TryGetValue(key, out value) &&
                   bool.TryParse(Convert.ToString(value), out parsed)
                ? parsed
                : fallback;
        }

        private static bool TryGetLocalTimestamp(
            IDictionary<string, object> arguments,
            string key,
            out DateTime? value)
        {
            value = null;
            object raw;
            if (!arguments.TryGetValue(key, out raw) ||
                string.IsNullOrWhiteSpace(Convert.ToString(raw)))
            {
                return true;
            }

            var text = Convert.ToString(raw).Trim();
            var hasZulu = text.EndsWith(
                "Z",
                StringComparison.OrdinalIgnoreCase);
            var hasOffset = text.Length >= 6 &&
                (text[text.Length - 6] == '+' ||
                 text[text.Length - 6] == '-') &&
                text[text.Length - 3] == ':';
            DateTimeOffset parsed;
            if (!hasZulu && !hasOffset)
            {
                return false;
            }

            if (!DateTimeOffset.TryParse(
                    text,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AllowWhiteSpaces |
                    DateTimeStyles.RoundtripKind,
                    out parsed))
            {
                return false;
            }

            value = TimeZoneInfo.ConvertTime(
                parsed,
                TimeZoneInfo.Local).DateTime;
            return true;
        }

        private static IEnumerable<string> GetStringList(
            IDictionary<string, object> arguments,
            string key)
        {
            object value;
            if (!arguments.TryGetValue(key, out value))
            {
                return new string[0];
            }

            var array = value as object[];
            if (array != null)
            {
                return array
                    .Select(item =>
                        TextBoundary.PlainText(
                            Convert.ToString(item),
                            100))
                    .Where(item => item.Length > 0);
            }

            var list = value as ArrayList;
            if (list != null)
            {
                return list
                    .Cast<object>()
                    .Select(item =>
                        TextBoundary.PlainText(
                            Convert.ToString(item),
                            100))
                    .Where(item => item.Length > 0);
            }

            return new[]
            {
                TextBoundary.PlainText(
                    Convert.ToString(value),
                    100)
            }.Where(item => item.Length > 0);
        }
    }
}

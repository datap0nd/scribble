using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
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
        private static int MaxThreadMessages
        {
            get { return MailboxSearchBudget.MaxThreadMessages; }
        }

        // Bodies are the expensive intake, so they keep a budget of
        // their own. An approved working set is read in full and
        // nothing else; a free request may read further than the
        // pinned set, up to the reviewed body budget.
        private int MaxBodyMessages
        {
            get
            {
                return _workingSetOnly
                    ? MailboxWorkingSet.MaxMessages
                    : MailboxSearchBudget.MaxBodyMessages;
            }
        }

        private static int ResultBudget
        {
            get
            {
                return ContextScale.Scaled(
                    TextBoundary.MaxToolResultCharacters);
            }
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
        private int _searchExecutedCount;
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
                        return ReadMessages(call.id, arguments);
                    case MailboxToolCatalog.ReadThread:
                        return ReadThread(call.id, arguments);
                    default:
                        return Error(
                            call.id,
                            "MAILBOX_TOOL_NOT_ALLOWED",
                            "The requested mailbox tool is not allowed.");
                }
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

            if (_searchExecutedCount >=
                MailboxSearchBudget.MaxSearchesPerRequest)
            {
                return Error(
                    callId,
                    "MAILBOX_SEARCH_LIMIT_REACHED",
                    MailboxSearchBudget.MaxSearchesPerRequest.ToString(
                        CultureInfo.InvariantCulture) +
                    " mailbox searches have already run for this request. Widen max_results instead of searching again.");
            }

            _searchExecutedCount++;
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
                MailboxSearchBudget.RecommendedResults,
                1,
                MailboxSearchBudget.MaxResults);
            var hits = _mailbox.Search(
                query,
                folder,
                daysBack,
                maxResults);

            var subjectCharacters =
                MailboxSearchBudget.SubjectCharacters(maxResults);
            var senderCharacters =
                MailboxSearchBudget.SenderCharacters(maxResults);
            var withRecipients =
                MailboxSearchBudget.IncludesRecipients(maxResults);
            var results = new List<object>();
            foreach (var hit in hits)
            {
                var handle = Register(hit.Message);
                var summary = new Dictionary<string, object>
                {
                    { "handle", handle },
                    { "folder", hit.FolderName },
                    {
                        "subject",
                        TextBoundary.SingleLine(
                            hit.Message.Subject,
                            subjectCharacters)
                    },
                    {
                        "from",
                        TextBoundary.SingleLine(
                            hit.Message.Sender,
                            senderCharacters)
                    },
                    {
                        "received",
                        hit.Message.ReceivedAt?.ToString("O") ??
                        "unknown"
                    }
                };
                if (withRecipients)
                {
                    summary["to"] = hit.Message.Recipients;
                }

                if (hit.Snippet.Length > 0)
                {
                    summary["snippet"] = hit.Snippet;
                }

                results.Add(summary);
            }

            // A wide sweep is packed to whatever the request's
            // tool-result budget actually holds rather than
            // refused outright: the model gets the newest
            // summaries and is told the tail was dropped.
            var returned = PackSummaries(results);
            var trimmed = returned.Count < results.Count;
            var payload = new Dictionary<string, object>
            {
                { "untrusted_email_data", true },
                { "query", query },
                { "folder", folder },
                { "requested_count", maxResults },
                { "match_count", results.Count },
                { "result_count", returned.Count },
                { "results", returned }
            };
            if (trimmed)
            {
                payload["truncated_for_context"] = true;
            }

            return Success(
                callId,
                payload,
                "Mailbox search loaded " +
                returned.Count.ToString(CultureInfo.InvariantCulture) +
                " result summaries" +
                (trimmed
                    ? " of " +
                      results.Count.ToString(
                          CultureInfo.InvariantCulture) +
                      " matches; the rest did not fit this request."
                    : "."));
        }

        // Keeps the newest summaries that fit the tool-result
        // budget. Each entry is measured once, so a 500-result
        // sweep costs one pass rather than a serialize-and-retry
        // loop.
        private List<object> PackSummaries(List<object> results)
        {
            var budget = ResultBudget - 600;
            var used = 0;
            var packed = new List<object>(results.Count);
            foreach (var result in results)
            {
                var length = _serializer.Serialize(result).Length + 1;
                if (packed.Count > 0 &&
                    used + length > budget)
                {
                    break;
                }

                used += length;
                packed.Add(result);
            }

            return packed;
        }

        private MailboxToolResult ReadMessages(
            string callId,
            IDictionary<string, object> arguments)
        {
            var handles = GetStringList(arguments, "handles")
                .Distinct(StringComparer.Ordinal)
                .Take(MaxBodyMessages)
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
            var loadedCount = 0;
            var used = 0;
            foreach (var handle in handles)
            {
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

                if (_loadedBodyHandles.Count >= MaxBodyMessages)
                {
                    messages.Add(new Dictionary<string, object>
                    {
                        { "handle", handle },
                        { "error_code", "MAILBOX_CONTEXT_LIMIT_REACHED" }
                    });
                    continue;
                }

                // Bodies carry attachment text, so a wide read is
                // measured as it is built: what does not fit is
                // deferred with its handle intact rather than
                // costing the whole result.
                var pendingImages = new List<VisionImagePayload>();
                var serialized = SerializeMessage(
                    handle,
                    message,
                    MailboxSearchBudget.BodyCharacters(
                        handles.Length,
                        ResultBudget,
                        MaxDirectMessageBodyCharacters),
                    pendingImages);
                var length =
                    _serializer.Serialize(serialized).Length + 1;
                if (loadedCount > 0 &&
                    used + length > ResultBudget - 600)
                {
                    messages.Add(new Dictionary<string, object>
                    {
                        { "handle", handle },
                        { "deferred_for_context", true },
                        {
                            "note",
                            "This body did not fit the remaining result " +
                            "budget. Request it again in a later step."
                        }
                    });
                    break;
                }

                used += length;
                messages.Add(serialized);
                foreach (var image in pendingImages)
                {
                    visionImages.Add(image);
                }

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
                MaxBodyMessages.ToString(
                    CultureInfo.InvariantCulture) +
                ".",
                visionImages);
        }

        private MailboxToolResult ReadThread(
            string callId,
            IDictionary<string, object> arguments)
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

            var remaining = MaxBodyMessages -
                _loadedBodyHandles.Count;
            if (remaining <= 0)
            {
                return Error(
                    callId,
                    "MAILBOX_CONTEXT_LIMIT_REACHED",
                    MaxBodyMessages.ToString(
                        CultureInfo.InvariantCulture) +
                    " unique message bodies are already loaded for this request.");
            }

            var conversation = _mailbox.ReadConversation(
                source,
                Math.Min(remaining, MaxThreadMessages));
            var messages = new List<object>();
            var visionImages = new List<VisionImagePayload>();
            foreach (var message in conversation)
            {
                var messageHandle = Register(message);
                if (_loadedBodyHandles.Contains(messageHandle))
                {
                    continue;
                }

                if (_loadedBodyHandles.Count >= MaxBodyMessages)
                {
                    break;
                }

                messages.Add(
                    SerializeMessage(
                        messageHandle,
                        message,
                        MaxThreadMessageBodyCharacters,
                        visionImages));
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
                MaxBodyMessages.ToString(
                    CultureInfo.InvariantCulture) +
                ".",
                visionImages);
        }

        private object SerializeMessage(
            string handle,
            MessageSnapshot message,
            int maximumBodyCharacters,
            IList<VisionImagePayload> visionImages)
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

            var attachments = _mailbox.ReadAttachments(message);
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
            if (json.Length >
                ContextScale.Scaled(
                    TextBoundary.MaxToolResultCharacters))
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

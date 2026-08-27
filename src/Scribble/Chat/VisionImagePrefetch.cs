using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;
using Scribble.Configuration;
using Scribble.Outlook;
using Scribble.Security;

namespace Scribble.Chat
{
    public static class VisionImagePrefetch
    {
        private const string PrefetchToolCallId = "prefetch_images";

        public static bool TryInject(
            ChatCompletionRequest request,
            string modelId,
            MailboxToolHost mailboxTools,
            MessageSnapshot selectedMessage,
            IReadOnlyList<MessageSnapshot> workingMessages)
        {
            if (request == null ||
                mailboxTools == null ||
                !ModelCatalog.IsVisionCapable(modelId))
            {
                return false;
            }

            var handles = CollectImageHandles(
                selectedMessage,
                workingMessages);
            if (handles.Count == 0)
            {
                return false;
            }

            var arguments = BuildArguments(handles);
            var result = mailboxTools.Execute(
                new ChatToolCall
                {
                    id = PrefetchToolCallId,
                    type = "function",
                    function = new ChatToolCallFunction
                    {
                        name = MailboxToolCatalog.ReadMessages,
                        arguments = arguments
                    }
                });
            if (result.VisionImages == null ||
                result.VisionImages.Count == 0)
            {
                return false;
            }

            request.messages.Add(new ChatCompletionAssistantToolMessage
            {
                role = "assistant",
                content = string.Empty,
                tool_calls = new List<ChatToolCall>
                {
                    new ChatToolCall
                    {
                        id = PrefetchToolCallId,
                        type = "function",
                        function = new ChatToolCallFunction
                        {
                            name = MailboxToolCatalog.ReadMessages,
                            arguments = arguments
                        }
                    }
                }
            });
            request.messages.Add(new ChatCompletionToolResultMessage
            {
                role = "tool",
                tool_call_id = PrefetchToolCallId,
                content = TextBoundary.PlainText(
                    result.Content,
                    ContextScale.Scaled(
                        TextBoundary.MaxToolResultCharacters))
            });
            VisionAttachmentExchange.AppendVisionContext(
                request,
                modelId,
                new[] { result });
            return true;
        }

        private static string BuildArguments(IReadOnlyList<string> handles)
        {
            var serializer = new JavaScriptSerializer();
            return serializer.Serialize(
                new Dictionary<string, object>
                {
                    { "handles", handles.ToArray() }
                });
        }

        private static List<string> CollectImageHandles(
            MessageSnapshot selectedMessage,
            IReadOnlyList<MessageSnapshot> workingMessages)
        {
            var workingSet = MailboxWorkingSet.Normalize(
                workingMessages);
            if (workingSet.Count > 0)
            {
                var handles = new List<string>();
                for (var index = 0;
                     index < workingSet.Count;
                     index++)
                {
                    if (ModelRouting.MessageHasImageAttachment(
                            workingSet[index]))
                    {
                        handles.Add(
                            MailboxWorkingSet.HandleAt(index));
                    }
                }

                return handles;
            }

            if (ModelRouting.MessageHasImageAttachment(
                    selectedMessage))
            {
                return new List<string> { "selected" };
            }

            return new List<string>();
        }
    }
}

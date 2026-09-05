using System.Collections.Generic;

namespace Scribble.Chat
{
    public sealed class ChatTurn
    {
        public ChatTurn(string role, string content)
        {
            Role = role;
            Content = content;
        }

        public string Role { get; }

        public string Content { get; }
    }

    public sealed class ChatCompletionRequest
    {
        public string model { get; set; }

        public List<object> messages { get; set; }

        public bool stream { get; set; }

        public List<ChatToolDefinition> tools { get; set; }

        public object tool_choice { get; set; }

        public int? max_tokens { get; set; }

        public double? temperature { get; set; }

        public bool? parallel_tool_calls { get; set; }
    }

    public sealed class ChatCompletionInputMessage
    {
        public string role { get; set; }

        public object content { get; set; }
    }

    public sealed class ChatMultimodalTextPart
    {
        public string type { get; set; }

        public string text { get; set; }
    }

    public sealed class ChatMultimodalImagePart
    {
        public string type { get; set; }

        public ChatMultimodalImageUrl image_url { get; set; }
    }

    public sealed class ChatMultimodalImageUrl
    {
        public string url { get; set; }

        public string detail { get; set; }
    }

    public sealed class VisionImagePayload
    {
        public VisionImagePayload(
            string fileName,
            string dataUrl)
        {
            FileName = fileName ?? string.Empty;
            DataUrl = dataUrl ?? string.Empty;
        }

        public string FileName { get; }

        public string DataUrl { get; }
    }

    public sealed class ChatCompletionAssistantToolMessage
    {
        public string role { get; set; }

        public string content { get; set; }

        public List<ChatToolCall> tool_calls { get; set; }
    }

    public sealed class ChatCompletionToolResultMessage
    {
        public string role { get; set; }

        public string tool_call_id { get; set; }

        public string content { get; set; }
    }

    public sealed class ChatCompletionResponse
    {
        public List<ChatCompletionChoice> choices { get; set; }

        public ChatCompletionError error { get; set; }
    }

    public sealed class ModelListResponse
    {
        public List<ModelListItem> data { get; set; }
    }

    public sealed class ModelListItem
    {
        public string id { get; set; }
    }

    public sealed class ChatCompletionChoice
    {
        public ChatCompletionResponseMessage message { get; set; }
    }

    public sealed class ChatCompletionResponseMessage
    {
        [System.Web.Script.Serialization.ScriptIgnore]
        public string RawContent { get; set; }

        public string role { get; set; }

        public string content { get; set; }

        public List<ChatToolCall> tool_calls { get; set; }
    }

    public sealed class ChatCompletionError
    {
        public string message { get; set; }

        public string type { get; set; }

        public string code { get; set; }
    }

    public sealed class ChatToolDefinition
    {
        public string type { get; set; }

        public ChatToolFunctionDefinition function { get; set; }
    }

    public sealed class ChatToolFunctionDefinition
    {
        public string name { get; set; }

        public string description { get; set; }

        public object parameters { get; set; }
    }

    public sealed class ChatToolCall
    {
        public string id { get; set; }

        public string type { get; set; }

        public ChatToolCallFunction function { get; set; }
    }

    public sealed class ChatToolCallFunction
    {
        public string name { get; set; }

        public string arguments { get; set; }
    }

    public sealed class MailboxToolResult
    {
        public MailboxToolResult(
            string toolCallId,
            string content,
            string statusText,
            IReadOnlyList<VisionImagePayload> visionImages = null,
            int contentCharacterLimit = 0)
        {
            ToolCallId = toolCallId ?? string.Empty;
            Content = content ?? string.Empty;
            StatusText = statusText ?? string.Empty;
            VisionImages = visionImages ??
                new VisionImagePayload[0];
            ContentCharacterLimit = contentCharacterLimit;
        }

        public string ToolCallId { get; }

        public string Content { get; }

        public string StatusText { get; }

        public IReadOnlyList<VisionImagePayload> VisionImages { get; }

        public int ContentCharacterLimit { get; }
    }

    // One completed browser tool round replayed by the extension:
    // the assistant turn that requested tools, and the bounded
    // results the panel and host produced for it.
    public sealed class BrowserExchangeTurn
    {
        public string AssistantContent { get; set; }

        public List<ChatToolCall> ToolCalls { get; set; }

        public List<BrowserExchangeResult> Results { get; set; }
    }

    public sealed class BrowserExchangeResult
    {
        public string Id { get; set; }

        public string Content { get; set; }
    }

    // A tool call the extension must execute (or display) before
    // the conversation can continue.
    public sealed class BrowserToolRequest
    {
        public BrowserToolRequest(
            string id,
            string name,
            string arguments)
        {
            Id = id ?? string.Empty;
            Name = name ?? string.Empty;
            Arguments = arguments ?? string.Empty;
        }

        public string Id { get; }

        public string Name { get; }

        public string Arguments { get; }
    }
}

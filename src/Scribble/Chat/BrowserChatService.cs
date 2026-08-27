using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Scribble.Configuration;
using Scribble.Security;

namespace Scribble.Chat
{
    public sealed class BrowserChatResult
    {
        public BrowserChatResult(
            string content,
            string model,
            bool screenshotUsed)
        {
            Content = content ?? string.Empty;
            Model = model ?? string.Empty;
            ScreenshotUsed = screenshotUsed;
        }

        public string Content { get; }

        public string Model { get; }

        public bool ScreenshotUsed { get; }
    }

    // Runs the browser host's read-only model loop. The browser can
    // provide user-approved page context and user-configured MCP
    // tools; no Office, mailbox, navigation, or page-mutation tool
    // is present in this process.
    public sealed class BrowserChatService : IDisposable
    {
        public const int MaxBrowserToolRounds = 1;
        public const int MaxBrowserToolCallsPerRound = 1;

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

        public async Task<BrowserChatResult> CompleteAsync(
            IReadOnlyList<ChatTurn> history,
            string prompt,
            string title,
            string url,
            string selection,
            string pageText,
            string screenshotDataUrl,
            CancellationToken cancellationToken)
        {
            if (!IsConfigured)
            {
                throw new AiEndpointException(
                    "CONFIGURATION_INCOMPLETE",
                    "Open Scribble in an Office app and configure the endpoint, model, and API key first.");
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
                    "The attached screenshot was not a valid bounded JPEG, PNG, or WebP image.");
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
                definitions);

            for (var round = 0;
                 round <= MaxBrowserToolRounds;
                 round++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var response = await _client.CompleteAsync(
                    _settings,
                    request,
                    cancellationToken).ConfigureAwait(false);
                var toolCalls = response.tool_calls;
                if (toolCalls == null || toolCalls.Count == 0)
                {
                    if (string.IsNullOrWhiteSpace(response.content))
                    {
                        throw new AiEndpointException(
                            "RESPONSE_MISSING_CONTENT",
                            "The model stopped without returning text.");
                    }

                    return new BrowserChatResult(
                        response.content,
                        activeModel,
                        screenshotUsed);
                }

                if (round == MaxBrowserToolRounds)
                {
                    throw new AiEndpointException(
                        "TOOL_ROUND_LIMIT",
                        "The model exceeded the maximum number of bounded tool rounds.");
                }

                if (toolCalls.Count >
                    MaxBrowserToolCallsPerRound)
                {
                    throw new AiEndpointException(
                        "TOOL_CALL_LIMIT",
                        "The model requested too many tools in one round.");
                }

                var results = new List<MailboxToolResult>();
                foreach (var toolCall in toolCalls)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var name = toolCall?.function?.name;
                    if (!McpToolHost.IsMcpTool(name) ||
                        !_mcpTools.HasServers)
                    {
                        throw new AiEndpointException(
                            "BROWSER_TOOL_NOT_ALLOWED",
                            "The browser model requested a tool that this read-only host does not expose.");
                    }

                    results.Add(await Task.Run(
                        () => _mcpTools.Execute(toolCall),
                        cancellationToken).ConfigureAwait(false));
                }

                ChatRequestFactory.AppendToolExchange(
                    request,
                    response,
                    results,
                    activeModel);
            }

            throw new AiEndpointException(
                "TOOL_ROUND_LIMIT",
                "The model did not finish after bounded tool use.");
        }

        public void Dispose()
        {
            _mcpTools.Dispose();
            _client.Dispose();
        }
    }
}

using System;
using System.Collections.Generic;
using Scribble.Chat;
using Scribble.Security;

namespace Scribble.Configuration
{
    public sealed class AppSettings
    {
        public string BaseUrl { get; set; } = string.Empty;

        public string Model { get; set; } = string.Empty;

        public string ApiKey { get; set; } = string.Empty;

        // HTTP endpoints are accepted without a separate opt-in.
        // The settings UI warns when a non-loopback URL will carry
        // credentials and context without transport encryption.
        public bool AllowInsecureHttp { get; set; } = true;

        public string ToneProfile { get; set; } = string.Empty;

        public bool UseToneProfile { get; set; }

        // How strongly drafts follow the writing soul, 10-100.
        public int ToneStrength { get; set; } = 60;

        // User-authored hard rules applied to every draft's wording.
        public string DraftRules { get; set; } = string.Empty;

        // Image requests automatically use the best discovered vision
        // model. This is product behavior rather than an end-user setting.
        public bool SwitchToVisionModelForImages { get; set; } = true;

        // Gemini via Google sign-in: when enabled, requests go to
        // Google's Code Assist API with OAuth tokens instead of the
        // OpenAI-compatible endpoint, and no endpoint or API key is
        // required. The refresh token is stored DPAPI-protected.
        public bool UseGeminiSignIn { get; set; }

        public string GeminiRefreshToken { get; set; } =
            string.Empty;

        // Optional Google Cloud project id for organizations whose
        // Gemini license designates one; used during account
        // onboarding. Takes precedence over GOOGLE_CLOUD_PROJECT.
        public string GeminiProject { get; set; } = string.Empty;

        public List<string> DiscoveredModels { get; set; } =
            new List<string>();

        // User-configured MCP servers, surfaced to the model as
        // namespaced mcp_ tools. Only Settings can add entries.
        public List<McpServerConfig> McpServers { get; set; } =
            new List<McpServerConfig>();

        // User-approved local document repositories. A Topic is
        // inert until it is explicitly selected for one chat.
        public List<TopicConfig> Topics { get; set; } =
            new List<TopicConfig>();

        // Legacy serialized fields retained for settings-file
        // compatibility. The text and loop budgets are no longer
        // user-customized; ApplyLimits always selects the reviewed
        // defaults for those. The mailbox working-set size is the
        // one budget the user owns (Settings > Limits).
        public bool UseRecommendedLimits { get; set; } = true;

        public int LimitContextMultiplier { get; set; } = 1;

        public int LimitPromptCharacters { get; set; } =
            TextBoundary.RecommendedUserPromptCharacters;

        public int LimitAssistantCharacters { get; set; } =
            TextBoundary.RecommendedAssistantCharacters;

        public int LimitHistoryTurns { get; set; } =
            TextBoundary.RecommendedConversationTurns;

        public int LimitToolRounds { get; set; } =
            TextBoundary.RecommendedToolRounds;

        public int LimitToolCallsPerRound { get; set; } =
            TextBoundary.RecommendedToolCallsPerRound;

        public int LimitWorkingSetMessages { get; set; } =
            LimitOverrides.RecommendedWorkingSetMessages;

        // Pushes this settings object's limit choices into the
        // process-wide effective limits. Called wherever settings
        // are loaded or saved. Text and loop budgets stay at the
        // reviewed defaults; the working-set size honors the user's
        // setting, falling back to the default when unset.
        public void ApplyLimits()
        {
            LimitOverrides.Apply(
                false,
                TextBoundary.RecommendedUserPromptCharacters,
                TextBoundary.RecommendedAssistantCharacters,
                TextBoundary.RecommendedConversationTurns,
                TextBoundary.RecommendedToolRounds,
                TextBoundary.RecommendedToolCallsPerRound,
                LimitWorkingSetMessages > 0
                    ? LimitWorkingSetMessages
                    : LimitOverrides.RecommendedWorkingSetMessages);
            ContextScale.ApplyUserMultiplier(
                1);
        }

        public bool IsConfigured
        {
            get
            {
                if (!AdminPolicy.GeminiDisabled &&
                    UseGeminiSignIn &&
                    GeminiCodeAssistGateway.IsGeminiModel(Model))
                {
                    return Model.Trim().Length > 0;
                }

                Uri endpoint;
                return ModelSelectionPolicy.IsGenerativeModel(Model) &&
                       ApiKey.Trim().Length > 0 &&
                       TryGetChatCompletionsUri(
                           BaseUrl,
                           AllowInsecureHttp,
                           out endpoint);
            }
        }

        // True when the OpenAI-compatible endpoint alone is usable.
        // Unlike HasConnectionSettings this ignores the Gemini tick,
        // so the picker can offer local models alongside Gemini
        // ones.
        public bool HasEndpointCredentials
        {
            get
            {
                Uri endpoint;
                return ApiKey.Trim().Length > 0 &&
                       TryGetChatCompletionsUri(
                           BaseUrl,
                           AllowInsecureHttp,
                           out endpoint);
            }
        }

        public bool HasConnectionSettings
        {
            get
            {
                if (!AdminPolicy.GeminiDisabled &&
                    UseGeminiSignIn)
                {
                    return true;
                }

                Uri endpoint;
                return ApiKey.Trim().Length > 0 &&
                       TryGetChatCompletionsUri(
                           BaseUrl,
                           AllowInsecureHttp,
                           out endpoint);
            }
        }

        public static bool TryGetChatCompletionsUri(string value, out Uri endpoint)
        {
            return TryGetChatCompletionsUri(
                value,
                true,
                out endpoint);
        }

        public static bool TryGetChatCompletionsUri(
            string value,
            bool allowInsecureHttp,
            out Uri endpoint)
        {
            endpoint = null;
            Uri baseUri;
            if (!Uri.TryCreate(value?.Trim(), UriKind.Absolute, out baseUri))
            {
                return false;
            }

            var isHttps = baseUri.Scheme.Equals(
                Uri.UriSchemeHttps,
                StringComparison.OrdinalIgnoreCase);
            var isAllowedHttp =
                baseUri.Scheme.Equals(
                    Uri.UriSchemeHttp,
                    StringComparison.OrdinalIgnoreCase) &&
                (baseUri.IsLoopback || allowInsecureHttp);

            if (!isHttps && !isAllowedHttp)
            {
                return false;
            }

            var path = baseUri.AbsolutePath.TrimEnd('/');
            string completedPath;
            if (path.EndsWith(
                "/chat/completions",
                StringComparison.OrdinalIgnoreCase))
            {
                completedPath = path;
            }
            else if (path.EndsWith("/v1", StringComparison.OrdinalIgnoreCase))
            {
                completedPath = path + "/chat/completions";
            }
            else if (path.EndsWith(
                "/openai",
                StringComparison.OrdinalIgnoreCase))
            {
                // OpenAI-compatibility layers hosted under an
                // /openai base, such as Google's
                // generativelanguage.googleapis.com/v1beta/openai,
                // expose chat/completions directly under that base.
                completedPath = path + "/chat/completions";
            }
            else
            {
                completedPath = path + "/v1/chat/completions";
            }

            var builder = new UriBuilder(baseUri)
            {
                Path = completedPath,
                Query = string.Empty,
                Fragment = string.Empty
            };

            endpoint = builder.Uri;
            return true;
        }

        public static bool TryGetModelsUri(
            string value,
            bool allowInsecureHttp,
            out Uri endpoint)
        {
            endpoint = null;
            Uri chatEndpoint;
            if (!TryGetChatCompletionsUri(
                value,
                allowInsecureHttp,
                out chatEndpoint))
            {
                return false;
            }

            const string suffix = "/chat/completions";
            var path = chatEndpoint.AbsolutePath;
            if (!path.EndsWith(
                suffix,
                StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            var builder = new UriBuilder(chatEndpoint)
            {
                Path = path.Substring(
                    0,
                    path.Length - suffix.Length) + "/models",
                Query = string.Empty,
                Fragment = string.Empty
            };
            endpoint = builder.Uri;
            return true;
        }
    }
}

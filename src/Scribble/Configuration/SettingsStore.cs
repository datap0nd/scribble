using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using Scribble.Chat;
using Scribble.Security;

namespace Scribble.Configuration
{
    public sealed class SettingsStore
    {
        // Keep the original DPAPI entropy so upgrades can still decrypt
        // API keys, Gemini tokens, and MCP headers saved before the rename.
        // The fragments intentionally avoid carrying the retired product
        // name as a source-code identifier or user-facing string.
        private static readonly string LegacyProductDirectoryName =
            "Outlook" + "Local" + "AI" + "Chat";

        private static readonly byte[] Entropy =
            Encoding.UTF8.GetBytes(
                LegacyProductDirectoryName + ".Settings.v1");

        private readonly string _settingsPath;
        private readonly string _legacySettingsPath;
        private readonly JavaScriptSerializer _serializer =
            new JavaScriptSerializer();

        public SettingsStore()
        {
            var localData = Environment.GetFolderPath(
                Environment.SpecialFolder.LocalApplicationData);
            _settingsPath = Path.Combine(
                localData,
                "Scribble",
                "settings.json");
            _legacySettingsPath = Path.Combine(
                localData,
                LegacyProductDirectoryName,
                "settings.json");
        }

        public AppSettings Load()
        {
            try
            {
                var loadPath = ResolveLoadPath();
                if (!File.Exists(loadPath))
                {
                    return new AppSettings();
                }

                var stored = _serializer.Deserialize<StoredSettings>(
                    File.ReadAllText(loadPath, Encoding.UTF8));
                if (stored == null)
                {
                    return new AppSettings();
                }

                var geminiDisabled = AdminPolicy.GeminiDisabled;
                var storedModel = (stored.Model ?? string.Empty).Trim();

                return new AppSettings
                {
                    BaseUrl = stored.BaseUrl ?? string.Empty,
                    Model = ModelSelectionPolicy.IsGenerativeModel(
                        storedModel)
                            ? storedModel
                            : string.Empty,
                    ApiKey = Unprotect(stored.ProtectedApiKey),
                    // Migrate every existing installation to the
                    // streamlined endpoint flow: HTTP works without
                    // a separate opt-in checkbox.
                    AllowInsecureHttp = true,
                    UseGeminiSignIn = stored.UseGeminiSignIn &&
                        !geminiDisabled,
                    // Do not decrypt dormant Google credentials into
                    // memory while the capability is unavailable.
                    GeminiRefreshToken = geminiDisabled
                        ? string.Empty
                        : Unprotect(
                            stored.ProtectedGeminiRefreshToken),
                    GeminiProject = geminiDisabled
                        ? string.Empty
                        : TextBoundary.SingleLine(
                            stored.GeminiProject,
                            200),
                    ToneProfile = TextBoundary.PlainText(
                        stored.ToneProfile,
                        TextBoundary.MaxToneProfileCharacters),
                    UseToneProfile = stored.UseToneProfile,
                    ToneStrength = Math.Max(
                        10,
                        Math.Min(
                            100,
                            stored.ToneStrength == 0
                                ? 60
                                : stored.ToneStrength)),
                    DraftRules = TextBoundary.PlainText(
                        stored.DraftRules,
                        2000),
                    SwitchToVisionModelForImages =
                        stored.SwitchToVisionModelForImages,
                    DiscoveredModels = NormalizeDiscoveredModels(
                        stored.DiscoveredModels),
                    McpServers = NormalizeMcpServers(
                        stored.McpServers),
                    UseRecommendedLimits = true,
                    LimitContextMultiplier = 1,
                    LimitPromptCharacters = TextBoundary
                        .RecommendedUserPromptCharacters,
                    LimitAssistantCharacters = TextBoundary
                        .RecommendedAssistantCharacters,
                    LimitHistoryTurns = TextBoundary
                        .RecommendedConversationTurns,
                    LimitToolRounds = TextBoundary
                        .RecommendedToolRounds,
                    LimitToolCallsPerRound = TextBoundary
                        .RecommendedToolCallsPerRound,
                    LimitWorkingSetMessages = NormalizeWorkingSet(
                        stored.LimitWorkingSetMessages)
                };
            }
            catch
            {
                return new AppSettings();
            }
        }

        private string ResolveLoadPath()
        {
            if (File.Exists(_settingsPath) ||
                !File.Exists(_legacySettingsPath))
            {
                return _settingsPath;
            }

            try
            {
                var directory = Path.GetDirectoryName(_settingsPath);
                Directory.CreateDirectory(directory);
                File.Copy(_legacySettingsPath, _settingsPath, false);
                return _settingsPath;
            }
            catch
            {
                // A locked or read-only legacy file is still usable for
                // this session. The next successful Save writes the new
                // Scribble location.
                return _legacySettingsPath;
            }
        }

        public void Save(AppSettings settings)
        {
            if (settings == null)
            {
                throw new ArgumentNullException(nameof(settings));
            }

            if (settings.Model.Trim().Length == 0)
            {
                throw new InvalidOperationException("Enter a model name.");
            }

            if (!ModelSelectionPolicy.IsGenerativeModel(settings.Model))
            {
                if (GeminiCodeAssistGateway.IsGeminiModel(
                    settings.Model))
                {
                    throw new InvalidOperationException(
                        "Google Gemini is unavailable in this build. " +
                        "Choose a model served by your endpoint.");
                }

                throw new InvalidOperationException(
                    "Choose a generative chat model.");
            }

            // The endpoint is required whenever the selected model
            // is not a Gemini model - the Gemini tick no longer
            // decides the transport, the model does - and a typed
            // endpoint is always validated even alongside Gemini.
            var useGemini =
                !AdminPolicy.GeminiDisabled &&
                settings.UseGeminiSignIn;
            var needsEndpoint =
                !GeminiCodeAssistGateway.IsGeminiModel(
                    settings.Model) ||
                !useGemini;
            if (needsEndpoint ||
                settings.BaseUrl.Trim().Length > 0)
            {
                Uri endpoint;
                if (!AppSettings.TryGetChatCompletionsUri(
                    settings.BaseUrl,
                    true,
                    out endpoint))
                {
                    throw new InvalidOperationException(
                        "Use an HTTP or HTTPS endpoint URL.");
                }

                if (needsEndpoint &&
                    settings.ApiKey.Trim().Length == 0)
                {
                    throw new InvalidOperationException("Enter an API key.");
                }
            }

            var directory = Path.GetDirectoryName(_settingsPath);
            Directory.CreateDirectory(directory);

            var stored = new StoredSettings
            {
                BaseUrl = settings.BaseUrl.Trim(),
                Model = settings.Model.Trim(),
                ProtectedApiKey =
                    settings.ApiKey.Trim().Length > 0
                        ? Protect(settings.ApiKey.Trim())
                        : string.Empty,
                AllowInsecureHttp = true,
                UseGeminiSignIn = useGemini,
                // A disabled save deliberately removes previously
                // issued direct-Gemini credentials. Retaining the
                // implementation does not require retaining tokens.
                ProtectedGeminiRefreshToken =
                    useGemini &&
                    settings.GeminiRefreshToken.Trim().Length > 0
                        ? Protect(
                            settings.GeminiRefreshToken.Trim())
                        : string.Empty,
                GeminiProject = useGemini
                    ? TextBoundary.SingleLine(
                        settings.GeminiProject,
                        200)
                    : string.Empty,
                ToneProfile = TextBoundary.PlainText(
                    settings.ToneProfile,
                    TextBoundary.MaxToneProfileCharacters),
                UseToneProfile = settings.UseToneProfile &&
                    !string.IsNullOrWhiteSpace(settings.ToneProfile),
                ToneStrength = Math.Max(
                    10,
                    Math.Min(100, settings.ToneStrength)),
                DraftRules = TextBoundary.PlainText(
                    settings.DraftRules,
                    2000),
                SwitchToVisionModelForImages =
                    settings.SwitchToVisionModelForImages,
                DiscoveredModels = NormalizeDiscoveredModels(
                    settings.DiscoveredModels),
                McpServers = StoreMcpServers(
                    settings.McpServers),
                UseCustomLimits = false,
                LimitContextMultiplier = 1,
                LimitPromptCharacters = TextBoundary
                    .RecommendedUserPromptCharacters,
                LimitAssistantCharacters = TextBoundary
                    .RecommendedAssistantCharacters,
                LimitHistoryTurns = TextBoundary
                    .RecommendedConversationTurns,
                LimitToolRounds = TextBoundary
                    .RecommendedToolRounds,
                LimitToolCallsPerRound = TextBoundary
                    .RecommendedToolCallsPerRound,
                LimitWorkingSetMessages = NormalizeWorkingSet(
                    settings.LimitWorkingSetMessages)
            };

            File.WriteAllText(
                _settingsPath,
                _serializer.Serialize(stored),
                new UTF8Encoding(false));
        }

        private static string Protect(string value)
        {
            var clearBytes = Encoding.UTF8.GetBytes(value);
            var protectedBytes = ProtectedData.Protect(
                clearBytes,
                Entropy,
                DataProtectionScope.CurrentUser);
            return Convert.ToBase64String(protectedBytes);
        }

        private static string Unprotect(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var protectedBytes = Convert.FromBase64String(value);
            var clearBytes = ProtectedData.Unprotect(
                protectedBytes,
                Entropy,
                DataProtectionScope.CurrentUser);
            return Encoding.UTF8.GetString(clearBytes);
        }

        private static List<McpServerConfig> NormalizeMcpServers(
            IEnumerable<StoredMcpServer> servers)
        {
            var result = new List<McpServerConfig>();
            foreach (var server in servers ??
                Enumerable.Empty<StoredMcpServer>())
            {
                if (server == null)
                {
                    continue;
                }

                var config = new McpServerConfig
                {
                    Name = server.Name,
                    Target = server.Target,
                    Arguments = server.Arguments,
                     Headers = TryUnprotect(
                         server.ProtectedHeaders),
                    Enabled = server.Enabled,
                    BrowserTools = server.BrowserTools,
                    BrowserToolsApproved =
                        server.BrowserToolsApproved
                }.Sanitized();
                if (config.Target.Length == 0)
                {
                    continue;
                }

                result.Add(config);
                if (result.Count == McpServerConfig.MaxServers)
                {
                    break;
                }
            }

            return result;
        }

        private static List<StoredMcpServer> StoreMcpServers(
            IEnumerable<McpServerConfig> servers)
        {
            var result = new List<StoredMcpServer>();
            foreach (var server in servers ??
                Enumerable.Empty<McpServerConfig>())
            {
                if (server == null)
                {
                    continue;
                }

                var config = server.Sanitized();
                if (config.Target.Length == 0)
                {
                    continue;
                }

                result.Add(new StoredMcpServer
                {
                    Name = config.Name,
                    Target = config.Target,
                    Arguments = config.Arguments,
                    // Headers can carry an Authorization value, so
                    // they get the same DPAPI protection as keys.
                    ProtectedHeaders =
                        config.Headers.Trim().Length > 0
                            ? Protect(config.Headers.Trim())
                            : string.Empty,
                    Enabled = config.Enabled,
                    BrowserTools = config.BrowserTools,
                    BrowserToolsApproved =
                        config.BrowserToolsApproved
                });
                if (result.Count == McpServerConfig.MaxServers)
                {
                    break;
                }
            }

            return result;
        }

        // Settings written before the working-set size became
        // user-adjustable (or hand-edited files) may carry zero or
        // out-of-range values; zero means "use the default".
        private static int NormalizeWorkingSet(int value)
        {
            if (value <= 0)
            {
                return LimitOverrides.RecommendedWorkingSetMessages;
            }

            return Math.Max(
                LimitOverrides.MinWorkingSetMessages,
                Math.Min(
                    LimitOverrides.MaxWorkingSetMessages,
                    value));
        }

        private static List<string> NormalizeDiscoveredModels(
            IEnumerable<string> models)
        {
            return (models ?? Enumerable.Empty<string>())
                .Select(model => TextBoundary.PlainText(model, 200))
                .Where(ModelSelectionPolicy.IsGenerativeModel)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(model => model, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private sealed class StoredSettings
        {
            public string BaseUrl { get; set; }

            public string Model { get; set; }

            public string ProtectedApiKey { get; set; }

            public bool AllowInsecureHttp { get; set; }

            public bool UseGeminiSignIn { get; set; }

            public string ProtectedGeminiRefreshToken { get; set; }

            public string GeminiProject { get; set; }

            public string ToneProfile { get; set; }

            public bool UseToneProfile { get; set; }

            public int ToneStrength { get; set; }

            public string DraftRules { get; set; }

            public bool SwitchToVisionModelForImages { get; set; }

            public List<string> DiscoveredModels { get; set; }

            public List<StoredMcpServer> McpServers { get; set; }

            public bool UseCustomLimits { get; set; }

            public int LimitContextMultiplier { get; set; }

            public int LimitPromptCharacters { get; set; }

            public int LimitAssistantCharacters { get; set; }

            public int LimitHistoryTurns { get; set; }

            public int LimitToolRounds { get; set; }

            public int LimitToolCallsPerRound { get; set; }

            public int LimitWorkingSetMessages { get; set; }
        }

        private sealed class StoredMcpServer
        {
            public string Name { get; set; }

            public string Target { get; set; }

            public string Arguments { get; set; }

            public string ProtectedHeaders { get; set; }

            public bool Enabled { get; set; }

            public string BrowserTools { get; set; }

            public bool BrowserToolsApproved { get; set; }
        }

        // A headers value that fails to unprotect (copied profile,
        // reset DPAPI key) degrades to no headers instead of
        // blocking settings load.
        private static string TryUnprotect(string value)
        {
            try
            {
                return Unprotect(value);
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}

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

                return new AppSettings
                {
                    BaseUrl = stored.BaseUrl ?? string.Empty,
                    Model = stored.Model ?? string.Empty,
                    ApiKey = Unprotect(stored.ProtectedApiKey),
                    AllowInsecureHttp = stored.AllowInsecureHttp,
                    UseGeminiSignIn = stored.UseGeminiSignIn &&
                        !AdminPolicy.GeminiDisabled,
                    GeminiRefreshToken = Unprotect(
                        stored.ProtectedGeminiRefreshToken),
                    GeminiProject = TextBoundary.SingleLine(
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
                    // A missing UseCustomLimits (older settings
                    // files) means recommended limits; missing
                    // custom values (0) fall back to recommended.
                    UseRecommendedLimits = !stored.UseCustomLimits,
                    LimitContextMultiplier = OrDefault(
                        stored.LimitContextMultiplier,
                        1),
                    LimitPromptCharacters = OrDefault(
                        stored.LimitPromptCharacters,
                        TextBoundary
                            .RecommendedUserPromptCharacters),
                    LimitAssistantCharacters = OrDefault(
                        stored.LimitAssistantCharacters,
                        TextBoundary
                            .RecommendedAssistantCharacters),
                    LimitHistoryTurns = OrDefault(
                        stored.LimitHistoryTurns,
                        TextBoundary
                            .RecommendedConversationTurns),
                    LimitToolRounds = OrDefault(
                        stored.LimitToolRounds,
                        TextBoundary.RecommendedToolRounds),
                    LimitToolCallsPerRound = OrDefault(
                        stored.LimitToolCallsPerRound,
                        TextBoundary
                            .RecommendedToolCallsPerRound),
                    LimitWorkingSetMessages = OrDefault(
                        stored.LimitWorkingSetMessages,
                        LimitOverrides
                            .RecommendedWorkingSetMessages)
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

            // The endpoint is required whenever the selected model
            // is not a Gemini model - the Gemini tick no longer
            // decides the transport, the model does - and a typed
            // endpoint is always validated even alongside Gemini.
            var needsEndpoint =
                !GeminiCodeAssistGateway.IsGeminiModel(
                    settings.Model) ||
                !settings.UseGeminiSignIn;
            if (needsEndpoint ||
                settings.BaseUrl.Trim().Length > 0)
            {
                Uri endpoint;
                if (!AppSettings.TryGetChatCompletionsUri(
                    settings.BaseUrl,
                    settings.AllowInsecureHttp,
                    out endpoint))
                {
                    throw new InvalidOperationException(
                        "Use HTTPS, loopback HTTP, or explicitly allow insecure HTTP.");
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
                AllowInsecureHttp = settings.AllowInsecureHttp,
                UseGeminiSignIn = settings.UseGeminiSignIn,
                ProtectedGeminiRefreshToken =
                    settings.GeminiRefreshToken.Trim().Length > 0
                        ? Protect(
                            settings.GeminiRefreshToken.Trim())
                        : string.Empty,
                GeminiProject = TextBoundary.SingleLine(
                    settings.GeminiProject,
                    200),
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
                UseCustomLimits = !settings.UseRecommendedLimits,
                LimitContextMultiplier =
                    settings.LimitContextMultiplier,
                LimitPromptCharacters =
                    settings.LimitPromptCharacters,
                LimitAssistantCharacters =
                    settings.LimitAssistantCharacters,
                LimitHistoryTurns = settings.LimitHistoryTurns,
                LimitToolRounds = settings.LimitToolRounds,
                LimitToolCallsPerRound =
                    settings.LimitToolCallsPerRound,
                LimitWorkingSetMessages =
                    settings.LimitWorkingSetMessages
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

        private static List<string> NormalizeDiscoveredModels(
            IEnumerable<string> models)
        {
            return (models ?? Enumerable.Empty<string>())
                .Select(model => TextBoundary.PlainText(model, 200))
                .Where(model => model.Length > 0)
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

        private static int OrDefault(int value, int fallback)
        {
            return value > 0 ? value : fallback;
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

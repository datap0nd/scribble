using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Scribble.Configuration;
using Scribble.Security;

namespace Scribble.Chat
{
    // Gemini over Google sign-in, end to end inside Scribble. The user
    // clicks Sign in with Google in Settings (GoogleSignInFlow runs
    // the browser OAuth flow); the refresh token is stored
    // DPAPI-protected in Scribble's own settings and exchanged here
    // for short-lived access tokens. Requests go to the same Code
    // Assist generateContent API the Gemini CLI uses, so enterprise
    // Gemini licensing resolves server-side from the account alone -
    // no API key and no cloud project setup. An existing Gemini CLI
    // session on the machine is honored as a silent fallback. The
    // OpenAI-shaped requests the rest of the app builds are
    // translated to Gemini's native format and back, so every
    // existing guardrail (read-only mailbox, one-shot drafts, no
    // send capability) applies unchanged.
    public sealed class GeminiCliCredentials
    {
        public GeminiCliCredentials(
            string accessToken,
            string refreshToken,
            long expiryUtcMilliseconds)
        {
            AccessToken = accessToken ?? string.Empty;
            RefreshToken = refreshToken ?? string.Empty;
            ExpiryUtcMilliseconds = expiryUtcMilliseconds;
        }

        public string AccessToken { get; }

        public string RefreshToken { get; }

        public long ExpiryUtcMilliseconds { get; }
    }

    public sealed class GeminiCodeAssistGateway
    {
        public const string ApiBase =
            "https://cloudcode-pa.googleapis.com/v1internal";
        // Access tokens are refreshed this long before their expiry.
        private const long RefreshMarginMilliseconds = 120000;

        // The model set the current Gemini CLI ships (the Code
        // Assist API has no listing endpoint). Quota buckets are
        // per model, so a model at capacity does not block the
        // others. Preview models are tier-gated: if the account
        // lacks access, the request fails with a permission error
        // and another model can be picked.
        public static readonly IReadOnlyList<string> KnownModels =
            new[]
            {
                "gemini-3.5-flash",
                "gemini-3-flash",
                "gemini-2.5-flash",
                "gemini-3.1-flash-lite",
                "gemini-2.5-flash-lite",
                "gemini-3.1-pro-preview",
                "gemini-3-pro-preview",
                "gemini-2.5-pro"
            };

        // The single source of truth for "is this a Google model".
        // Transport is chosen from the SELECTED MODEL, never from a
        // mode switch, so the model picker is always the decider: a
        // gemini-* id goes to the Gemini gateway and every other id
        // goes to the configured OpenAI-compatible endpoint.
        public static bool IsGeminiModel(string model)
        {
            var name = (model ?? string.Empty).Trim();
            if (name.Length == 0)
            {
                return false;
            }

            if (name.StartsWith(
                "gemini",
                StringComparison.OrdinalIgnoreCase) ||
                name.StartsWith(
                    "models/gemini",
                    StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            foreach (var known in KnownModels)
            {
                if (string.Equals(
                    known,
                    name,
                    StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private readonly JavaScriptSerializer _serializer =
            new JavaScriptSerializer
            {
                MaxJsonLength = int.MaxValue
            };

        private string _cachedAccessToken = string.Empty;
        private long _cachedTokenExpiryMs;
        private string _cachedProject;
        // Thought signatures captured from functionCall parts, keyed
        // by the tool-call id they were issued under; echoed back
        // when history replays those calls. Signatures only matter
        // within a live tool loop, so the map is cleared once it
        // outgrows any plausible loop.
        private readonly Dictionary<string, string>
            _thoughtSignatures =
                new Dictionary<string, string>(
                    StringComparer.Ordinal);

        // Optional UI hook: called with a short human-readable note
        // when the gateway is silently waiting (e.g. a quota reset)
        // so slowness is never a mystery.
        public Action<string> StatusListener { get; set; }

        // Retry policy mirroring the Gemini CLI: up to ten attempts
        // per request, exponential backoff (5s doubling to a 30s
        // cap) with jitter, the server's retry hint honored when
        // longer, and a persistent-429 fallback from pro to flash
        // that sticks for the rest of the session.
        public const int MaxRetryAttempts = 10;
        // Every known model is a separate quota bucket, and hopping
        // costs only one HTTP round trip, so on a capacity error the
        // request hops the whole family instantly - fast models
        // first, pro-class last - and only waits when every bucket
        // is dry. Tier-blocked models simply answer 429 and cost a
        // fraction of a second to skip.
        public static readonly IReadOnlyList<string>
            CapacityFallbackChain =
                new[]
                {
                    "gemini-3.5-flash",
                    "gemini-3-flash",
                    "gemini-2.5-flash",
                    "gemini-3.1-flash-lite",
                    "gemini-2.5-flash-lite",
                    "gemini-3.1-pro-preview",
                    "gemini-3-pro-preview",
                    "gemini-2.5-pro"
                };
        private readonly Random _retryJitter = new Random();
        // Session-sticky reroutes: once a model exhausts, later
        // requests for it start directly at the model that worked.
        private readonly Dictionary<string, string>
            _stickyModelRoutes =
                new Dictionary<string, string>(
                    StringComparer.OrdinalIgnoreCase);

        public static int ComputeRetryDelaySeconds(
            int attempt,
            string body,
            double jitter01)
        {
            // The CLI's interactive capacity choreography: two fast
            // probes (1s, 3s) before the model fallback decision,
            // then exponential backoff capped at 30s, always
            // honoring a longer server hint up to 90s.
            if (attempt == 0)
            {
                return 1;
            }

            if (attempt == 1)
            {
                return 3;
            }

            var backoff = Math.Min(30, 5 << Math.Min(attempt, 3));
            var hint = ParseRetryDelaySeconds(body);
            var delay = Math.Min(90, Math.Max(backoff, hint));
            return delay +
                (int)(delay * 0.2 *
                      Math.Max(0.0, Math.Min(1.0, jitter01)));
        }

        // Next model in the capacity chain that has not been tried
        // yet this request. Any Gemini model participates; a null
        // result means every bucket in the chain was exhausted.
        public static string NextFallbackModel(
            string current,
            ICollection<string> triedModels)
        {
            var name = NormalizeModel(current).ToLowerInvariant();
            if (name.IndexOf(
                    "gemini",
                    StringComparison.Ordinal) < 0)
            {
                return null;
            }

            foreach (var candidate in CapacityFallbackChain)
            {
                if (!string.Equals(
                        candidate,
                        name,
                        StringComparison.OrdinalIgnoreCase) &&
                    (triedModels == null ||
                     !triedModels.Contains(candidate)))
                {
                    return candidate;
                }
            }

            return null;
        }

        // The thinking budget is model-family specific, so a hop
        // must rewrite it: 2.5 models get their explicit budget and
        // other families must not receive the 2.5-only field.
        public static void ApplyThinkingConfig(
            IDictionary<string, object> envelope,
            string model)
        {
            object requestValue;
            if (!envelope.TryGetValue(
                    "request",
                    out requestValue))
            {
                return;
            }

            var request = requestValue
                as IDictionary<string, object>;
            if (request == null)
            {
                return;
            }

            object configValue;
            var generationConfig = request.TryGetValue(
                "generationConfig",
                out configValue)
                ? configValue as IDictionary<string, object>
                : null;
            var budget = ThinkingBudgetFor(model);
            if (budget < 0)
            {
                if (generationConfig != null)
                {
                    generationConfig.Remove("thinkingConfig");
                    if (generationConfig.Count == 0)
                    {
                        request.Remove("generationConfig");
                    }
                }

                return;
            }

            if (generationConfig == null)
            {
                generationConfig =
                    new Dictionary<string, object>();
                request["generationConfig"] = generationConfig;
            }

            generationConfig["thinkingConfig"] =
                new Dictionary<string, object>
                {
                    { "thinkingBudget", budget }
                };
        }

        private string EffectiveModel(string model)
        {
            var normalized = NormalizeModel(model);
            string route;
            if (_stickyModelRoutes.TryGetValue(
                normalized,
                out route))
            {
                return route;
            }

            return normalized;
        }

        // Central 429 policy for generateContent. Returns 1 when the
        // caller should retry after a wait, 2 when the request
        // hopped to the next model in the chain (caller restarts its
        // attempt counter), and 0 when everything is exhausted. Hops
        // happen immediately - waiting is reserved for the case
        // where the entire chain is dry, and each waited cycle
        // re-probes the whole chain.
        private async Task<int> HandleCapacityAsync(
            IDictionary<string, object> envelope,
            int attempt,
            string body,
            string originalModel,
            HashSet<string> triedModels,
            CancellationToken cancellationToken)
        {
            var current = Convert.ToString(envelope["model"]);
            triedModels.Add(NormalizeModel(current)
                .ToLowerInvariant());
            var next = NextFallbackModel(current, triedModels);
            if (next != null)
            {
                envelope["model"] = next;
                ApplyThinkingConfig(envelope, next);
                _stickyModelRoutes[
                    NormalizeModel(originalModel)] = next;
                StatusListener?.Invoke(
                    current + " is at capacity - trying " + next);
                return 2;
            }

            if (attempt >= MaxRetryAttempts - 1)
            {
                return 0;
            }

            // Whole chain dry: wait, then re-probe every bucket.
            triedModels.Clear();
            var delay = ComputeRetryDelaySeconds(
                attempt,
                body,
                _retryJitter.NextDouble());
            StatusListener?.Invoke(
                "All Gemini models are at capacity - retry " +
                (attempt + 1) + " of " + MaxRetryAttempts +
                " in " + delay + "s");
            await Task.Delay(
                TimeSpan.FromSeconds(delay),
                cancellationToken).ConfigureAwait(true);
            return 1;
        }

        public static string CredentialsPath
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(
                        Environment.SpecialFolder.UserProfile),
                    ".gemini",
                    "oauth_creds.json");
            }
        }

        // ------------------------------------------------------------------
        // Credentials.
        // ------------------------------------------------------------------

        public static GeminiCliCredentials ParseCredentials(
            string json)
        {
            try
            {
                var serializer = new JavaScriptSerializer();
                var map = serializer.DeserializeObject(json)
                    as IDictionary<string, object>;
                if (map == null)
                {
                    return null;
                }

                object accessValue;
                object refreshValue;
                object expiryValue;
                map.TryGetValue("access_token", out accessValue);
                map.TryGetValue("refresh_token", out refreshValue);
                map.TryGetValue("expiry_date", out expiryValue);
                var refresh =
                    Convert.ToString(refreshValue) ?? string.Empty;
                if (refresh.Trim().Length == 0)
                {
                    return null;
                }

                long expiry = 0;
                if (expiryValue != null)
                {
                    long.TryParse(
                        Convert.ToString(
                            expiryValue,
                            System.Globalization.CultureInfo
                                .InvariantCulture),
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo
                            .InvariantCulture,
                        out expiry);
                }

                return new GeminiCliCredentials(
                    Convert.ToString(accessValue) ?? string.Empty,
                    refresh,
                    expiry);
            }
            catch
            {
                return null;
            }
        }

        public static bool NeedsRefresh(
            GeminiCliCredentials credentials,
            long nowUtcMilliseconds)
        {
            return credentials.AccessToken.Trim().Length == 0 ||
                   credentials.ExpiryUtcMilliseconds -
                   nowUtcMilliseconds < RefreshMarginMilliseconds;
        }

        // Primes the in-memory token cache right after a fresh
        // browser sign-in so the first request needs no refresh.
        public void PrimeAccessToken(
            string accessToken,
            long expiresInSeconds)
        {
            _cachedAccessToken = accessToken ?? string.Empty;
            _cachedTokenExpiryMs = NowUtcMilliseconds() +
                expiresInSeconds * 1000;
        }

        private async Task<string> GetAccessTokenAsync(
            HttpClient httpClient,
            AppSettings settings,
            CancellationToken cancellationToken)
        {
            EnsureGeminiAllowed();
            var nowMs = NowUtcMilliseconds();
            if (_cachedAccessToken.Length > 0 &&
                _cachedTokenExpiryMs - nowMs >
                RefreshMarginMilliseconds)
            {
                return _cachedAccessToken;
            }

            // Scribble's own browser sign-in is the primary source;
            // an existing Gemini CLI session on the same machine
            // works as a silent fallback.
            var refreshToken =
                (settings?.GeminiRefreshToken ?? string.Empty)
                .Trim();
            if (refreshToken.Length == 0)
            {
                var path = CredentialsPath;
                if (File.Exists(path))
                {
                    var credentials = ParseCredentials(
                        File.ReadAllText(path, Encoding.UTF8));
                    if (credentials != null)
                    {
                        if (!NeedsRefresh(credentials, nowMs))
                        {
                            _cachedAccessToken =
                                credentials.AccessToken;
                            _cachedTokenExpiryMs =
                                credentials.ExpiryUtcMilliseconds;
                            return _cachedAccessToken;
                        }

                        refreshToken = credentials.RefreshToken;
                    }
                }
            }

            if (refreshToken.Length == 0)
            {
                throw new AiEndpointException(
                    "GEMINI_SIGNIN_MISSING",
                    "No Google sign-in was found. Open Scribble " +
                    "Settings and click Sign in with Google.");
            }

            return await RefreshAccessTokenAsync(
                httpClient,
                refreshToken,
                cancellationToken).ConfigureAwait(true);
        }

        private async Task<string> RefreshAccessTokenAsync(
            HttpClient httpClient,
            string refreshToken,
            CancellationToken cancellationToken)
        {
            using (var request = new HttpRequestMessage(
                HttpMethod.Post,
                GoogleSignInFlow.TokenEndpoint))
            {
                request.Content = new FormUrlEncodedContent(
                    new Dictionary<string, string>
                    {
                        {
                            "client_id",
                            GoogleSignInFlow.OAuthClientId
                        },
                        {
                            "client_secret",
                            GoogleSignInFlow.OAuthClientSecret
                        },
                        { "refresh_token", refreshToken },
                        { "grant_type", "refresh_token" }
                    });
                using (var response = await httpClient
                    .SendAsync(request, cancellationToken)
                    .ConfigureAwait(true))
                {
                    var body = await ReadBodyAsync(
                        response,
                        cancellationToken).ConfigureAwait(true);
                    if (!response.IsSuccessStatusCode)
                    {
                        throw new AiEndpointException(
                            "GEMINI_SIGNIN_EXPIRED",
                            "The Google sign-in could not be " +
                            "refreshed (" +
                            (int)response.StatusCode +
                            "). Open Scribble Settings and click " +
                            "Sign in with Google again.",
                            responseSnippet: body);
                    }

                    var map = _serializer.DeserializeObject(body)
                        as IDictionary<string, object>;
                    object tokenValue;
                    object expiresValue;
                    if (map == null ||
                        !map.TryGetValue(
                            "access_token",
                            out tokenValue))
                    {
                        throw new AiEndpointException(
                            "GEMINI_SIGNIN_EXPIRED",
                            "Google returned no access token. " +
                            "Open Scribble Settings and click Sign " +
                            "in with Google again.",
                            responseSnippet: body);
                    }

                    map.TryGetValue("expires_in", out expiresValue);
                    long expiresSeconds = 3600;
                    if (expiresValue != null)
                    {
                        long.TryParse(
                            Convert.ToString(
                                expiresValue,
                                System.Globalization.CultureInfo
                                    .InvariantCulture),
                            out expiresSeconds);
                    }

                    _cachedAccessToken =
                        Convert.ToString(tokenValue) ??
                        string.Empty;
                    _cachedTokenExpiryMs = NowUtcMilliseconds() +
                        expiresSeconds * 1000;
                    return _cachedAccessToken;
                }
            }
        }

        // ------------------------------------------------------------------
        // Project discovery (loadCodeAssist / onboardUser).
        // ------------------------------------------------------------------

        // Mirrors the Gemini CLI's project resolution: an already
        // onboarded account (currentTier present) uses the project
        // from loadCodeAssist or the GOOGLE_CLOUD_PROJECT
        // environment variable (enterprise tiers designate one, the
        // same variable Gemini CLI uses); a new account is onboarded
        // once and its long-running operation is polled with GET.
        // generateContent is never called with an empty project -
        // that is exactly what Google answers with an opaque
        // HTTP 500 "Internal error encountered".
        private async Task<string> GetProjectAsync(
            HttpClient httpClient,
            AppSettings settings,
            CancellationToken cancellationToken)
        {
            if (_cachedProject != null)
            {
                return _cachedProject;
            }

            // The Settings field wins over the environment variable
            // so colleagues can be set up without touching their
            // machine configuration.
            var environmentProject =
                (settings?.GeminiProject ?? string.Empty).Trim();
            if (environmentProject.Length == 0)
            {
                environmentProject =
                    (Environment.GetEnvironmentVariable(
                         "GOOGLE_CLOUD_PROJECT") ?? string.Empty)
                    .Trim();
            }
            var loadBody = new Dictionary<string, object>
            {
                { "metadata", ClientMetadata(environmentProject) }
            };
            if (environmentProject.Length > 0)
            {
                loadBody["cloudaicompanionProject"] =
                    environmentProject;
            }

            var load = await PostJsonAsync(
                httpClient,
                settings,
                ":loadCodeAssist",
                loadBody,
                cancellationToken).ConfigureAwait(true);
            var loadProject = ReadString(
                load,
                "cloudaicompanionProject");
            object tierValue = null;
            load?.TryGetValue("currentTier", out tierValue);
            if (tierValue is IDictionary<string, object>)
            {
                var resolved = loadProject.Length > 0
                    ? loadProject
                    : environmentProject;
                if (resolved.Length == 0)
                {
                    throw ProjectUnresolved();
                }

                _cachedProject = resolved;
                return resolved;
            }

            var tierId = FindDefaultTierId(load);
            var onboardBody = new Dictionary<string, object>
            {
                { "tierId", tierId },
                { "metadata", ClientMetadata(environmentProject) }
            };
            if (environmentProject.Length > 0)
            {
                onboardBody["cloudaicompanionProject"] =
                    environmentProject;
            }

            var operation = await PostJsonAsync(
                httpClient,
                settings,
                ":onboardUser",
                onboardBody,
                cancellationToken).ConfigureAwait(true);
            for (var attempt = 0;
                 attempt < 15 &&
                 operation != null &&
                 !ReadBool(operation, "done");
                 attempt++)
            {
                await Task.Delay(2000, cancellationToken)
                    .ConfigureAwait(true);
                var name = ReadString(operation, "name");
                if (name.Length == 0)
                {
                    break;
                }

                operation = await GetJsonAsync(
                    httpClient,
                    settings,
                    "/" + name,
                    cancellationToken).ConfigureAwait(true);
            }

            var onboardedProject = string.Empty;
            if (operation != null && ReadBool(operation, "done"))
            {
                object responseValue;
                operation.TryGetValue(
                    "response",
                    out responseValue);
                var responseMap = responseValue
                    as IDictionary<string, object>;
                object projectValue = null;
                responseMap?.TryGetValue(
                    "cloudaicompanionProject",
                    out projectValue);
                // The project arrives as {"id": ...} on some tiers
                // and as a plain string on others.
                var projectMap = projectValue
                    as IDictionary<string, object>;
                onboardedProject = projectMap != null
                    ? ReadString(projectMap, "id")
                    : (projectValue as string ?? string.Empty);
            }

            var final = onboardedProject.Length > 0
                ? onboardedProject
                : environmentProject;
            if (final.Length == 0)
            {
                throw ProjectUnresolved();
            }

            _cachedProject = final;
            return final;
        }

        private static AiEndpointException ProjectUnresolved()
        {
            return new AiEndpointException(
                "GEMINI_PROJECT_UNRESOLVED",
                "Google did not provide a Gemini project for this " +
                "account. Enter your organization's Google Cloud " +
                "project id in Scribble Settings (the Google Cloud " +
                "project field next to the Gemini sign-in - the " +
                "same id for everyone in the organization, ask " +
                "your admin or a colleague where it works), then " +
                "sign in and try again.");
        }

        private static string FindDefaultTierId(
            IDictionary<string, object> load)
        {
            object tiersValue = null;
            load?.TryGetValue("allowedTiers", out tiersValue);
            var tiers = tiersValue as object[];
            if (tiers != null)
            {
                foreach (var entry in tiers)
                {
                    var tier = entry as IDictionary<string, object>;
                    if (tier == null)
                    {
                        continue;
                    }

                    object isDefaultValue;
                    tier.TryGetValue(
                        "isDefault",
                        out isDefaultValue);
                    if (isDefaultValue is bool &&
                        (bool)isDefaultValue)
                    {
                        var id = ReadString(tier, "id");
                        if (id.Length > 0)
                        {
                            return id;
                        }
                    }
                }
            }

            return "free-tier";
        }

        private static Dictionary<string, object> ClientMetadata(
            string duetProject)
        {
            var metadata = new Dictionary<string, object>
            {
                { "ideType", "IDE_UNSPECIFIED" },
                { "platform", "PLATFORM_UNSPECIFIED" },
                { "pluginType", "GEMINI" }
            };
            if (!string.IsNullOrEmpty(duetProject))
            {
                metadata["duetProject"] = duetProject;
            }

            return metadata;
        }

        private static bool ReadBool(
            IDictionary<string, object> map,
            string key)
        {
            object value;
            return map != null &&
                   map.TryGetValue(key, out value) &&
                   value is bool &&
                   (bool)value;
        }

        // ------------------------------------------------------------------
        // Public entry points.
        // ------------------------------------------------------------------

        private static void EnsureGeminiAllowed()
        {
            if (!AdminPolicy.GeminiDisabled)
            {
                return;
            }

            throw new AiEndpointException(
                "GEMINI_DISABLED_BY_POLICY",
                "Google Gemini is unavailable in this build. Use " +
                "an OpenAI-compatible endpoint in Settings.");
        }

        public async Task<IReadOnlyList<string>> VerifySignInAsync(
            HttpClient httpClient,
            AppSettings settings,
            CancellationToken cancellationToken)
        {
            EnsureGeminiAllowed();
            await GetAccessTokenAsync(
                httpClient,
                settings,
                cancellationToken).ConfigureAwait(true);
            await GetProjectAsync(
                httpClient,
                settings,
                cancellationToken).ConfigureAwait(true);
            return KnownModels;
        }

        public async Task<ChatCompletionResponseMessage>
            GenerateAsync(
                HttpClient httpClient,
                AppSettings settings,
                ChatCompletionRequest requestModel,
                CancellationToken cancellationToken)
        {
            EnsureGeminiAllowed();

            var project = await GetProjectAsync(
                httpClient,
                settings,
                cancellationToken).ConfigureAwait(true);
            var envelope = new Dictionary<string, object>
            {
                { "model", EffectiveModel(requestModel.model) },
                { "project", project },
                {
                    "request",
                    TranslateRequest(
                        requestModel,
                        _thoughtSignatures)
                }
            };
            ApplyThinkingConfig(
                envelope,
                Convert.ToString(envelope["model"]));
            IDictionary<string, object> root = null;
            var triedModels = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            for (var attempt = 0;
                 attempt < MaxRetryAttempts;
                 attempt++)
            {
                try
                {
                    root = await PostJsonAsync(
                        httpClient,
                        settings,
                        ":generateContent",
                        envelope,
                        cancellationToken).ConfigureAwait(true);
                    break;
                }
                catch (AiEndpointException exception)
                    when (exception.HttpStatus == 429)
                {
                    var action = await HandleCapacityAsync(
                        envelope,
                        attempt,
                        exception.ResponseSnippet,
                        requestModel.model,
                        triedModels,
                        cancellationToken).ConfigureAwait(true);
                    if (action == 0)
                    {
                        throw;
                    }

                    if (action == 2)
                    {
                        attempt = -1;
                    }
                }
            }

            object responseValue = null;
            root?.TryGetValue("response", out responseValue);
            if (_thoughtSignatures.Count > 500)
            {
                _thoughtSignatures.Clear();
            }

            var message = TranslateResponse(
                responseValue as IDictionary<string, object> ??
                root,
                _thoughtSignatures);
            if (message == null)
            {
                throw new AiEndpointException(
                    "RESPONSE_MISSING_CONTENT",
                    "The Gemini endpoint returned neither message " +
                    "text nor tool calls.");
            }

            return message;
        }

        // Streams a generateContent call over SSE so text reaches
        // the chat as it is produced instead of after the full
        // response. Function calls and thought signatures are
        // collected across chunks; thought parts are never streamed.
        public async Task<ChatCompletionResponseMessage>
            GenerateStreamAsync(
                HttpClient httpClient,
                AppSettings settings,
                ChatCompletionRequest requestModel,
                Action<string> onTextDelta,
                CancellationToken cancellationToken)
        {
            EnsureGeminiAllowed();
            var project = await GetProjectAsync(
                httpClient,
                settings,
                cancellationToken).ConfigureAwait(true);
            var envelope = new Dictionary<string, object>
            {
                { "model", EffectiveModel(requestModel.model) },
                { "project", project },
                {
                    "request",
                    TranslateRequest(
                        requestModel,
                        _thoughtSignatures)
                }
            };
            ApplyThinkingConfig(
                envelope,
                Convert.ToString(envelope["model"]));
            var triedModels = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            for (var attempt = 0;
                 attempt < MaxRetryAttempts;
                 attempt++)
            {
                var json = _serializer.Serialize(envelope);
                var token = await GetAccessTokenAsync(
                    httpClient,
                    settings,
                    cancellationToken).ConfigureAwait(true);
                using (var request = new HttpRequestMessage(
                    HttpMethod.Post,
                    ApiBase + ":streamGenerateContent?alt=sse"))
                {
                    request.Headers.Authorization =
                        new AuthenticationHeaderValue(
                            "Bearer",
                            token);
                    request.Headers.Accept.Add(
                        new MediaTypeWithQualityHeaderValue(
                            "text/event-stream"));
                    request.Content = new StringContent(
                        json,
                        Encoding.UTF8,
                        "application/json");
                    using (var response = await httpClient
                        .SendAsync(
                            request,
                            HttpCompletionOption
                                .ResponseHeadersRead,
                            cancellationToken)
                        .ConfigureAwait(true))
                    {
                        if (!response.IsSuccessStatusCode)
                        {
                            var body = await ReadBodyAsync(
                                response,
                                cancellationToken)
                                .ConfigureAwait(true);
                            var status =
                                (int)response.StatusCode;
                            if (status == 429)
                            {
                                var action =
                                    await HandleCapacityAsync(
                                        envelope,
                                        attempt,
                                        body,
                                        requestModel.model,
                                        triedModels,
                                        cancellationToken)
                                        .ConfigureAwait(true);
                                if (action == 2)
                                {
                                    attempt = -1;
                                }

                                if (action != 0)
                                {
                                    continue;
                                }

                                throw new AiEndpointException(
                                    "GEMINI_RATE_LIMITED",
                                    "The Gemini quota for this " +
                                    "model stayed exhausted " +
                                    "through " + MaxRetryAttempts +
                                    " retries. Wait a few minutes " +
                                    "and try again, or switch to " +
                                    "gemini-2.5-flash.",
                                    httpStatus: status,
                                    responseSnippet: body);
                            }

                            var hint =
                                status == 401 || status == 403
                                    ? " The Google sign-in may " +
                                      "have expired - open " +
                                      "Scribble Settings and click " +
                                      "Sign in with Google again."
                                    : string.Empty;
                            throw new AiEndpointException(
                                "GEMINI_HTTP_" + status,
                                "The Gemini endpoint rejected " +
                                "streamGenerateContent: " +
                                status + "." + hint,
                                httpStatus: status,
                                responseSnippet: body);
                        }

                        if (_thoughtSignatures.Count > 500)
                        {
                            _thoughtSignatures.Clear();
                        }

                        var message = await ReadSseAsync(
                            response,
                            onTextDelta,
                            cancellationToken)
                            .ConfigureAwait(true);
                        if (message == null)
                        {
                            throw new AiEndpointException(
                                "RESPONSE_MISSING_CONTENT",
                                "The Gemini endpoint returned " +
                                "neither message text nor tool " +
                                "calls.");
                        }

                        return message;
                    }
                }
            }

            throw new AiEndpointException(
                "GEMINI_RATE_LIMITED",
                "The Gemini quota stayed exhausted through " +
                MaxRetryAttempts + " retries. Wait a few minutes " +
                "and try again.");
        }

        private async Task<ChatCompletionResponseMessage>
            ReadSseAsync(
                HttpResponseMessage response,
                Action<string> onTextDelta,
                CancellationToken cancellationToken)
        {
            var text = new StringBuilder();
            var toolCalls = new List<ChatToolCall>();
            var rawCap =
                TextBoundary.MaxAssistantCharacters * 4;
            using (var stream = await response.Content
                .ReadAsStreamAsync().ConfigureAwait(true))
            // A stalled connection would block ReadLineAsync past
            // the between-chunk cancellation checks; closing the
            // stream on cancel faults the pending read.
            using (cancellationToken.Register(stream.Close))
            using (var reader = new StreamReader(
                stream,
                Encoding.UTF8))
            {
                var dataLines = new List<string>();
                while (true)
                {
                    cancellationToken
                        .ThrowIfCancellationRequested();
                    string line;
                    try
                    {
                        line = await reader.ReadLineAsync()
                            .ConfigureAwait(true);
                    }
                    catch (Exception exception)
                        when (cancellationToken
                            .IsCancellationRequested)
                    {
                        throw new OperationCanceledException(
                            "The stream read was cancelled.",
                            exception,
                            cancellationToken);
                    }

                    if (line == null)
                    {
                        ConsumeSseChunk(
                            dataLines,
                            text,
                            toolCalls,
                            onTextDelta,
                            rawCap);
                        break;
                    }

                    if (line.StartsWith(
                        "data:",
                        StringComparison.Ordinal))
                    {
                        dataLines.Add(
                            line.Substring(5).Trim());
                        continue;
                    }

                    if (line.Length == 0)
                    {
                        ConsumeSseChunk(
                            dataLines,
                            text,
                            toolCalls,
                            onTextDelta,
                            rawCap);
                    }
                }
            }

            var boundedText = TextBoundary.PlainText(
                text.ToString(),
                TextBoundary.MaxAssistantCharacters);
            if (boundedText.Length == 0 && toolCalls.Count == 0)
            {
                return null;
            }

            return new ChatCompletionResponseMessage
            {
                role = "assistant",
                content = boundedText,
                tool_calls = toolCalls.Count > 0
                    ? toolCalls
                    : null
            };
        }

        private void ConsumeSseChunk(
            List<string> dataLines,
            StringBuilder text,
            List<ChatToolCall> toolCalls,
            Action<string> onTextDelta,
            int rawCap)
        {
            if (dataLines.Count == 0)
            {
                return;
            }

            var payload = string.Join("\n", dataLines);
            dataLines.Clear();
            IDictionary<string, object> root;
            try
            {
                root = _serializer.DeserializeObject(payload)
                    as IDictionary<string, object>;
            }
            catch
            {
                return;
            }

            if (root == null)
            {
                return;
            }

            object responseValue;
            var inner = root.TryGetValue(
                "response",
                out responseValue)
                ? responseValue as IDictionary<string, object>
                : root;
            object candidatesValue = null;
            inner?.TryGetValue(
                "candidates",
                out candidatesValue);
            var candidates = candidatesValue as object[];
            var candidate = candidates != null &&
                candidates.Length > 0
                ? candidates[0] as IDictionary<string, object>
                : null;
            object contentValue = null;
            candidate?.TryGetValue("content", out contentValue);
            var content = contentValue
                as IDictionary<string, object>;
            object partsValue = null;
            content?.TryGetValue("parts", out partsValue);
            var parts = partsValue as object[];
            if (parts == null)
            {
                return;
            }

            foreach (var entry in parts)
            {
                var part = entry as IDictionary<string, object>;
                if (part == null)
                {
                    continue;
                }

                object thoughtValue;
                if (part.TryGetValue(
                        "thought",
                        out thoughtValue) &&
                    thoughtValue is bool &&
                    (bool)thoughtValue)
                {
                    continue;
                }

                object textValue;
                if (part.TryGetValue("text", out textValue) &&
                    textValue is string)
                {
                    var chunkText = (string)textValue;
                    if (chunkText.Length > 0 &&
                        text.Length < rawCap)
                    {
                        text.Append(chunkText);
                        onTextDelta?.Invoke(chunkText);
                    }

                    continue;
                }

                object callValue;
                if (part.TryGetValue(
                        "functionCall",
                        out callValue))
                {
                    var call = callValue
                        as IDictionary<string, object>;
                    if (call == null)
                    {
                        continue;
                    }

                    object argsValue;
                    call.TryGetValue("args", out argsValue);
                    var callId = "gemini_call_" +
                        Guid.NewGuid().ToString("N")
                            .Substring(0, 12);
                    var signature = ReadString(
                        part,
                        "thoughtSignature");
                    if (signature.Length > 0)
                    {
                        _thoughtSignatures[callId] = signature;
                    }

                    toolCalls.Add(new ChatToolCall
                    {
                        id = callId,
                        type = "function",
                        function = new ChatToolCallFunction
                        {
                            name = ReadString(call, "name"),
                            arguments = _serializer.Serialize(
                                argsValue ??
                                new Dictionary<string, object>())
                        }
                    });
                }
            }
        }

        // ------------------------------------------------------------------
        // Request translation (OpenAI chat shape to Gemini native).
        // Public and pure so the guardrail tests can cover it.
        // ------------------------------------------------------------------

        public static string NormalizeModel(string model)
        {
            var value = (model ?? string.Empty).Trim();
            const string prefix = "models/";
            return value.StartsWith(
                prefix,
                StringComparison.OrdinalIgnoreCase)
                ? value.Substring(prefix.Length)
                : value;
        }

        public static Dictionary<string, object> TranslateRequest(
            ChatCompletionRequest request)
        {
            return TranslateRequest(request, null);
        }

        // Gemini 2.5 attaches an opaque thoughtSignature to every
        // functionCall part it returns and rejects follow-up
        // requests that replay the call without it, so signatures
        // captured from responses are echoed back here by tool-call
        // id.
        public static Dictionary<string, object> TranslateRequest(
            ChatCompletionRequest request,
            IDictionary<string, string> thoughtSignatures)
        {
            var contents = new List<object>();
            var systemText = new StringBuilder();
            var callNames =
                new Dictionary<string, string>(
                    StringComparer.Ordinal);
            foreach (var entry in request.messages ??
                new List<object>())
            {
                var input = entry as ChatCompletionInputMessage;
                if (input != null)
                {
                    if (string.Equals(
                        input.role,
                        "system",
                        StringComparison.OrdinalIgnoreCase))
                    {
                        if (systemText.Length > 0)
                        {
                            systemText.Append("\n");
                        }

                        systemText.Append(
                            Convert.ToString(input.content) ??
                            string.Empty);
                        continue;
                    }

                    var parts = TranslateParts(input.content);
                    if (parts.Count > 0)
                    {
                        contents.Add(new Dictionary<string, object>
                        {
                            {
                                "role",
                                string.Equals(
                                    input.role,
                                    "assistant",
                                    StringComparison
                                        .OrdinalIgnoreCase)
                                    ? "model"
                                    : "user"
                            },
                            { "parts", parts }
                        });
                    }

                    continue;
                }

                var assistant =
                    entry as ChatCompletionAssistantToolMessage;
                if (assistant != null)
                {
                    var parts = new List<object>();
                    if (!string.IsNullOrEmpty(assistant.content))
                    {
                        parts.Add(new Dictionary<string, object>
                        {
                            { "text", assistant.content }
                        });
                    }

                    foreach (var call in assistant.tool_calls ??
                        new List<ChatToolCall>())
                    {
                        if (call?.function == null)
                        {
                            continue;
                        }

                        var name = call.function.name ??
                            string.Empty;
                        if (!string.IsNullOrEmpty(call.id))
                        {
                            callNames[call.id] = name;
                        }

                        var callPart =
                            new Dictionary<string, object>
                            {
                                {
                                    "functionCall",
                                    new Dictionary<string, object>
                                    {
                                        { "name", name },
                                        {
                                            "args",
                                            ParseArguments(
                                                call.function
                                                    .arguments)
                                        }
                                    }
                                }
                            };
                        string signature;
                        if (thoughtSignatures != null &&
                            call.id != null &&
                            thoughtSignatures.TryGetValue(
                                call.id,
                                out signature) &&
                            signature.Length > 0)
                        {
                            callPart["thoughtSignature"] =
                                signature;
                        }

                        parts.Add(callPart);
                    }

                    if (parts.Count > 0)
                    {
                        contents.Add(new Dictionary<string, object>
                        {
                            { "role", "model" },
                            { "parts", parts }
                        });
                    }

                    continue;
                }

                var toolResult =
                    entry as ChatCompletionToolResultMessage;
                if (toolResult != null)
                {
                    string name;
                    if (toolResult.tool_call_id == null ||
                        !callNames.TryGetValue(
                            toolResult.tool_call_id,
                            out name) ||
                        name.Length == 0)
                    {
                        name = "tool";
                    }

                    contents.Add(new Dictionary<string, object>
                    {
                        { "role", "user" },
                        {
                            "parts",
                            new List<object>
                            {
                                new Dictionary<string, object>
                                {
                                    {
                                        "functionResponse",
                                        new Dictionary<string, object>
                                        {
                                            { "name", name },
                                            {
                                                "response",
                                                new Dictionary<string, object>
                                                {
                                                    {
                                                        "result",
                                                        toolResult
                                                            .content ??
                                                        string.Empty
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    });
                }
            }

            var translated = new Dictionary<string, object>
            {
                { "contents", contents }
            };
            if (systemText.Length > 0)
            {
                translated["systemInstruction"] =
                    new Dictionary<string, object>
                    {
                        { "role", "user" },
                        {
                            "parts",
                            new List<object>
                            {
                                new Dictionary<string, object>
                                {
                                    {
                                        "text",
                                        systemText.ToString()
                                    }
                                }
                            }
                        }
                    };
            }

            if (request.tools != null && request.tools.Count > 0)
            {
                var declarations = new List<object>();
                foreach (var tool in request.tools)
                {
                    if (tool?.function == null)
                    {
                        continue;
                    }

                    declarations.Add(new Dictionary<string, object>
                    {
                        { "name", tool.function.name },
                        {
                            "description",
                            tool.function.description
                        },
                        {
                            "parameters",
                            SanitizeSchema(
                                tool.function.parameters)
                        }
                    });
                }

                translated["tools"] = new List<object>
                {
                    new Dictionary<string, object>
                    {
                        { "functionDeclarations", declarations }
                    }
                };

                var forcedName = ReadForcedToolName(
                    request.tool_choice);
                if (forcedName.Length > 0)
                {
                    translated["toolConfig"] =
                        new Dictionary<string, object>
                        {
                            {
                                "functionCallingConfig",
                                new Dictionary<string, object>
                                {
                                    { "mode", "ANY" },
                                    {
                                        "allowedFunctionNames",
                                        new List<object>
                                        {
                                            forcedName
                                        }
                                    }
                                }
                            }
                        };
                }
            }

            var generationConfig =
                new Dictionary<string, object>();
            var thinkingBudget = ThinkingBudgetFor(request.model);
            if (thinkingBudget >= 0)
            {
                generationConfig["thinkingConfig"] =
                    new Dictionary<string, object>
                    {
                        { "thinkingBudget", thinkingBudget }
                    };
            }

            if (request.max_tokens.HasValue)
            {
                // Gemini models spend internal thinking tokens from
                // the same output budget, so a tight cap (the
                // endpoint probe, question generation) would return
                // no visible text at all. Headroom keeps the caps
                // meaningful while leaving room for thinking.
                generationConfig["maxOutputTokens"] =
                    request.max_tokens.Value + 4096;
            }

            if (generationConfig.Count > 0)
            {
                translated["generationConfig"] = generationConfig;
            }

            return translated;
        }

        // Latency: Gemini 2.5 Flash defaults to dynamic thinking,
        // which multiplies response time on ordinary mailbox
        // questions. Flash and Flash-Lite allow disabling thinking
        // entirely (budget 0); Pro cannot go below 128, so it gets
        // the minimum. Only the 2.5 family accepts thinkingBudget -
        // Gemini 3 models use a different thinking control and can
        // reject it - so everything else stays at server defaults.
        public static int ThinkingBudgetFor(string model)
        {
            var name = NormalizeModel(model).ToLowerInvariant();
            if (name.IndexOf(
                    "gemini-2.5",
                    StringComparison.Ordinal) < 0)
            {
                return -1;
            }

            return name.IndexOf(
                       "pro",
                       StringComparison.Ordinal) >= 0
                ? 128
                : 0;
        }

        private static string ReadForcedToolName(object toolChoice)
        {
            var map = toolChoice as IDictionary<string, object>;
            if (map == null)
            {
                return string.Empty;
            }

            object functionValue;
            map.TryGetValue("function", out functionValue);
            var function = functionValue
                as IDictionary<string, object>;
            return function != null
                ? ReadString(function, "name")
                : string.Empty;
        }

        private static List<object> TranslateParts(object content)
        {
            var parts = new List<object>();
            var text = content as string;
            if (text != null)
            {
                if (text.Length > 0)
                {
                    parts.Add(new Dictionary<string, object>
                    {
                        { "text", text }
                    });
                }

                return parts;
            }

            var list = content as System.Collections.IEnumerable;
            if (list == null)
            {
                return parts;
            }

            foreach (var item in list)
            {
                var textPart = item as ChatMultimodalTextPart;
                if (textPart != null &&
                    !string.IsNullOrEmpty(textPart.text))
                {
                    parts.Add(new Dictionary<string, object>
                    {
                        { "text", textPart.text }
                    });
                    continue;
                }

                var imagePart = item as ChatMultimodalImagePart;
                var url = imagePart?.image_url?.url ?? string.Empty;
                var inline = TranslateDataUrl(url);
                if (inline != null)
                {
                    parts.Add(new Dictionary<string, object>
                    {
                        { "inlineData", inline }
                    });
                }
            }

            return parts;
        }

        public static Dictionary<string, object> TranslateDataUrl(
            string url)
        {
            const string prefix = "data:";
            const string marker = ";base64,";
            if (url == null ||
                !url.StartsWith(
                    prefix,
                    StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            var markerIndex = url.IndexOf(
                marker,
                StringComparison.OrdinalIgnoreCase);
            if (markerIndex <= prefix.Length)
            {
                return null;
            }

            return new Dictionary<string, object>
            {
                {
                    "mimeType",
                    url.Substring(
                        prefix.Length,
                        markerIndex - prefix.Length)
                },
                {
                    "data",
                    url.Substring(markerIndex + marker.Length)
                }
            };
        }

        // Gemini's schema accepts a strict subset of JSON Schema, so
        // only supported keywords survive and type names become the
        // uppercase enum values the API expects.
        public static object SanitizeSchema(object schema)
        {
            var map = schema as IDictionary<string, object>;
            if (map == null)
            {
                return new Dictionary<string, object>
                {
                    { "type", "OBJECT" }
                };
            }

            var result = new Dictionary<string, object>();
            foreach (var pair in map)
            {
                switch (pair.Key)
                {
                    case "type":
                        result["type"] =
                            (Convert.ToString(pair.Value) ??
                             "object").ToUpperInvariant();
                        break;
                    case "description":
                    case "required":
                    case "enum":
                        result[pair.Key] = pair.Value;
                        break;
                    case "properties":
                        var properties = pair.Value
                            as IDictionary<string, object>;
                        if (properties != null)
                        {
                            var sanitized =
                                new Dictionary<string, object>();
                            foreach (var property in properties)
                            {
                                sanitized[property.Key] =
                                    SanitizeSchema(property.Value);
                            }

                            result["properties"] = sanitized;
                        }

                        break;
                    case "items":
                        result["items"] = SanitizeSchema(
                            pair.Value);
                        break;
                }
            }

            if (!result.ContainsKey("type"))
            {
                result["type"] = "OBJECT";
            }

            return result;
        }

        private static object ParseArguments(string arguments)
        {
            try
            {
                var serializer = new JavaScriptSerializer();
                return serializer.DeserializeObject(
                    arguments ?? string.Empty)
                    as IDictionary<string, object> ??
                    (object)new Dictionary<string, object>();
            }
            catch
            {
                return new Dictionary<string, object>();
            }
        }

        // ------------------------------------------------------------------
        // Response translation (Gemini native to OpenAI chat shape).
        // ------------------------------------------------------------------

        public static ChatCompletionResponseMessage
            TranslateResponse(IDictionary<string, object> response)
        {
            return TranslateResponse(response, null);
        }

        public static ChatCompletionResponseMessage
            TranslateResponse(
                IDictionary<string, object> response,
                IDictionary<string, string> thoughtSignatureSink)
        {
            if (response == null)
            {
                return null;
            }

            object candidatesValue;
            response.TryGetValue(
                "candidates",
                out candidatesValue);
            var candidates = candidatesValue as object[];
            var candidate = candidates != null &&
                candidates.Length > 0
                ? candidates[0] as IDictionary<string, object>
                : null;
            object contentValue = null;
            candidate?.TryGetValue("content", out contentValue);
            var content = contentValue
                as IDictionary<string, object>;
            object partsValue = null;
            content?.TryGetValue("parts", out partsValue);
            var parts = partsValue as object[];
            if (parts == null)
            {
                return null;
            }

            var text = new StringBuilder();
            var toolCalls = new List<ChatToolCall>();
            var serializer = new JavaScriptSerializer
            {
                MaxJsonLength = int.MaxValue
            };
            foreach (var entry in parts)
            {
                var part = entry as IDictionary<string, object>;
                if (part == null)
                {
                    continue;
                }

                // Parts flagged as thought are the model's internal
                // reasoning summary, not the answer.
                object thoughtValue;
                if (part.TryGetValue("thought", out thoughtValue) &&
                    thoughtValue is bool &&
                    (bool)thoughtValue)
                {
                    continue;
                }

                object textValue;
                if (part.TryGetValue("text", out textValue) &&
                    textValue is string)
                {
                    text.Append((string)textValue);
                    continue;
                }

                object callValue;
                if (part.TryGetValue(
                        "functionCall",
                        out callValue))
                {
                    var call = callValue
                        as IDictionary<string, object>;
                    if (call == null)
                    {
                        continue;
                    }

                    object argsValue;
                    call.TryGetValue("args", out argsValue);
                    // Globally unique ids so replayed history from
                    // earlier tool rounds never collides in the
                    // signature map.
                    var callId = "gemini_call_" +
                        Guid.NewGuid().ToString("N")
                            .Substring(0, 12);
                    var signature = ReadString(
                        part,
                        "thoughtSignature");
                    if (thoughtSignatureSink != null &&
                        signature.Length > 0)
                    {
                        thoughtSignatureSink[callId] = signature;
                    }

                    toolCalls.Add(new ChatToolCall
                    {
                        id = callId,
                        type = "function",
                        function = new ChatToolCallFunction
                        {
                            name = ReadString(call, "name"),
                            arguments = serializer.Serialize(
                                argsValue ??
                                new Dictionary<string, object>())
                        }
                    });
                }
            }

            var boundedText = TextBoundary.PlainText(
                text.ToString(),
                TextBoundary.MaxAssistantCharacters);
            if (boundedText.Length == 0 && toolCalls.Count == 0)
            {
                return null;
            }

            return new ChatCompletionResponseMessage
            {
                role = "assistant",
                content = boundedText,
                tool_calls = toolCalls.Count > 0
                    ? toolCalls
                    : null
            };
        }

        // ------------------------------------------------------------------
        // HTTP plumbing.
        // ------------------------------------------------------------------

        private async Task<IDictionary<string, object>>
            PostJsonAsync(
                HttpClient httpClient,
                AppSettings settings,
                string method,
                Dictionary<string, object> payload,
                CancellationToken cancellationToken)
        {
            var json = _serializer.Serialize(payload);
            var token = await GetAccessTokenAsync(
                httpClient,
                settings,
                cancellationToken).ConfigureAwait(true);
            using (var request = new HttpRequestMessage(
                HttpMethod.Post,
                ApiBase + method))
            {
                request.Headers.Authorization =
                    new AuthenticationHeaderValue(
                        "Bearer",
                        token);
                request.Headers.Accept.Add(
                    new MediaTypeWithQualityHeaderValue(
                        "application/json"));
                request.Content = new StringContent(
                    json,
                    Encoding.UTF8,
                    "application/json");
                using (var response = await httpClient
                    .SendAsync(request, cancellationToken)
                    .ConfigureAwait(true))
                {
                    var body = await ReadBodyAsync(
                        response,
                        cancellationToken)
                        .ConfigureAwait(true);
                    if (response.IsSuccessStatusCode)
                    {
                        return _serializer
                            .DeserializeObject(body)
                            as IDictionary<string, object>;
                    }

                    var status = (int)response.StatusCode;
                    // Capacity handling lives in the callers'
                    // retry loops (fast probes, backoff, and
                    // model fallback); this layer only reports.
                    if (status == 429)
                    {
                        throw new AiEndpointException(
                            "GEMINI_RATE_LIMITED",
                            "The Gemini quota for this model is " +
                            "exhausted right now. Wait a minute " +
                            "and try again, or switch to " +
                            "gemini-2.5-flash.",
                            httpStatus: status,
                            responseSnippet: body);
                    }

                    var hint = status == 401 || status == 403
                        ? " The Google sign-in may have " +
                          "expired - open Scribble Settings " +
                          "and click Sign in with Google " +
                          "again."
                        : string.Empty;
                    throw new AiEndpointException(
                        "GEMINI_HTTP_" + status,
                        "The Gemini endpoint rejected " +
                        method.TrimStart(':') + ": " + status +
                        "." + hint,
                        httpStatus: status,
                        responseSnippet: body);
                }
            }
        }

        // Reads the retry hint from a 429 body: either RetryInfo's
        // "retryDelay": "56s" or message text like "after 56s".
        public static int ParseRetryDelaySeconds(string body)
        {
            var match = System.Text.RegularExpressions.Regex.Match(
                body ?? string.Empty,
                "\"retryDelay\"\\s*:\\s*\"(\\d+)");
            if (!match.Success)
            {
                match = System.Text.RegularExpressions.Regex.Match(
                    body ?? string.Empty,
                    "after\\s+(\\d+)\\s*s");
            }

            int seconds;
            return match.Success &&
                   int.TryParse(
                       match.Groups[1].Value,
                       out seconds)
                ? seconds
                : 0;
        }

        // GET against the API base; used to poll the onboarding
        // long-running operation (path "/operations/...").
        private async Task<IDictionary<string, object>>
            GetJsonAsync(
                HttpClient httpClient,
                AppSettings settings,
                string path,
                CancellationToken cancellationToken)
        {
            var token = await GetAccessTokenAsync(
                httpClient,
                settings,
                cancellationToken).ConfigureAwait(true);
            using (var request = new HttpRequestMessage(
                HttpMethod.Get,
                ApiBase + path))
            {
                request.Headers.Authorization =
                    new AuthenticationHeaderValue(
                        "Bearer",
                        token);
                request.Headers.Accept.Add(
                    new MediaTypeWithQualityHeaderValue(
                        "application/json"));
                using (var response = await httpClient
                    .SendAsync(request, cancellationToken)
                    .ConfigureAwait(true))
                {
                    var body = await ReadBodyAsync(
                        response,
                        cancellationToken).ConfigureAwait(true);
                    if (!response.IsSuccessStatusCode)
                    {
                        return null;
                    }

                    return _serializer.DeserializeObject(body)
                        as IDictionary<string, object>;
                }
            }
        }

        private static async Task<string> ReadBodyAsync(
            HttpResponseMessage response,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var body = await response.Content
                .ReadAsStringAsync().ConfigureAwait(true);
            if (body.Length >
                TextBoundary.MaxHttpResponseCharacters)
            {
                throw new AiEndpointException(
                    "RESPONSE_TOO_LARGE",
                    "The Gemini endpoint response was too large.");
            }

            return body;
        }

        private static string ReadString(
            IDictionary<string, object> map,
            string key)
        {
            object value;
            return map != null && map.TryGetValue(key, out value)
                ? Convert.ToString(value) ?? string.Empty
                : string.Empty;
        }

        private static long NowUtcMilliseconds()
        {
            return (long)(DateTime.UtcNow -
                new DateTime(
                    1970,
                    1,
                    1,
                    0,
                    0,
                    0,
                    DateTimeKind.Utc)).TotalMilliseconds;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net;
using System.Net.Http;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace Scribble.Configuration
{
    public sealed class GoogleSignInResult
    {
        public GoogleSignInResult(
            string refreshToken,
            string accessToken,
            long expiresInSeconds)
        {
            RefreshToken = refreshToken ?? string.Empty;
            AccessToken = accessToken ?? string.Empty;
            ExpiresInSeconds = expiresInSeconds;
        }

        public string RefreshToken { get; }

        public string AccessToken { get; }

        public long ExpiresInSeconds { get; }
    }

    // Browser-based Google sign-in for Gemini: the standard OAuth
    // installed-app flow with PKCE and a loopback redirect - the
    // same flow, OAuth client, and scopes the open-source Gemini CLI
    // uses, so anything an admin has already allowed for Gemini CLI
    // works here identically. The user's password never touches
    // Scribble: Google's own pages handle the sign-in and Scribble only
    // receives tokens on 127.0.0.1.
    public static class GoogleSignInFlow
    {
        // Published installed-app OAuth client of the open-source
        // Gemini CLI (Apache-2.0). Installed-app client secrets are
        // not confidential; token requests require them.
        public const string OAuthClientId =
            "681255809395-oo8ft2oprdrnp9e3aqf6av3hmdib135j" +
            ".apps.googleusercontent.com";
        public const string OAuthClientSecret =
            "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";
        public const string TokenEndpoint =
            "https://oauth2.googleapis.com/token";
        private const string AuthorizeEndpoint =
            "https://accounts.google.com/o/oauth2/v2/auth";
        private const string Scopes =
            "https://www.googleapis.com/auth/cloud-platform " +
            "https://www.googleapis.com/auth/userinfo.email " +
            "https://www.googleapis.com/auth/userinfo.profile";

        public static async Task<GoogleSignInResult> SignInAsync(
            HttpClient httpClient,
            TimeSpan timeout)
        {
            if (AdminPolicy.GeminiDisabled)
            {
                throw new InvalidOperationException(
                    "Google Gemini sign-in is unavailable in this build.");
            }

            var port = FindFreeLoopbackPort();
            var redirectUri =
                "http://localhost:" + port + "/oauth2callback";
            var verifier = RandomUrlSafeString(64);
            var challenge = Base64Url(
                Sha256(Encoding.ASCII.GetBytes(verifier)));
            var state = RandomUrlSafeString(32);
            var authorizeUrl = AuthorizeEndpoint +
                "?client_id=" +
                Uri.EscapeDataString(OAuthClientId) +
                "&redirect_uri=" +
                Uri.EscapeDataString(redirectUri) +
                "&response_type=code" +
                "&scope=" + Uri.EscapeDataString(Scopes) +
                "&access_type=offline" +
                "&prompt=consent" +
                "&code_challenge=" + challenge +
                "&code_challenge_method=S256" +
                "&state=" + state;

            string code;
            using (var listener = new HttpListener())
            {
                listener.Prefixes.Add(
                    "http://localhost:" + port + "/");
                listener.Start();
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = authorizeUrl,
                        UseShellExecute = true
                    });

                    // Browsers also request favicon.ico and may
                    // probe the port, so requests are accepted in a
                    // loop until the one carrying the OAuth callback
                    // (code or error query) arrives; everything else
                    // gets a 404 and the wait continues.
                    var deadline = DateTime.UtcNow + timeout;
                    string error;
                    string returnedState;
                    while (true)
                    {
                        var remaining = deadline - DateTime.UtcNow;
                        if (remaining <= TimeSpan.Zero)
                        {
                            throw new InvalidOperationException(
                                "The Google sign-in timed out. " +
                                "Try again and complete the " +
                                "sign-in in the browser.");
                        }

                        var contextTask =
                            listener.GetContextAsync();
                        var finished = await Task.WhenAny(
                            contextTask,
                            Task.Delay(remaining))
                            .ConfigureAwait(true);
                        if (finished != contextTask)
                        {
                            ObserveAbandonedTask(contextTask);
                            throw new InvalidOperationException(
                                "The Google sign-in timed out. " +
                                "Try again and complete the " +
                                "sign-in in the browser.");
                        }

                        var context = contextTask.Result;
                        var query = context.Request.QueryString;
                        error = query["error"] ?? string.Empty;
                        returnedState =
                            query["state"] ?? string.Empty;
                        code = query["code"] ?? string.Empty;
                        if (code.Length == 0 && error.Length == 0)
                        {
                            try
                            {
                                context.Response.StatusCode = 404;
                                context.Response.Close();
                            }
                            catch
                            {
                            }

                            continue;
                        }

                        var success = error.Length == 0 &&
                            code.Length > 0 &&
                            string.Equals(
                                returnedState,
                                state,
                                StringComparison.Ordinal);
                        WriteBrowserResponse(
                            context.Response,
                            success);
                        break;
                    }

                    if (error.Length > 0)
                    {
                        throw new InvalidOperationException(
                            "Google reported: " + error);
                    }

                    if (!string.Equals(
                        returnedState,
                        state,
                        StringComparison.Ordinal))
                    {
                        throw new InvalidOperationException(
                            "The sign-in response did not match " +
                            "this request. Try again.");
                    }
                }
                finally
                {
                    try
                    {
                        listener.Stop();
                    }
                    catch
                    {
                    }
                }
            }

            using (var request = new HttpRequestMessage(
                HttpMethod.Post,
                TokenEndpoint))
            {
                request.Content = new FormUrlEncodedContent(
                    new Dictionary<string, string>
                    {
                        { "code", code },
                        { "client_id", OAuthClientId },
                        { "client_secret", OAuthClientSecret },
                        { "redirect_uri", redirectUri },
                        { "grant_type", "authorization_code" },
                        { "code_verifier", verifier }
                    });
                using (var response = await httpClient
                    .SendAsync(request).ConfigureAwait(true))
                {
                    var body = await response.Content
                        .ReadAsStringAsync().ConfigureAwait(true);
                    if (!response.IsSuccessStatusCode)
                    {
                        throw new InvalidOperationException(
                            "The Google token exchange failed (" +
                            (int)response.StatusCode + "): " +
                            Truncate(body, 300));
                    }

                    var serializer = new JavaScriptSerializer();
                    var map = serializer.DeserializeObject(body)
                        as IDictionary<string, object>;
                    var refreshToken = ReadString(
                        map,
                        "refresh_token");
                    var accessToken = ReadString(
                        map,
                        "access_token");
                    if (refreshToken.Length == 0)
                    {
                        throw new InvalidOperationException(
                            "Google returned no refresh token. " +
                            "Remove Scribble's access at " +
                            "myaccount.google.com/permissions " +
                            "and sign in again.");
                    }

                    long expiresIn = 3600;
                    object expiresValue = null;
                    map?.TryGetValue(
                        "expires_in",
                        out expiresValue);
                    if (expiresValue != null)
                    {
                        long.TryParse(
                            Convert.ToString(
                                expiresValue,
                                System.Globalization.CultureInfo
                                    .InvariantCulture),
                            out expiresIn);
                    }

                    return new GoogleSignInResult(
                        refreshToken,
                        accessToken,
                        expiresIn);
                }
            }
        }

        // A pending GetContextAsync faults when the listener stops
        // after a timeout; observing the task keeps that expected
        // fault away from the unobserved-exception handler.
        private static void ObserveAbandonedTask(
            Task<HttpListenerContext> task)
        {
            task.ContinueWith(
                completed =>
                {
                    var ignored = completed.Exception;
                },
                TaskContinuationOptions.OnlyOnFaulted);
        }

        private static void WriteBrowserResponse(
            HttpListenerResponse response,
            bool success)
        {
            try
            {
                var html =
                    "<!DOCTYPE html><html><head><meta charset=" +
                    "\"utf-8\"><title>Scribble</title></head>" +
                    "<body style=\"font-family:Segoe UI," +
                    "sans-serif;background:#1a1b1e;color:#e8e8ec;" +
                    "display:flex;align-items:center;" +
                    "justify-content:center;height:100vh;\">" +
                    "<div style=\"text-align:center;\"><h2>" +
                    (success
                        ? "Signed in"
                        : "Sign-in did not complete") +
                    "</h2><p>" +
                    (success
                        ? "You can close this tab and return " +
                          "to Outlook."
                        : "Close this tab and try again from " +
                          "Scribble Settings.") +
                    "</p></div></body></html>";
                var bytes = Encoding.UTF8.GetBytes(html);
                response.ContentType = "text/html; charset=utf-8";
                response.ContentLength64 = bytes.Length;
                response.OutputStream.Write(
                    bytes,
                    0,
                    bytes.Length);
                response.OutputStream.Close();
            }
            catch
            {
            }
        }

        private static int FindFreeLoopbackPort()
        {
            var probe = new TcpListener(IPAddress.Loopback, 0);
            probe.Start();
            var port =
                ((IPEndPoint)probe.LocalEndpoint).Port;
            probe.Stop();
            return port;
        }

        private static string RandomUrlSafeString(int byteCount)
        {
            var bytes = new byte[byteCount];
            using (var random = new RNGCryptoServiceProvider())
            {
                random.GetBytes(bytes);
            }

            return Base64Url(bytes);
        }

        private static byte[] Sha256(byte[] bytes)
        {
            using (var sha = SHA256.Create())
            {
                return sha.ComputeHash(bytes);
            }
        }

        private static string Base64Url(byte[] bytes)
        {
            return Convert.ToBase64String(bytes)
                .TrimEnd('=')
                .Replace('+', '-')
                .Replace('/', '_');
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

        private static string Truncate(string value, int length)
        {
            var text = value ?? string.Empty;
            return text.Length <= length
                ? text
                : text.Substring(0, length);
        }
    }
}

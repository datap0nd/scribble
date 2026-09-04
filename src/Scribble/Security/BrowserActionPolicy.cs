using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Scribble.Security
{
    public sealed class BrowserActionDescriptor
    {
        public string Action { get; set; }

        public string TagName { get; set; }

        public string InputType { get; set; }

        public string Role { get; set; }

        public string HtmlName { get; set; }

        public string Name { get; set; }

        public string VisibleLabel { get; set; }

        public string GroupLabel { get; set; }

        public string Placeholder { get; set; }

        public string Autocomplete { get; set; }

        public string LinkTarget { get; set; }

        public string Url { get; set; }

        public string Value { get; set; }

        public string SourceText { get; set; }

        public string Key { get; set; }

        public bool IsSubmit { get; set; }

        public bool FormHasPassword { get; set; }

        public bool FormHasPayment { get; set; }

        public bool FormHasPersonalData { get; set; }
    }

    public sealed class BrowserActionDecision
    {
        public BrowserActionDecision(
            bool allowed,
            string code,
            string message)
        {
            Allowed = allowed;
            Code = code ?? string.Empty;
            Message = message ?? string.Empty;
        }

        public bool Allowed { get; }

        public string Code { get; }

        public string Message { get; }
    }

    // Pure policy gate for model-requested browser actions. The extension
    // supplies value-free DOM metadata and cannot dispatch CDP input until
    // this policy returns Allowed. Keep this class independent of Chrome and
    // COM so the hand-rolled guardrail test runner can exercise every rule.
    public static class BrowserActionPolicy
    {
        public const int MaxTypedCharacters = 200;

        private static readonly ISet<string> AllowedActions =
            new HashSet<string>(StringComparer.Ordinal)
            {
                "click",
                "type",
                "select",
                "check",
                "press",
                "hover",
                "scroll",
                "wait"
            };

        private static readonly Regex SensitiveAutocomplete =
            new Regex(
                @"(^|\s)(name|honorific-prefix|given-name|additional-name|family-name|" +
                @"nickname|username|new-password|current-password|one-time-code|" +
                @"organization|street-address|address-line[123]|address-level[1-4]|" +
                @"postal-code|country|country-name|email|tel|tel-country-code|" +
                @"tel-national|tel-area-code|tel-local|cc-[^\s]+|transaction-[^\s]+)(\s|$)",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex SensitiveFieldText =
            new Regex(
                @"\b(password|passcode|log[ -]?in|sign[ -]?in|sign[ -]?up|register|" +
                @"email(?: address)?|phone|telephone|mobile|street address|postal|zip code|" +
                @"card(?:holder| number)?|credit card|debit card|cvv|cvc|iban|bank account|" +
                @"billing|travell?er name|passenger name|first name|last name|full name)\b",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex ConsequentialControlText =
            new Regex(
                @"\b(buy|purchase|checkout|check out|pay|payment|place order|confirm order|" +
                @"book now|continue to book|complete booking|confirm booking|reserve now|make reservation|" +
                @"sign[ -]?in|log[ -]?in|sign[ -]?up|register|" +
                @"subscribe|unsubscribe|send|post|upload|download|delete|remove account|" +
                @"submit application|add to (?:cart|basket|bag)|agree|accept terms|consent)\b",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex ReversibleCommerceLinkText =
            new Regex(
                @"\b(buy|purchase)\b",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static BrowserActionDecision Evaluate(
            BrowserActionDescriptor descriptor)
        {
            if (descriptor == null)
            {
                return Deny(
                    "ACTION_DESCRIPTOR_MISSING",
                    "I couldn't evaluate the browser action because it did not include a control descriptor.");
            }

            var action = Clean(descriptor.Action, 40).ToLowerInvariant();
            if (!AllowedActions.Contains(action))
            {
                return Deny(
                    "ACTION_NOT_ALLOWED",
                    "I can't run that browser action because it is not in my allowlist.");
            }

            Uri page;
            if (!Uri.TryCreate(descriptor.Url, UriKind.Absolute, out page) ||
                (page.Scheme != Uri.UriSchemeHttp &&
                 page.Scheme != Uri.UriSchemeHttps))
            {
                return Deny(
                    "ACTION_URL_BLOCKED",
                    "I restrict browser actions to HTTP and HTTPS work tabs.");
            }

            if (action == "scroll" || action == "wait")
            {
                return Allow();
            }

            var googleSearchAction = IsGoogleSearchAction(
                descriptor,
                page,
                action);

            var valueEntryAction = action == "type" ||
                action == "select" || action == "check";
            var inputType = Clean(descriptor.InputType, 40)
                .ToLowerInvariant();
            if (inputType == "password" || inputType == "file" ||
                (valueEntryAction && IsSensitiveField(descriptor)))
            {
                return Deny(
                    "ACTION_SENSITIVE_FIELD",
                    "I can't interact with credential, payment, or personal-data fields.");
            }

            if (descriptor.FormHasPassword ||
                (!googleSearchAction && descriptor.IsSubmit &&
                 (descriptor.FormHasPayment ||
                  descriptor.FormHasPersonalData)))
            {
                return Deny(
                    "ACTION_SENSITIVE_FORM",
                    "I can't use that target because it is inside a credential, payment, or personal-data form.");
            }

            var controlText = Clean(
                (descriptor.Name ?? string.Empty) + " " +
                (descriptor.Placeholder ?? string.Empty),
                500);
            if (ConsequentialControlText.IsMatch(controlText) &&
                !IsReversibleCommerceLink(descriptor, controlText))
            {
                return Deny(
                    "ACTION_CONSEQUENTIAL",
                    "I stop before booking, payment, authentication, messaging, downloads, or destructive actions.");
            }

            if (action == "type")
            {
                var value = Clean(descriptor.Value, MaxTypedCharacters + 1);
                if (value.Length == 0)
                {
                    return Deny(
                        "TYPE_VALUE_MISSING",
                        "I need a non-empty typed value.");
                }

                if (value.Length > MaxTypedCharacters)
                {
                    return Deny(
                        "TYPE_VALUE_TOO_LONG",
                        "I limit typed browser values to 200 characters.");
                }

                if (!IsTypedValueAuthorized(
                    value,
                    descriptor.SourceText))
                {
                    return Deny(
                        "TYPE_SOURCE_NOT_USER",
                        "I can type only text from your prompt, a locally validated public alias, or a clarification answer.");
                }
            }

            return Allow();
        }

        public static bool IsSensitiveField(
            BrowserActionDescriptor descriptor)
        {
            if (descriptor == null)
            {
                return true;
            }

            var type = Clean(descriptor.InputType, 40)
                .ToLowerInvariant();
            if (type == "password" ||
                type == "email" ||
                type == "tel" ||
                type == "file")
            {
                return true;
            }

            var autocomplete = Clean(
                descriptor.Autocomplete,
                300);
            if (SensitiveAutocomplete.IsMatch(autocomplete))
            {
                return true;
            }

            var text = Clean(
                (descriptor.Name ?? string.Empty) + " " +
                (descriptor.Placeholder ?? string.Empty),
                500);
            // Passenger counts are public search criteria, not identity.
            if (Regex.IsMatch(
                    text,
                    @"\b(passengers?|travell?ers?|adults?|children|infants?)\b.*\b(count|number|how many)\b|" +
                    @"\b(count|number|how many)\b.*\b(passengers?|travell?ers?|adults?|children|infants?)\b",
                    RegexOptions.IgnoreCase))
            {
                return false;
            }

            return SensitiveFieldText.IsMatch(text);
        }

        private static bool IsReversibleCommerceLink(
            BrowserActionDescriptor descriptor,
            string controlText)
        {
            if (descriptor == null || descriptor.IsSubmit ||
                !(string.Equals(
                      Clean(descriptor.TagName, 40),
                      "a",
                      StringComparison.OrdinalIgnoreCase) ||
                  string.Equals(
                      Clean(descriptor.Role, 80),
                      "link",
                      StringComparison.OrdinalIgnoreCase)) ||
                !ReversibleCommerceLinkText.IsMatch(controlText))
            {
                return false;
            }

            Uri destination;
            return Uri.TryCreate(
                    descriptor.LinkTarget,
                    UriKind.Absolute,
                    out destination) &&
                (destination.Scheme == Uri.UriSchemeHttp ||
                 destination.Scheme == Uri.UriSchemeHttps);
        }

        public static bool IsContiguousUserPhrase(
            string value,
            string sourceText)
        {
            var wanted = NormalizePhrase(value);
            var source = NormalizePhrase(sourceText);
            if (wanted.Length == 0 || source.Length == 0)
            {
                return false;
            }

            return (" " + source + " ").IndexOf(
                " " + wanted + " ",
                StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static bool IsTypedValueDerivedFromUser(
            string value,
            string sourceText)
        {
            var wanted = QueryTokens(value);
            var source = new HashSet<string>(
                QueryTokens(sourceText),
                StringComparer.OrdinalIgnoreCase);
            if (wanted.Count == 0 || source.Count == 0)
            {
                return false;
            }

            foreach (var token in wanted)
            {
                if (!source.Contains(token))
                {
                    return false;
                }
            }

            return true;
        }

        public static bool IsTypedValueAuthorized(
            string value,
            string sourceText)
        {
            if (IsTypedValueDerivedFromUser(value, sourceText))
            {
                return true;
            }

            var source = new HashSet<string>(
                QueryTokens(sourceText),
                StringComparer.OrdinalIgnoreCase);
            var allowed = new HashSet<string>(
                source,
                StringComparer.OrdinalIgnoreCase);
            foreach (var inference in PublicSearchInferences)
            {
                if (!source.Contains(inference.Key))
                {
                    continue;
                }

                foreach (var alias in inference.Value)
                {
                    allowed.Add(CanonicalQueryToken(alias));
                }
            }

            var wanted = QueryTokens(value);
            return wanted.Count > 0 &&
                wanted.All(token => allowed.Contains(token));
        }

        private static readonly IDictionary<string, string[]>
            PublicSearchInferences =
                new Dictionary<string, string[]>(
                    StringComparer.OrdinalIgnoreCase)
                {
                    { "dubai", new[] { "dxb", "international", "airport" } },
                    { "sharjah", new[] { "shj", "international", "airport" } },
                    { "lisbon", new[] { "lis", "airport" } },
                    { "seoul", new[] { "icn", "gmp", "incheon", "airport" } },
                    { "london", new[] { "lhr", "lgw", "airport" } },
                    { "york", new[] { "jfk", "lga", "ewr", "airport" } },
                    { "paris", new[] { "cdg", "ory", "airport" } },
                    { "tokyo", new[] { "hnd", "nrt", "airport" } },
                    { "singapore", new[] { "sin", "changi", "airport" } },
                    { "january", new[] { "jan", "1", "01" } },
                    { "february", new[] { "feb", "2", "02" } },
                    { "march", new[] { "mar", "3", "03" } },
                    { "april", new[] { "apr", "4", "04" } },
                    { "may", new[] { "5", "05" } },
                    { "june", new[] { "jun", "6", "06" } },
                    { "july", new[] { "jul", "7", "07" } },
                    { "august", new[] { "aug", "8", "08" } },
                    { "september", new[] { "sep", "sept", "9", "09" } },
                    { "october", new[] { "oct", "10" } },
                    { "november", new[] { "nov", "11" } },
                    { "december", new[] { "dec", "12" } }
                };

        private static bool IsGoogleSearchAction(
            BrowserActionDescriptor descriptor,
            Uri page,
            string action)
        {
            if (descriptor == null || page == null ||
                !IsGoogleHost(page.Host))
            {
                return false;
            }

            var role = Clean(descriptor.Role, 80)
                .ToLowerInvariant();
            var tagName = Clean(descriptor.TagName, 40)
                .ToLowerInvariant();
            var inputType = Clean(descriptor.InputType, 40)
                .ToLowerInvariant();
            var htmlName = Clean(descriptor.HtmlName, 120)
                .ToLowerInvariant();
            var label = Clean(
                (descriptor.Name ?? string.Empty) + " " +
                (descriptor.Placeholder ?? string.Empty),
                500);
            var looksLikeSearch =
                (role == "searchbox" ||
                 role == "textbox" ||
                 role == "combobox" ||
                 inputType == "search") &&
                (tagName.Length == 0 || tagName == "input" ||
                 tagName == "textarea") &&
                (htmlName == "q" || inputType == "search" ||
                 Regex.IsMatch(
                     label,
                     @"\b(search|google)\b",
                     RegexOptions.IgnoreCase));
            if (!looksLikeSearch)
            {
                return false;
            }

            if (action == "type")
            {
                return true;
            }

            return action == "press" &&
                Clean(descriptor.Key, 40).Equals(
                    "Enter",
                    StringComparison.OrdinalIgnoreCase);
        }

        public static bool IsGoogleHost(string host)
        {
            return Regex.IsMatch(
                Clean(host, 253),
                @"^(?:[a-z0-9-]+\.)*google\.[a-z]{2,3}(?:\.[a-z]{2})?$",
                RegexOptions.IgnoreCase);
        }

        private static IList<string> QueryTokens(string value)
        {
            var normalized = NormalizePhrase(value);
            var result = new List<string>();
            foreach (var raw in normalized.Split(
                new[] { ' ' },
                StringSplitOptions.RemoveEmptyEntries))
            {
                var token = CanonicalQueryToken(raw);
                if (token.Length > 4 && token.EndsWith(
                    "ies",
                    StringComparison.OrdinalIgnoreCase))
                {
                    token = token.Substring(0, token.Length - 3) + "y";
                }
                else if (token.Length > 3 &&
                         token.EndsWith(
                             "s",
                             StringComparison.OrdinalIgnoreCase) &&
                         !token.EndsWith(
                             "ss",
                             StringComparison.OrdinalIgnoreCase))
                {
                    token = token.Substring(0, token.Length - 1);
                }

                result.Add(token);
            }

            return result;
        }

        private static string CanonicalQueryToken(string value)
        {
            var token = value ?? string.Empty;
            int numeric;
            return int.TryParse(token, out numeric)
                ? numeric.ToString(
                    System.Globalization.CultureInfo.InvariantCulture)
                : token;
        }

        private static BrowserActionDecision Allow()
        {
            return new BrowserActionDecision(
                true,
                "ACTION_ALLOWED",
                "I allowed the bounded browser action.");
        }

        private static BrowserActionDecision Deny(
            string code,
            string message)
        {
            return new BrowserActionDecision(false, code, message);
        }

        private static string Clean(string value, int maximum)
        {
            return TextBoundary.SingleLine(value, maximum);
        }

        private static string NormalizePhrase(string value)
        {
            var safe = TextBoundary.PlainText(
                value,
                TextBoundary.MaxUserPromptCharacters);
            var result = new StringBuilder(safe.Length);
            var pendingSpace = false;
            foreach (var character in safe)
            {
                if (char.IsLetterOrDigit(character))
                {
                    if (pendingSpace && result.Length > 0)
                    {
                        result.Append(' ');
                    }

                    result.Append(char.ToLowerInvariant(character));
                    pendingSpace = false;
                }
                else
                {
                    pendingSpace = true;
                }
            }

            return result.ToString().Trim();
        }
    }
}

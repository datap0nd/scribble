using System;
using System.Collections.Generic;
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

        public string Name { get; set; }

        public string Placeholder { get; set; }

        public string Autocomplete { get; set; }

        public string Url { get; set; }

        public string Value { get; set; }

        public string SourceText { get; set; }

        public string Key { get; set; }

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
                @"submit application|add to (?:cart|basket|bag))\b",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static BrowserActionDecision Evaluate(
            BrowserActionDescriptor descriptor)
        {
            if (descriptor == null)
            {
                return Deny(
                    "ACTION_DESCRIPTOR_MISSING",
                    "The browser action did not include a control descriptor.");
            }

            var action = Clean(descriptor.Action, 40).ToLowerInvariant();
            if (!AllowedActions.Contains(action))
            {
                return Deny(
                    "ACTION_NOT_ALLOWED",
                    "That browser action is not in Scribble's allowlist.");
            }

            Uri page;
            if (!Uri.TryCreate(descriptor.Url, UriKind.Absolute, out page) ||
                (page.Scheme != Uri.UriSchemeHttp &&
                 page.Scheme != Uri.UriSchemeHttps))
            {
                return Deny(
                    "ACTION_URL_BLOCKED",
                    "Browser actions are restricted to HTTP and HTTPS work tabs.");
            }

            if (action == "scroll" || action == "wait")
            {
                return Allow();
            }

            var googleSearchAction = IsGoogleSearchAction(
                descriptor,
                page,
                action);

            if (IsSensitiveField(descriptor))
            {
                return Deny(
                    "ACTION_SENSITIVE_FIELD",
                    "Scribble cannot interact with credential, payment, or personal-data fields.");
            }

            if (descriptor.FormHasPassword ||
                (!googleSearchAction &&
                 (descriptor.FormHasPayment ||
                  descriptor.FormHasPersonalData)))
            {
                return Deny(
                    "ACTION_SENSITIVE_FORM",
                    "The target is inside a credential, payment, or personal-data form.");
            }

            var controlText = Clean(
                (descriptor.Name ?? string.Empty) + " " +
                (descriptor.Placeholder ?? string.Empty),
                500);
            if (ConsequentialControlText.IsMatch(controlText))
            {
                return Deny(
                    "ACTION_CONSEQUENTIAL",
                    "Scribble stops before booking, payment, authentication, messaging, downloads, or destructive actions.");
            }

            if (action == "type")
            {
                var value = Clean(descriptor.Value, MaxTypedCharacters + 1);
                if (value.Length == 0)
                {
                    return Deny(
                        "TYPE_VALUE_MISSING",
                        "A non-empty typed value is required.");
                }

                if (value.Length > MaxTypedCharacters)
                {
                    return Deny(
                        "TYPE_VALUE_TOO_LONG",
                        "Typed browser values are limited to 200 characters.");
                }

                var sourceAllowed = googleSearchAction
                    ? IsGoogleQueryDerivedFromUser(
                        value,
                        descriptor.SourceText)
                    : IsContiguousUserPhrase(
                        value,
                        descriptor.SourceText);
                if (!sourceAllowed)
                {
                    return Deny(
                        "TYPE_SOURCE_NOT_USER",
                        "Typed text must come directly from the user's prompt or clarification answers.");
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

        public static bool IsGoogleQueryDerivedFromUser(
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

        private static bool IsGoogleSearchAction(
            BrowserActionDescriptor descriptor,
            Uri page,
            string action)
        {
            if (descriptor == null || page == null ||
                !(page.Host.Equals(
                      "google.com",
                      StringComparison.OrdinalIgnoreCase) ||
                  page.Host.EndsWith(
                      ".google.com",
                      StringComparison.OrdinalIgnoreCase)))
            {
                return false;
            }

            var role = Clean(descriptor.Role, 80)
                .ToLowerInvariant();
            var inputType = Clean(descriptor.InputType, 40)
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
                Regex.IsMatch(
                    label,
                    @"\b(search|google)\b",
                    RegexOptions.IgnoreCase);
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

        private static IList<string> QueryTokens(string value)
        {
            var normalized = NormalizePhrase(value);
            var result = new List<string>();
            foreach (var raw in normalized.Split(
                new[] { ' ' },
                StringSplitOptions.RemoveEmptyEntries))
            {
                var token = raw;
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

        private static BrowserActionDecision Allow()
        {
            return new BrowserActionDecision(
                true,
                "ACTION_ALLOWED",
                "The bounded browser action is allowed.");
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

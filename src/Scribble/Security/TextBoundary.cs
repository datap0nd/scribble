using System;
using System.Text;

namespace Scribble.Security
{
    public static class TextBoundary
    {
        // Recommended defaults. The effective Max* values below can
        // be adjusted by the user from the Settings Limits tab
        // (LimitOverrides), always inside hard clamps; drafting and
        // sending capability rules are never adjustable.
        public const int RecommendedUserPromptCharacters = 4000;
        public const int RecommendedAssistantCharacters = 12000;
        public const int RecommendedConversationTurns = 12;
        public const int RecommendedToolRounds = 4;
        public const int RecommendedToolCallsPerRound = 4;

        public const int MaxMessageBodyCharacters = 24000;
        public const int MaxToneProfileCharacters = 5000;
        public const int MaxToolResultCharacters = 120000;
        public const int MaxHttpResponseCharacters = 1048576;

        public static int MaxUserPromptCharacters
        {
            get { return LimitOverrides.PromptCharacters; }
        }

        public static int MaxAssistantCharacters
        {
            get { return LimitOverrides.AssistantCharacters; }
        }

        public static int MaxConversationTurns
        {
            get { return LimitOverrides.HistoryTurns; }
        }

        public static int MaxToolRounds
        {
            get { return LimitOverrides.ToolRounds; }
        }

        public static int MaxToolCallsPerRound
        {
            get { return LimitOverrides.ToolCallsPerRound; }
        }

        public static string PlainText(string value, int maximumLength)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            var limit = Math.Max(0, maximumLength);
            var builder = new StringBuilder(Math.Min(value.Length, limit));

            foreach (var character in value)
            {
                if (builder.Length >= limit)
                {
                    break;
                }

                if (character == '\r' || character == '\n' || character == '\t')
                {
                    builder.Append(character);
                    continue;
                }

                if (!char.IsControl(character))
                {
                    builder.Append(character);
                }
            }

            return builder.ToString().Trim();
        }

        public static string SingleLine(
            string value,
            int maximumLength)
        {
            return PlainText(value, maximumLength)
                .Replace('\r', ' ')
                .Replace('\n', ' ')
                .Replace('\t', ' ')
                .Trim();
        }
    }

    // Provider-aware context scaling. The character budgets in this
    // assembly are sized for local models with modest context
    // windows; Gemini models carry ~1M-token windows, so Gemini
    // mode multiplies the TEXT budgets while capability caps
    // (message counts, attachment counts, tool rounds, byte intake)
    // stay fixed. Applied from settings at startup and on save.
    public static class ContextScale
    {
        public const int LargeContextMultiplier = 4;
        public const int MaxUserMultiplier = 8;

        private static bool _largeContext;
        private static int _userMultiplier = 1;
        private static int _multiplier = 1;

        public static int Multiplier
        {
            get { return _multiplier; }
        }

        public static void Apply(bool largeContext)
        {
            _largeContext = largeContext;
            Recompute();
        }

        // User-chosen context multiplier from the Settings Limits
        // tab; the larger of the provider multiplier and the user's
        // choice wins, clamped to MaxUserMultiplier.
        public static void ApplyUserMultiplier(int multiplier)
        {
            _userMultiplier = Math.Max(
                1,
                Math.Min(MaxUserMultiplier, multiplier));
            Recompute();
        }

        private static void Recompute()
        {
            _multiplier = Math.Max(
                _largeContext ? LargeContextMultiplier : 1,
                _userMultiplier);
        }

        public static int Scaled(int baseCharacters)
        {
            return baseCharacters * _multiplier;
        }
    }

    // Effective request limits, settable only from the Settings
    // Limits tab within the hard clamps below. Everything here is a
    // text, loop, or working-set budget - drafting and sending
    // capability rules (one draft per request, no sending, marked
    // draft surfaces only) live elsewhere and are never adjustable.
    public static class LimitOverrides
    {
        public const int MinPromptCharacters = 2000;
        public const int MaxPromptCharacters = 16000;
        public const int MinAssistantCharacters = 4000;
        public const int MaxAssistantCharactersLimit = 48000;
        public const int MinHistoryTurns = 4;
        public const int MaxHistoryTurns = 24;
        public const int MinToolRounds = 2;
        public const int MaxToolRoundsLimit = 8;
        public const int MinToolCallsPerRound = 2;
        public const int MaxToolCallsPerRoundLimit = 8;
        public const int RecommendedWorkingSetMessages = 10;
        public const int MinWorkingSetMessages = 3;
        public const int MaxWorkingSetMessages = 50;

        private static int _promptCharacters =
            TextBoundary.RecommendedUserPromptCharacters;
        private static int _assistantCharacters =
            TextBoundary.RecommendedAssistantCharacters;
        private static int _historyTurns =
            TextBoundary.RecommendedConversationTurns;
        private static int _toolRounds =
            TextBoundary.RecommendedToolRounds;
        private static int _toolCallsPerRound =
            TextBoundary.RecommendedToolCallsPerRound;
        private static int _workingSetMessages =
            RecommendedWorkingSetMessages;

        public static int PromptCharacters
        {
            get { return _promptCharacters; }
        }

        public static int AssistantCharacters
        {
            get { return _assistantCharacters; }
        }

        public static int HistoryTurns
        {
            get { return _historyTurns; }
        }

        public static int ToolRounds
        {
            get { return _toolRounds; }
        }

        public static int ToolCallsPerRound
        {
            get { return _toolCallsPerRound; }
        }

        public static int WorkingSetMessages
        {
            get { return _workingSetMessages; }
        }

        public static void Apply(
            bool useRecommended,
            int promptCharacters,
            int assistantCharacters,
            int historyTurns,
            int toolRounds,
            int toolCallsPerRound,
            int workingSetMessages)
        {
            if (useRecommended)
            {
                _promptCharacters =
                    TextBoundary.RecommendedUserPromptCharacters;
                _assistantCharacters =
                    TextBoundary.RecommendedAssistantCharacters;
                _historyTurns =
                    TextBoundary.RecommendedConversationTurns;
                _toolRounds =
                    TextBoundary.RecommendedToolRounds;
                _toolCallsPerRound =
                    TextBoundary.RecommendedToolCallsPerRound;
                _workingSetMessages =
                    RecommendedWorkingSetMessages;
                return;
            }

            _promptCharacters = Clamp(
                promptCharacters,
                MinPromptCharacters,
                MaxPromptCharacters);
            _assistantCharacters = Clamp(
                assistantCharacters,
                MinAssistantCharacters,
                MaxAssistantCharactersLimit);
            _historyTurns = Clamp(
                historyTurns,
                MinHistoryTurns,
                MaxHistoryTurns);
            _toolRounds = Clamp(
                toolRounds,
                MinToolRounds,
                MaxToolRoundsLimit);
            _toolCallsPerRound = Clamp(
                toolCallsPerRound,
                MinToolCallsPerRound,
                MaxToolCallsPerRoundLimit);
            _workingSetMessages = Clamp(
                workingSetMessages,
                MinWorkingSetMessages,
                MaxWorkingSetMessages);
        }

        private static int Clamp(
            int value,
            int minimum,
            int maximum)
        {
            return Math.Max(
                minimum,
                Math.Min(maximum, value));
        }
    }

}

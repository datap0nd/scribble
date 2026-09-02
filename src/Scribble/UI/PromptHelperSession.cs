using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Scribble.Chat;
using Scribble.Security;

namespace Scribble.UI
{
    // Bridges the suite-wide ask_user tool to the shared WebView chat UI.
    // There can be only one live question per pane. The model loop pauses
    // until a chip/free-text answer arrives or Stop cancels the request.
    internal sealed class PromptHelperSession
    {
        private readonly object _sync = new object();
        private readonly Action<IDictionary<string, object>> _post;
        private readonly JavaScriptSerializer _serializer =
            new JavaScriptSerializer();
        private TaskCompletionSource<string> _pending;
        private PromptHelperQuestion _question;

        public PromptHelperSession(
            Action<IDictionary<string, object>> post)
        {
            _post = post ?? throw new ArgumentNullException(nameof(post));
        }

        public async Task<MailboxToolResult> AskAsync(
            ChatToolCall call,
            CancellationToken cancellationToken)
        {
            PromptHelperQuestion question;
            try
            {
                question = PromptHelperTool.Parse(call);
            }
            catch (Exception exception)
            {
                return new MailboxToolResult(
                    call?.id ?? string.Empty,
                    "[ASK_USER_INVALID] " + TextBoundary.SingleLine(
                        exception.Message,
                        500),
                    "Clarification question was invalid");
            }

            TaskCompletionSource<string> pending;
            lock (_sync)
            {
                if (_pending != null)
                {
                    return new MailboxToolResult(
                        call?.id ?? string.Empty,
                        "[ASK_USER_BUSY] Another clarification question is already open.",
                        "A clarification question is already open");
                }

                pending = new TaskCompletionSource<string>(
                    TaskCreationOptions.RunContinuationsAsynchronously);
                _pending = pending;
                _question = question;
            }

            PostQuestion(question);
            string answer;
            using (cancellationToken.Register(
                () => pending.TrySetCanceled()))
            {
                try
                {
                    answer = await pending.Task;
                }
                finally
                {
                    lock (_sync)
                    {
                        if (ReferenceEquals(_pending, pending))
                        {
                            _pending = null;
                            _question = null;
                        }
                    }
                }
            }

            return new MailboxToolResult(
                call?.id ?? string.Empty,
                _serializer.Serialize(
                    new Dictionary<string, object>
                    {
                        { "answer", answer }
                    }),
                "Continuing with your answer");
        }

        public void HandleAnswer(object answerValue)
        {
            var answer = TextBoundary.SingleLine(
                Convert.ToString(answerValue) ?? string.Empty,
                PromptHelperTool.MaxAnswerCharacters);
            if (answer.Length == 0)
            {
                return;
            }

            TaskCompletionSource<string> pending;
            lock (_sync)
            {
                pending = _pending;
            }

            pending?.TrySetResult(answer);
        }

        public void Cancel()
        {
            TaskCompletionSource<string> pending;
            lock (_sync)
            {
                pending = _pending;
            }

            if (pending != null)
            {
                _post(new Dictionary<string, object>
                {
                    { "type", "dismissAskUser" }
                });
                pending.TrySetCanceled();
            }
        }

        public void RestoreIfPending()
        {
            PromptHelperQuestion question;
            lock (_sync)
            {
                question = _question;
            }

            if (question != null)
            {
                PostQuestion(question);
            }
        }

        private void PostQuestion(PromptHelperQuestion question)
        {
            var options = new List<object>();
            foreach (var option in question.Options)
            {
                options.Add(new Dictionary<string, object>
                {
                    { "label", option.Label },
                    { "description", option.Description }
                });
            }

            _post(new Dictionary<string, object>
            {
                { "type", "askUser" },
                { "question", question.Question },
                { "reason", question.Reason },
                { "options", options }
            });
        }
    }
}

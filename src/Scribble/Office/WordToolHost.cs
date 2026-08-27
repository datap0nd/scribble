using System;
using System.Collections.Generic;
using System.Globalization;
using System.Web.Script.Serialization;
using Scribble.Chat;
using Scribble.Security;
using Scribble.Utilities;

namespace Scribble.Office
{
    // Read-only Word access for the Scribble pane. Document text is
    // bounded before it reaches the model and always travels as
    // untrusted data. Draft writes live in DocumentDraftHost behind
    // the one-shot authorization.
    public sealed class WordToolHost
    {
        public const int MaxReadCharacters = 24000;

        private readonly object _wordApplication;
        private readonly JavaScriptSerializer _serializer =
            new JavaScriptSerializer();

        public WordToolHost(object wordApplication)
        {
            _wordApplication = wordApplication ??
                throw new ArgumentNullException(
                    nameof(wordApplication));
        }

        public MailboxToolResult Execute(ChatToolCall call)
        {
            if (call?.function == null ||
                string.IsNullOrWhiteSpace(call.id))
            {
                return Error(
                    call?.id,
                    "DOCUMENT_TOOL_CALL_INVALID",
                    "The model returned an invalid tool call.");
            }

            var name = call.function.name ?? string.Empty;
            if (!WordToolCatalog.IsApproved(name))
            {
                return Error(
                    call.id,
                    "DOCUMENT_TOOL_NOT_ALLOWED",
                    "The requested document tool is not allowed.");
            }

            try
            {
                var arguments = ToolArguments.Parse(
                    _serializer,
                    call.function.arguments);
                return ReadDocument(call.id, arguments);
            }
            catch (Exception exception)
            {
                Log.Error("WordTool." + name, exception);
                return Error(
                    call.id,
                    "DOCUMENT_TOOL_FAILED",
                    DiagnosticDetails.ForException(
                        exception,
                        "DOCUMENT_TOOL_FAILED"));
            }
        }

        public string DescribeActiveContext()
        {
            try
            {
                dynamic document = ActiveDocument();
                if (document == null)
                {
                    return "No document is open in Word.";
                }

                var lines = new List<string>
                {
                    "Document: " + TextBoundary.SingleLine(
                        Convert.ToString(document.Name),
                        180)
                };
                var saved = string.Empty;
                try
                {
                    saved = Convert.ToString(document.Path) ??
                        string.Empty;
                }
                catch
                {
                }

                lines.Add(saved.Length > 0
                    ? "Saved on disk: yes"
                    : "Saved on disk: no (unsaved document)");
                try
                {
                    int characters =
                        (int)document.Content.End;
                    lines.Add(
                        "Length: about " +
                        characters.ToString(
                            CultureInfo.InvariantCulture) +
                        " characters, " +
                        ((int)document.Paragraphs.Count)
                        .ToString(CultureInfo.InvariantCulture) +
                        " paragraphs");
                }
                catch
                {
                }

                return string.Join("\n", lines);
            }
            catch (Exception exception)
            {
                Log.Error("WordDescribe", exception);
                return "The document context could not be read.";
            }
        }

        // Bounded snapshot of the current selection for the context
        // tray ("add selected text").
        public string DescribeSelection(out string title)
        {
            title = "Word selection";
            dynamic application = _wordApplication;
            dynamic selection = application.Selection;
            var text = Convert.ToString(selection.Text) ??
                string.Empty;
            if (text.Trim().Length == 0)
            {
                throw new InvalidOperationException(
                    "Select some text in Word first.");
            }

            return "Selected Word text:\n" +
                TextBoundary.PlainText(
                    text,
                    ContextScale.Scaled(MaxReadCharacters));
        }

        private dynamic ActiveDocument()
        {
            dynamic application = _wordApplication;
            try
            {
                return application.ActiveDocument;
            }
            catch
            {
                return null;
            }
        }

        private MailboxToolResult ReadDocument(
            string callId,
            IDictionary<string, object> arguments)
        {
            dynamic document = ActiveDocument();
            if (document == null)
            {
                return Error(
                    callId,
                    "DOCUMENT_NOT_OPEN",
                    "No document is open in Word.");
            }

            var start = ToolArguments.GetInteger(
                arguments,
                "start",
                0,
                0,
                10000000);
            string fullText = Convert.ToString(
                document.Content.Text) ?? string.Empty;
            var totalLength = fullText.Length;
            if (start >= totalLength && totalLength > 0)
            {
                return Error(
                    callId,
                    "DOCUMENT_OFFSET_PAST_END",
                    "The start offset is past the end of the document (" +
                    totalLength.ToString(
                        CultureInfo.InvariantCulture) +
                    " characters).");
            }

            var cap = ContextScale.Scaled(MaxReadCharacters);
            var slice = fullText.Substring(
                Math.Min(start, totalLength));
            var truncated = slice.Length > cap;
            slice = TextBoundary.PlainText(slice, cap);
            var payload = new Dictionary<string, object>
            {
                { "untrusted_document_data", true },
                {
                    "document",
                    TextBoundary.SingleLine(
                        Convert.ToString(document.Name),
                        180)
                },
                { "total_characters", totalLength },
                { "start", start },
                { "truncated", truncated },
                { "text", slice }
            };
            var json = _serializer.Serialize(payload);
            if (json.Length >
                ContextScale.Scaled(
                    TextBoundary.MaxToolResultCharacters))
            {
                return Error(
                    callId,
                    "DOCUMENT_TOOL_RESULT_TOO_LARGE",
                    "The bounded document result was still too large to return safely.");
            }

            return new MailboxToolResult(
                callId,
                json,
                "Read " +
                slice.Length.ToString(
                    CultureInfo.InvariantCulture) +
                " document characters from offset " +
                start.ToString(CultureInfo.InvariantCulture) +
                ".");
        }

        private MailboxToolResult Error(
            string callId,
            string code,
            string message)
        {
            return new MailboxToolResult(
                callId ?? string.Empty,
                _serializer.Serialize(
                    new Dictionary<string, object>
                    {
                        { "ok", false },
                        { "error_code", code },
                        {
                            "message",
                            TextBoundary.PlainText(message, 1200)
                        }
                    }),
                "[" + code + "] " + message);
        }
    }
}

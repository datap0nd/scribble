using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Web.Script.Serialization;
using Scribble.Chat;
using Scribble.Security;
using Scribble.Utilities;

namespace Scribble.Office
{
    // Read-only PowerPoint access for the Scribble pane. Slide and
    // notes text is bounded before it reaches the model and always
    // travels as untrusted data. Draft writes live in
    // DocumentDraftHost behind the one-shot authorization.
    public sealed class PresentationToolHost
    {
        public const int MaxSlides = 100;
        public const int MaxSlideCharacters = 8000;
        public const int MaxPreviewCharacters = 240;

        private readonly object _powerPointApplication;
        private readonly JavaScriptSerializer _serializer =
            new JavaScriptSerializer();

        public PresentationToolHost(object powerPointApplication)
        {
            _powerPointApplication = powerPointApplication ??
                throw new ArgumentNullException(
                    nameof(powerPointApplication));
        }

        public MailboxToolResult Execute(ChatToolCall call)
        {
            if (call?.function == null ||
                string.IsNullOrWhiteSpace(call.id))
            {
                return Error(
                    call?.id,
                    "PRESENTATION_TOOL_CALL_INVALID",
                    "The model returned an invalid tool call.");
            }

            var name = call.function.name ?? string.Empty;
            if (!PresentationToolCatalog.IsApproved(name))
            {
                return Error(
                    call.id,
                    "PRESENTATION_TOOL_NOT_ALLOWED",
                    "The requested presentation tool is not allowed.");
            }

            try
            {
                var arguments = ToolArguments.Parse(
                    _serializer,
                    call.function.arguments);
                switch (name)
                {
                    case PresentationToolCatalog.ListSlides:
                        return ListSlides(call.id);
                    case PresentationToolCatalog.ReadSlide:
                        return ReadSlide(call.id, arguments);
                    default:
                        return Error(
                            call.id,
                            "PRESENTATION_TOOL_NOT_ALLOWED",
                            "The requested presentation tool is not allowed.");
                }
            }
            catch (Exception exception)
            {
                Log.Error("PresentationTool." + name, exception);
                return Error(
                    call.id,
                    "PRESENTATION_TOOL_FAILED",
                    DiagnosticDetails.ForException(
                        exception,
                        "PRESENTATION_TOOL_FAILED"));
            }
        }

        public string DescribeActiveContext()
        {
            try
            {
                dynamic presentation = ActivePresentation();
                if (presentation == null)
                {
                    return "No presentation is open in PowerPoint.";
                }

                var lines = new List<string>
                {
                    "Presentation: " + TextBoundary.SingleLine(
                        Convert.ToString(presentation.Name),
                        180),
                    "Slides: " +
                    ((int)presentation.Slides.Count).ToString(
                        CultureInfo.InvariantCulture)
                };
                var saved = string.Empty;
                try
                {
                    saved = Convert.ToString(presentation.Path) ??
                        string.Empty;
                }
                catch
                {
                }

                lines.Add(saved.Length > 0
                    ? "Saved on disk: yes"
                    : "Saved on disk: no (unsaved presentation)");
                try
                {
                    dynamic application = _powerPointApplication;
                    int current = (int)application.ActiveWindow
                        .View.Slide.SlideIndex;
                    lines.Add("Current slide: " + current);
                }
                catch
                {
                }

                return string.Join("\n", lines);
            }
            catch (Exception exception)
            {
                Log.Error("PresentationDescribe", exception);
                return "The presentation context could not be read.";
            }
        }

        // Bounded snapshot of the current slide for the context tray
        // ("add current slide").
        public string DescribeCurrentSlide(out string title)
        {
            title = "PowerPoint slide";
            dynamic application = _powerPointApplication;
            dynamic slide = application.ActiveWindow.View.Slide;
            if (slide == null)
            {
                throw new InvalidOperationException(
                    "Open a slide in PowerPoint first.");
            }

            int index = (int)slide.SlideIndex;
            title = "PowerPoint slide " + index;
            return "Current PowerPoint slide " + index + ":\n" +
                CollectSlideText(slide, true);
        }

        private dynamic ActivePresentation()
        {
            dynamic application = _powerPointApplication;
            try
            {
                return application.ActivePresentation;
            }
            catch
            {
                return null;
            }
        }

        private MailboxToolResult ListSlides(string callId)
        {
            dynamic presentation = ActivePresentation();
            if (presentation == null)
            {
                return Error(
                    callId,
                    "PRESENTATION_NOT_OPEN",
                    "No presentation is open in PowerPoint.");
            }

            var slides = new List<object>();
            var total = (int)presentation.Slides.Count;
            var count = Math.Min(total, MaxSlides);
            for (var index = 1; index <= count; index++)
            {
                dynamic slide = presentation.Slides[index];
                var text = CollectSlideText(slide, false);
                var slideTitle = string.Empty;
                try
                {
                    slideTitle = Convert.ToString(
                        slide.Shapes.Title.TextFrame
                            .TextRange.Text) ?? string.Empty;
                }
                catch
                {
                }

                slides.Add(new Dictionary<string, object>
                {
                    { "index", index },
                    {
                        "title",
                        TextBoundary.SingleLine(slideTitle, 200)
                    },
                    {
                        "preview",
                        TextBoundary.SingleLine(
                            text,
                            MaxPreviewCharacters)
                    }
                });
            }

            return Success(
                callId,
                new Dictionary<string, object>
                {
                    { "untrusted_document_data", true },
                    {
                        "presentation",
                        TextBoundary.SingleLine(
                            Convert.ToString(presentation.Name),
                            180)
                    },
                    { "slide_count", total },
                    { "listed", count },
                    { "slides", slides }
                },
                "Listed " +
                count.ToString(CultureInfo.InvariantCulture) +
                " slides.");
        }

        private MailboxToolResult ReadSlide(
            string callId,
            IDictionary<string, object> arguments)
        {
            dynamic presentation = ActivePresentation();
            if (presentation == null)
            {
                return Error(
                    callId,
                    "PRESENTATION_NOT_OPEN",
                    "No presentation is open in PowerPoint.");
            }

            var total = (int)presentation.Slides.Count;
            var index = ToolArguments.GetInteger(
                arguments,
                "index",
                0,
                1,
                1000);
            if (index < 1 || index > total)
            {
                return Error(
                    callId,
                    "PRESENTATION_SLIDE_UNKNOWN",
                    "The slide index is out of range. Call list_slides first.");
            }

            dynamic slide = presentation.Slides[index];
            return Success(
                callId,
                new Dictionary<string, object>
                {
                    { "untrusted_document_data", true },
                    { "index", index },
                    {
                        "text",
                        TextBoundary.PlainText(
                            CollectSlideText(slide, true),
                            ContextScale.Scaled(
                                MaxSlideCharacters))
                    }
                },
                "Read slide " +
                index.ToString(CultureInfo.InvariantCulture) +
                ".");
        }

        // Gathers text from every shape on the slide, and the
        // speaker notes when includeNotes is set, bounded.
        private static string CollectSlideText(
            dynamic slide,
            bool includeNotes)
        {
            var builder = new StringBuilder();
            try
            {
                foreach (dynamic shape in slide.Shapes)
                {
                    AppendShapeText(builder, shape);
                    if (builder.Length >
                        ContextScale.Scaled(MaxSlideCharacters))
                    {
                        break;
                    }
                }
            }
            catch
            {
            }

            if (includeNotes)
            {
                try
                {
                    var notes = new StringBuilder();
                    foreach (dynamic shape in
                        slide.NotesPage.Shapes)
                    {
                        AppendShapeText(notes, shape);
                    }

                    if (notes.Length > 0)
                    {
                        builder.Append("\n[Speaker notes]\n");
                        builder.Append(notes);
                    }
                }
                catch
                {
                }
            }

            return TextBoundary.PlainText(
                builder.ToString(),
                ContextScale.Scaled(MaxSlideCharacters));
        }

        private static void AppendShapeText(
            StringBuilder builder,
            dynamic shape)
        {
            try
            {
                if ((int)shape.HasTextFrame == 0)
                {
                    return;
                }

                var text = Convert.ToString(
                    shape.TextFrame.TextRange.Text) ??
                    string.Empty;
                if (text.Trim().Length > 0)
                {
                    if (builder.Length > 0)
                    {
                        builder.Append('\n');
                    }

                    builder.Append(text.Trim());
                }
            }
            catch
            {
            }
        }

        private MailboxToolResult Success(
            string callId,
            object payload,
            string status)
        {
            var json = _serializer.Serialize(payload);
            if (json.Length >
                ContextScale.Scaled(
                    TextBoundary.MaxToolResultCharacters))
            {
                return Error(
                    callId,
                    "PRESENTATION_TOOL_RESULT_TOO_LARGE",
                    "The bounded presentation result was still too large to return safely.");
            }

            return new MailboxToolResult(callId, json, status);
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

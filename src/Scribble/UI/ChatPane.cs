/*
THESIS: A restrained Outlook sidebar makes mailbox retrieval and one linked,
human-reviewed draft visible without granting send capability.
OWN-WORLD: A dark web chat surface rendered by WebView2 from an embedded,
network-isolated page; every piece of model or mailbox text enters the DOM
as inert text nodes, never as HTML. The C# host keeps every capability
boundary; the page is presentation only.
STORY: Ask the mailbox, watch slim activity lines record what loaded, then
deliberately open an unsent draft.
*/

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Scribble.Chat;
using Scribble.Configuration;
using Scribble.Outlook;
using Scribble.Security;
using Scribble.Utilities;

namespace Scribble.UI
{
    [ComVisible(true)]
    [Guid("14D24FA1-4342-442F-B68B-B68D7372794C")]
    [ProgId("Scribble.ChatPane")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class ChatPane : UserControl
    {
        private const int MaxTranscriptEvents = 400;
        private const int MaxExternalImages = 4;

        private sealed class ExternalImageContext
        {
            public ExternalImageContext(
                VisionImagePayload payload,
                string thumbnail)
            {
                Payload = payload;
                Thumbnail = thumbnail ?? string.Empty;
            }

            public VisionImagePayload Payload { get; }

            public string Thumbnail { get; }
        }

        // Wraps an external document with the tray-chip state so files
        // that hit a cap (too large, unsupported, truncated) render as
        // amber warning chips while the bounded text still reaches the
        // model.
        private sealed class ExternalDocumentContext
        {
            public ExternalDocumentContext(
                ExternalContextDocument document,
                bool warn,
                string subtitle)
            {
                Document = document;
                Warn = warn;
                Subtitle = subtitle ?? string.Empty;
            }

            public ExternalContextDocument Document { get; }

            public bool Warn { get; }

            public string Subtitle { get; }
        }

        private readonly SettingsStore _settingsStore =
            new SettingsStore();
        private readonly OpenAiCompatibleClient _client =
            new OpenAiCompatibleClient();
        private readonly JavaScriptSerializer _serializer =
            new JavaScriptSerializer();
        private readonly List<ChatTurn> _history;
        private readonly List<MessageSnapshot> _workingMessages =
            new List<MessageSnapshot>();
        private readonly List<ExternalDocumentContext> _externalContext =
            new List<ExternalDocumentContext>();
        private readonly List<ExternalImageContext> _externalImages =
            new List<ExternalImageContext>();
        private string _pendingSuggestJson;
        private readonly System.Windows.Forms.Timer _elapsedTimer =
            new System.Windows.Forms.Timer
            {
                Interval = 5000
            };
        private DateTime _requestStartedAt = DateTime.UtcNow;
        private DateTime _statusChangedAt = DateTime.UtcNow;
        private readonly List<string> _transcriptEvents;
        private readonly PaneMemory.Slot _memory;
        private readonly DiagnosticsRecorder _diagnostics =
            new DiagnosticsRecorder();
        private readonly WebView2 _webView = new WebView2();

        private object _outlookApplication;
        private AppSettings _settings;
        private MessageSnapshot _selectedMessage;
        private DraftToolHost _draftTools;
        private Scribble.Office.DocumentDraftHost
            _crossAppTools;
        private McpToolHost _mcpTools;
        private CancellationTokenSource _requestCancellation;
        private string _lastAssistantText = string.Empty;
        // Incremented when a request starts and when the user stops:
        // an in-flight continuation compares its captured value and
        // discards itself when stale, so Stop always releases the UI
        // immediately even if the underlying HTTP call never returns.
        private int _requestGeneration;
        private bool _busy;
        private bool _shutdown;
        private bool _webReady;
        private string _scopeText =
            "No context - use /search or select emails";
        private string _statusText = "Ready";
        private bool _statusError;

        public ChatPane()
        {
            LastCreated = this;
            // Reopening the pane in the same Outlook session picks
            // the conversation back up from process memory.
            _memory = PaneMemory.For("outlook");
            _history = _memory.History;
            _transcriptEvents = _memory.Transcript;
            _lastAssistantText = _memory.LastAnswer;
            _settings = _settingsStore.Load();
            ContextScale.Apply(
                GeminiCodeAssistGateway.IsGeminiModel(
                    _settings.Model));
            _settings.ApplyLimits();
            _mcpTools = new McpToolHost(_settings.McpServers);
            // Surface silent gateway waits (quota retries) so a slow
            // response is never a mystery.
            _client.GeminiGateway.StatusListener =
                message => SetStatus(message, false);
            _elapsedTimer.Tick += ElapsedTick;
            _elapsedTimer.Start();

            Dock = DockStyle.Fill;
            BackColor = Color.FromArgb(26, 27, 30);
            MinimumSize = new Size(300, 480);
            AllowDrop = true;
            DragEnter += ChatPaneDragEnter;
            DragDrop += ChatPaneDragDrop;

            _webView.Dock = DockStyle.Fill;
            _webView.DefaultBackgroundColor =
                Color.FromArgb(26, 27, 30);
            Controls.Add(_webView);
            InitializeWebView();
        }

        internal static ChatPane LastCreated { get; private set; }

        internal void Initialize(
            object outlookApplication,
            bool refreshSelection = true)
        {
            if (_outlookApplication != null)
            {
                return;
            }

            _outlookApplication = outlookApplication ??
                throw new ArgumentNullException(nameof(outlookApplication));
            _draftTools = new DraftToolHost(
                _outlookApplication);
            _crossAppTools =
                new Scribble.Office.DocumentDraftHost(
                    "outlook",
                    _outlookApplication);
            if (refreshSelection)
            {
                RefreshSelectedMessage();
            }
            else
            {
                SetScopeUnavailable(
                    "No context - drag emails here, or use " +
                    "Add email or /search");
            }

            UpdateDraftState();
        }

        // ------------------------------------------------------------------
        // WebView2 hosting: the embedded page is the only content ever
        // loaded, remote navigation is cancelled, and script has no access
        // to anything but the JSON bridge below.
        // ------------------------------------------------------------------

        private async void InitializeWebView()
        {
            try
            {
                var dataFolder = Path.Combine(
                    Environment.GetFolderPath(
                        Environment.SpecialFolder
                            .LocalApplicationData),
                    "MetoAI",
                    "WebView2");
                var environment =
                    await CoreWebView2Environment.CreateAsync(
                        null,
                        dataFolder);
                await _webView.EnsureCoreWebView2Async(environment);
                if (_shutdown)
                {
                    return;
                }

                var settings = _webView.CoreWebView2.Settings;
                // Right-click works: copy, paste, select all, and
                // the usual editing actions. Devtools stay off and
                // the page still cannot navigate, so the menu adds
                // convenience without widening the boundary.
                settings.AreDefaultContextMenusEnabled = true;
                settings.AreDevToolsEnabled = false;
                settings.IsStatusBarEnabled = false;
                settings.AreBrowserAcceleratorKeysEnabled = false;
                settings.IsBuiltInErrorPageEnabled = false;
                settings.IsZoomControlEnabled = false;
                // External drop stays enabled so the page can receive
                // drag gestures; the page never navigates on drop and
                // forwards the gesture to the host instead.
                _webView.AllowExternalDrop = true;
                _webView.CoreWebView2.NavigationStarting +=
                    WebNavigationStarting;
                _webView.CoreWebView2.NewWindowRequested +=
                    (sender, eventArgs) => eventArgs.Handled = true;
                _webView.CoreWebView2.WebMessageReceived +=
                    WebMessageReceived;
                _webView.CoreWebView2.NavigateToString(
                    LoadChatPage());
            }
            catch (Exception exception)
            {
                Log.Error("WebViewInit", exception);
                ShowWebViewFallback(exception);
            }
        }

        private void ShowWebViewFallback(Exception exception)
        {
            try
            {
                Controls.Remove(_webView);
            }
            catch
            {
            }

            var notice = new Label
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(26, 27, 30),
                ForeColor = Color.FromArgb(232, 232, 236),
                Padding = new Padding(18),
                Text =
                    "Scribble needs the Microsoft Edge WebView2 " +
                    "runtime, which ships with Windows 10/11 and " +
                    "Microsoft Edge. Install it from Microsoft, then " +
                    "restart Outlook.\r\n\r\nDetails: " +
                    TextBoundary.SingleLine(
                        exception?.Message,
                        300)
            };
            Controls.Add(notice);
        }

        private static string LoadChatPage()
        {
            using (var stream = typeof(ChatPane).Assembly
                .GetManifestResourceStream(
                    "Scribble.UI.ChatPaneWeb.html"))
            {
                if (stream == null)
                {
                    throw new InvalidOperationException(
                        "The embedded chat page is missing.");
                }

                using (var reader = new StreamReader(
                    stream,
                    System.Text.Encoding.UTF8))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        private static void WebNavigationStarting(
            object sender,
            CoreWebView2NavigationStartingEventArgs eventArgs)
        {
            var uri = eventArgs.Uri ?? string.Empty;
            if (!uri.StartsWith(
                    "data:",
                    StringComparison.OrdinalIgnoreCase) &&
                !uri.StartsWith(
                    "about:",
                    StringComparison.OrdinalIgnoreCase))
            {
                eventArgs.Cancel = true;
            }
        }

        private void WebMessageReceived(
            object sender,
            CoreWebView2WebMessageReceivedEventArgs eventArgs)
        {
            try
            {
                var json = eventArgs.TryGetWebMessageAsString();
                var message = _serializer.DeserializeObject(json)
                    as IDictionary<string, object>;
                if (message == null)
                {
                    return;
                }

                object typeValue;
                message.TryGetValue("type", out typeValue);
                var type = Convert.ToString(typeValue) ??
                    string.Empty;
                switch (type)
                {
                    case "ready":
                        HandleWebReady();
                        break;
                    case "send":
                        object textValue;
                        message.TryGetValue("text", out textValue);
                        HandleSendMessage(
                            Convert.ToString(textValue) ??
                            string.Empty);
                        break;
                    case "stop":
                        HandleStop();
                        break;
                    case "newChat":
                        HandleNewChat();
                        break;
                    case "addEmail":
                        AddActiveSelection();
                        break;
                    case "addFiles":
                        HandleAddFiles();
                        break;
                    case "openSettings":
                        OpenSettings();
                        break;
                    case "clearContext":
                        HandleClearContext();
                        break;
                    case "removeContext":
                        object kindValue;
                        object indexValue;
                        message.TryGetValue("kind", out kindValue);
                        message.TryGetValue("index", out indexValue);
                        int removeIndex;
                        int.TryParse(
                            Convert.ToString(indexValue),
                            out removeIndex);
                        HandleRemoveContext(
                            Convert.ToString(kindValue) ??
                            string.Empty,
                            removeIndex);
                        break;
                    case "emailDrop":
                        if (!_busy)
                        {
                            AddActiveSelection();
                        }

                        break;
                    case "fileDrop":
                        HandleWebFileDrop(eventArgs);
                        break;
                    case "setModel":
                        object modelValue;
                        message.TryGetValue("model", out modelValue);
                        HandleSetModel(
                            Convert.ToString(modelValue) ??
                            string.Empty);
                        break;
                    case "suggestAnswers":
                        object answersValue;
                        message.TryGetValue(
                            "answers",
                            out answersValue);
                        HandleSuggestAnswers(answersValue);
                        break;
                    case "shareAnswer":
                        HandleShareAnswer();
                        break;
                    case "addShared":
                        HandleAddShared();
                        break;
                    case "copyDiag":
                        HandleCopyDiagnostics();
                        break;
                }
            }
            catch (Exception exception)
            {
                Log.Error("WebMessage", exception);
            }
        }

        private void HandleWebReady()
        {
            _webReady = true;
            RefreshModelPicker();
            PostToWeb(new Dictionary<string, object>
            {
                { "type", "scope" },
                { "text", _scopeText }
            });
            UpdateDraftState();
            PushContextToWeb();
            PostToWeb(new Dictionary<string, object>
            {
                { "type", "clear" }
            });
            foreach (var recorded in _transcriptEvents.ToArray())
            {
                PostRawToWeb(recorded);
            }

            if (_pendingSuggestJson != null)
            {
                PostRawToWeb(_pendingSuggestJson);
                _pendingSuggestJson = null;
            }

            PostToWeb(new Dictionary<string, object>
            {
                { "type", "busy" },
                { "value", _busy }
            });
            PostToWeb(new Dictionary<string, object>
            {
                { "type", "status" },
                { "text", _statusText },
                { "error", _statusError }
            });
        }

        private void PostToWeb(IDictionary<string, object> payload)
        {
            PostRawToWeb(_serializer.Serialize(payload));
        }

        private void PostRawToWeb(string json)
        {
            if (!_webReady || _shutdown)
            {
                return;
            }

            try
            {
                _webView.CoreWebView2?.PostWebMessageAsJson(json);
            }
            catch (Exception exception)
            {
                Log.Error("PostToWeb", exception);
            }
        }

        // Live token streaming: deltas render in a transient bubble
        // on the page and are never recorded in the transcript; the
        // final formatted assistant message replaces the bubble.
        private void PostStreamDelta(string text)
        {
            if (string.IsNullOrEmpty(text) || !_busy)
            {
                return;
            }

            PostToWeb(new Dictionary<string, object>
            {
                { "type", "delta" },
                { "text", text }
            });
        }

        private void PostStreamEnd()
        {
            PostToWeb(new Dictionary<string, object>
            {
                { "type", "deltaEnd" }
            });
        }

        private void PostTranscript(
            IDictionary<string, object> payload)
        {
            var json = _serializer.Serialize(payload);
            _transcriptEvents.Add(json);
            if (_transcriptEvents.Count > MaxTranscriptEvents)
            {
                _transcriptEvents.RemoveAt(0);
            }

            PostRawToWeb(json);
        }

        // ------------------------------------------------------------------
        // Transcript primitives. Model and mailbox text crosses the bridge
        // as plain strings; the page inserts it via textContent only.
        // ------------------------------------------------------------------

        private void AppendUserTurn(string text)
        {
            PostTranscript(new Dictionary<string, object>
            {
                { "type", "user" },
                { "text", text }
            });
        }

        private void AppendFormattedAssistantText(string text)
        {
            var formatted = SafeModelText.Format(
                text,
                TextBoundary.MaxAssistantCharacters);
            var ranges = new List<object>();
            foreach (var range in formatted.BoldRanges)
            {
                ranges.Add(new Dictionary<string, object>
                {
                    { "s", range.Start },
                    { "l", range.Length }
                });
            }

            // The page renders the bounded raw markdown (tables,
            // lists, code) as safe DOM nodes; the plain text plus
            // bold ranges remain the fallback path.
            PostTranscript(new Dictionary<string, object>
            {
                { "type", "assistant" },
                { "text", formatted.PlainText },
                { "bold", ranges },
                {
                    "md",
                    TextBoundary.PlainText(
                        text,
                        TextBoundary.MaxAssistantCharacters)
                }
            });
        }

        private void AppendContext(string text)
        {
            PostTranscript(new Dictionary<string, object>
            {
                { "type", "activity" },
                { "text", TextBoundary.SingleLine(text, 400) },
                { "kind", "context" }
            });
        }

        private void AppendDraftAction(string text)
        {
            PostTranscript(new Dictionary<string, object>
            {
                { "type", "activity" },
                { "text", TextBoundary.SingleLine(text, 400) },
                { "kind", "draft" }
            });
        }

        private void AppendError(string text)
        {
            PostTranscript(new Dictionary<string, object>
            {
                { "type", "activity" },
                { "text", TextBoundary.PlainText(text, 2400) },
                { "kind", "error" }
            });
        }

        private void SetStatus(string text, bool error)
        {
            _statusText = TextBoundary.SingleLine(text, 300);
            _statusError = error;
            _statusChangedAt = DateTime.UtcNow;
            PostToWeb(new Dictionary<string, object>
            {
                { "type", "status" },
                { "text", _statusText },
                { "error", error }
            });
        }

        // While a request runs, a ticker appends elapsed time to a
        // status that has not changed recently, so a long wait is
        // always visibly progressing instead of looking frozen.
        private void ElapsedTick(object sender, EventArgs eventArgs)
        {
            if (!_busy)
            {
                return;
            }

            var sinceChange =
                DateTime.UtcNow - _statusChangedAt;
            if (sinceChange.TotalSeconds < 8)
            {
                return;
            }

            var elapsed = (int)(DateTime.UtcNow - _requestStartedAt)
                .TotalSeconds;
            PostToWeb(new Dictionary<string, object>
            {
                { "type", "status" },
                {
                    "text",
                    _statusText + " (" + elapsed + "s)"
                },
                { "error", _statusError }
            });
        }

        private void SetScope(string text)
        {
            _scopeText = TextBoundary.SingleLine(text, 200);
            PostToWeb(new Dictionary<string, object>
            {
                { "type", "scope" },
                { "text", _scopeText }
            });
        }

        private void SetScopeUnavailable(string text)
        {
            SetScope(text);
        }

        private void SetBusy(bool busy)
        {
            _busy = busy;
            PostToWeb(new Dictionary<string, object>
            {
                { "type", "busy" },
                { "value", busy }
            });
            if (busy)
            {
                SetStatus("Thinking...", false);
            }
        }

        private void UpdateDraftState()
        {
            var linked =
                _draftTools != null &&
                _draftTools.HasActiveDraft;
            PostToWeb(new Dictionary<string, object>
            {
                { "type", "draft" },
                {
                    "text",
                    linked
                        ? "Draft linked - feedback updates it. Scribble cannot send."
                        : "Say 'create a draft' to open one. Scribble cannot send."
                },
                { "linked", linked }
            });
        }

        private void RefreshModelPicker()
        {
            var saved = (_settings?.Model ?? string.Empty).Trim();
            var current = ModelSelectionPolicy.IsGenerativeModel(saved)
                ? saved
                : string.Empty;
            var models = new List<string>(
                _settings?.DiscoveredModels ?? new List<string>());
            if (current.Length > 0 &&
                models.FindIndex(model =>
                    string.Equals(
                        model,
                        current,
                        StringComparison.OrdinalIgnoreCase)) < 0)
            {
                models.Insert(0, current);
            }

            var items = new List<object>();
            foreach (var model in models)
            {
                if (!ModelSelectionPolicy.IsGenerativeModel(model))
                {
                    continue;
                }

                items.Add(new Dictionary<string, object>
                {
                    { "id", model },
                    {
                        "vision",
                        ModelCatalog.IsVisionCapable(model)
                    }
                });
            }

            PostToWeb(new Dictionary<string, object>
            {
                { "type", "models" },
                { "items", items },
                { "current", current }
            });
        }

        private void PushContextToWeb()
        {
            var items = new List<object>();
            if (_selectedMessage != null)
            {
                var selectedCard =
                    (Dictionary<string, object>)
                    BuildWorkingSetCard(0, _selectedMessage);
                selectedCard["kind"] = "selected";
                selectedCard["index"] = 0;
                selectedCard["badge"] = "@";
                items.Add(selectedCard);
            }

            for (var index = 0;
                 index < _workingMessages.Count;
                 index++)
            {
                var card =
                    (Dictionary<string, object>)
                    BuildWorkingSetCard(
                        index,
                        _workingMessages[index]);
                card["kind"] = "email";
                card["index"] = index;
                items.Add(card);
            }

            for (var index = 0;
                 index < _externalContext.Count;
                 index++)
            {
                var card =
                    (Dictionary<string, object>)
                    BuildExternalContextCard(
                        _externalContext[index]);
                card["kind"] = "file";
                card["index"] = index;
                items.Add(card);
            }

            for (var index = 0;
                 index < _externalImages.Count;
                 index++)
            {
                var image = _externalImages[index];
                items.Add(new Dictionary<string, object>
                {
                    { "kind", "image" },
                    { "index", index },
                    {
                        "title",
                        TextBoundary.SingleLine(
                            image.Payload.FileName,
                            120)
                    },
                    { "subtitle", "image - vision input" },
                    { "thumb", image.Thumbnail }
                });
            }

            PostToWeb(new Dictionary<string, object>
            {
                { "type", "context" },
                { "items", items }
            });
        }

        private void HandleRemoveContext(string kind, int index)
        {
            if (_busy)
            {
                return;
            }

            switch (kind)
            {
                case "selected":
                    _selectedMessage = null;
                    break;
                case "email":
                    if (index >= 0 &&
                        index < _workingMessages.Count)
                    {
                        _workingMessages.RemoveAt(index);
                    }

                    break;
                case "file":
                    if (index >= 0 &&
                        index < _externalContext.Count)
                    {
                        _externalContext.RemoveAt(index);
                    }

                    break;
                case "image":
                    if (index >= 0 &&
                        index < _externalImages.Count)
                    {
                        _externalImages.RemoveAt(index);
                    }

                    break;
            }

            if (_selectedMessage == null &&
                _workingMessages.Count == 0)
            {
                SetScopeUnavailable(
                    "No context - use /search or select emails");
            }
            else if (_workingMessages.Count > 0)
            {
                SetScope(
                    "Working set: " +
                    _workingMessages.Count +
                    " of " +
                    MailboxWorkingSet.MaxMessages +
                    " emails");
            }

            RefreshContextLayer("External files");
            SetStatus("Removed from context", false);
        }

        private object BuildWorkingSetCard(
            int index,
            MessageSnapshot message)
        {
            var subject = TextBoundary.SingleLine(
                SubjectDisplay.Clean(message.Subject),
                180);
            if (subject.Length == 0)
            {
                subject = "(No subject)";
            }

            var sender = TextBoundary.SingleLine(
                message.Sender,
                120);
            if (sender.Length == 0)
            {
                sender = "Unknown sender";
            }

            var date = message.ReceivedAt?.ToString(
                "yyyy-MM-dd HH:mm") ??
                "Unknown date";
            return new Dictionary<string, object>
            {
                { "badge", (index + 1).ToString() },
                { "title", subject },
                { "subtitle", sender + "  |  " + date }
            };
        }

        private object BuildExternalContextCard(
            ExternalDocumentContext entry)
        {
            var subtitle = entry.Subtitle.Length > 0
                ? entry.Subtitle
                : entry.Document.Content.Length + " text characters";
            var card = new Dictionary<string, object>
            {
                { "badge", entry.Warn ? "!" : "F" },
                {
                    "title",
                    TextBoundary.SingleLine(entry.Document.Name, 180)
                },
                { "subtitle", subtitle }
            };
            if (entry.Warn)
            {
                card["warn"] = true;
            }

            return card;
        }

        private void RefreshContextLayer(string source)
        {
            PushContextToWeb();
        }

        // ------------------------------------------------------------------
        // Context management (unchanged capability boundaries).
        // ------------------------------------------------------------------

        // "Suggest a response" from the Outlook context menu: capture
        // the right-clicked email, ask up to three short questions in
        // the pane (tone plus model-suggested specifics), then feed
        // the answers into the normal drafting pipeline. The composed
        // prompt goes through HandleSendMessage, so every existing
        // boundary applies: explicit drafting language authorizes one
        // draft creation and nothing else changes.
        public void BeginSuggestResponse()
        {
            if (_busy)
            {
                SetStatus(
                    "Scribble is busy - try again in a moment",
                    true);
                return;
            }

            RefreshSelectedMessage();
            if (_selectedMessage == null)
            {
                SetStatus(
                    "Select an email first, then use Suggest a " +
                    "response",
                    true);
                return;
            }

            RunSuggestQuestionFlow();
        }

        private async void RunSuggestQuestionFlow()
        {
            SetBusy(true);
            _requestStartedAt = DateTime.UtcNow;
            SetStatus("Preparing reply questions...", false);
            var generation = ++_requestGeneration;
            var cancellation = new CancellationTokenSource();
            _requestCancellation = cancellation;
            // The generated questions are a nicety, so this call is
            // both stoppable and time-capped: a stalled endpoint
            // falls back to the tone question instead of wedging.
            cancellation.CancelAfter(TimeSpan.FromSeconds(45));
            var questions = new List<object>
            {
                new Dictionary<string, object>
                {
                    {
                        "text",
                        "What tone should the reply take?"
                    },
                    {
                        "options",
                        new List<string>
                        {
                            "Professional",
                            "Friendly",
                            "Brief and direct",
                            "Warm and personal"
                        }
                    }
                }
            };
            try
            {
                if (_settings != null && _settings.IsConfigured)
                {
                    var model = ModelRouting.ResolveForRequest(
                        _settings,
                        false);
                    var request =
                        SuggestQuestionsRequestFactory.Create(
                            model,
                            _selectedMessage);
                    var response = await _client.CompleteAsync(
                        _settings,
                        request,
                        cancellation.Token);
                    var parsed =
                        SuggestQuestionsRequestFactory.Parse(
                            response?.content ?? string.Empty);
                    foreach (var question in parsed)
                    {
                        questions.Add(
                            new Dictionary<string, object>
                            {
                                { "text", question.Text },
                                {
                                    "options",
                                    new List<string>(
                                        question.Options)
                                }
                            });
                    }
                }
            }
            catch (OperationCanceledException)
            {
                // User stop is handled below via the generation
                // check; a timeout continues with the tone question.
            }
            catch (Exception exception)
            {
                // The tone question alone still gives a useful flow.
                Log.Error("SuggestQuestions", exception);
            }
            finally
            {
                if (ReferenceEquals(
                    _requestCancellation,
                    cancellation))
                {
                    _requestCancellation = null;
                }

                cancellation.Dispose();
            }

            if (generation != _requestGeneration)
            {
                // The user pressed stop; HandleStop already reset
                // the UI, so the question card is not shown.
                return;
            }

            SetBusy(false);
            SetStatus(
                "Answer the questions to shape the draft",
                false);
            var payload = new Dictionary<string, object>
            {
                { "type", "suggest" },
                { "questions", questions }
            };
            if (_webReady)
            {
                PostToWeb(payload);
            }
            else
            {
                _pendingSuggestJson =
                    _serializer.Serialize(payload);
            }
        }

        private void HandleSuggestAnswers(object answersValue)
        {
            if (_busy)
            {
                return;
            }

            var hasDraft = _draftTools != null &&
                _draftTools.HasActiveDraft;
            var lines = new List<string>
            {
                hasDraft
                    ? "Rewrite the linked draft as a reply to " +
                      "the selected email."
                    : "Draft a reply to the selected email."
            };
            var answers = answersValue as object[];
            if (answers != null)
            {
                var count = 0;
                foreach (var entry in answers)
                {
                    if (count == 3)
                    {
                        break;
                    }

                    var map = entry as IDictionary<string, object>;
                    if (map == null)
                    {
                        continue;
                    }

                    object questionValue;
                    object answerValue;
                    map.TryGetValue("question", out questionValue);
                    map.TryGetValue("answer", out answerValue);
                    var question = TextBoundary.SingleLine(
                        Convert.ToString(questionValue) ??
                        string.Empty,
                        200);
                    var answer = TextBoundary.SingleLine(
                        Convert.ToString(answerValue) ??
                        string.Empty,
                        300);
                    if (answer.Length == 0)
                    {
                        continue;
                    }

                    lines.Add(question.Length > 0
                        ? question + " " + answer
                        : "Guidance: " + answer);
                    count++;
                }
            }

            HandleSendMessage(string.Join("\n", lines));
        }

        public void RefreshSelectedMessage()
        {
            if (_outlookApplication == null)
            {
                SetScopeUnavailable(
                    "Outlook is still initializing");
                return;
            }

            if (_busy)
            {
                SetStatus(
                    "Still working - try again in a moment",
                    true);
                return;
            }

            try
            {
                SetSelectedMessage(
                    new MessageReader(_outlookApplication)
                        .CaptureCurrent());
                SetStatus("Email selected", false);
            }
            catch (Exception exception)
            {
                _workingMessages.Clear();
                RefreshContextLayer("External files");
                _selectedMessage = null;
                SetScopeUnavailable(
                    "No context - use /search or select emails");
                SetStatus("Ready", false);
                Log.Error("CaptureCurrent", exception);
            }
        }

        public void UseRibbonSelection(object selection)
        {
            if (_outlookApplication == null)
            {
                return;
            }

            if (_busy)
            {
                SetStatus(
                    "Still working - try again in a moment",
                    true);
                return;
            }

            try
            {
                var reader = new MessageReader(
                    _outlookApplication);
                IReadOnlyList<MessageSnapshot> messages;
                try
                {
                    messages = selection == null
                        ? reader.CaptureActiveSelectionMany()
                        : reader.CaptureSelectionMany(selection);
                }
                catch when (selection != null)
                {
                    messages = reader.CaptureActiveSelectionMany();
                }

                ApplySelectedMessages(messages);
            }
            catch (Exception exception)
            {
                Log.Error("CaptureRibbonSelection", exception);
                var details = DiagnosticDetails.ForException(
                    exception,
                    "EMAIL_SELECTION_FAILED");
                SetStatus(FirstLine(details), true);
            }
        }

        public void AddActiveSelection()
        {
            UseRibbonSelection(null);
        }

        private void ApplySelectedMessages(
            IReadOnlyList<MessageSnapshot> messages)
        {
            if (messages == null || messages.Count == 0)
            {
                throw new InvalidOperationException(
                    "Select one to " +
                    MailboxWorkingSet.MaxMessages +
                    " emails in Outlook first.");
            }

            if (messages.Count == 1)
            {
                SetSelectedMessage(messages[0]);
                SetStatus("Email added to context", false);
                return;
            }

            SetWorkingMessages(
                messages,
                messages.Count + " emails selected in Outlook");
        }

        private void SetSelectedMessage(MessageSnapshot message)
        {
            _workingMessages.Clear();
            RefreshContextLayer("External files");
            _selectedMessage = message ??
                throw new ArgumentNullException(nameof(message));
            var displaySubject = SubjectDisplay.Clean(
                _selectedMessage.Subject);
            SetScope(
                "Selected: " +
                (string.IsNullOrWhiteSpace(displaySubject)
                    ? "(No subject)"
                    : displaySubject));
        }

        private void SetWorkingMessages(
            IEnumerable<MessageSnapshot> messages,
            string source)
        {
            var bounded = MailboxWorkingSet.Normalize(messages);
            _workingMessages.Clear();
            foreach (var message in bounded)
            {
                _workingMessages.Add(message);
            }

            _selectedMessage = null;
            SetScope(
                "Working set: " +
                _workingMessages.Count +
                " of " +
                MailboxWorkingSet.MaxMessages +
                " emails");
            RefreshContextLayer(source);
            AppendContext(
                TextBoundary.SingleLine(source, 260) +
                " - working set ready");
            SetStatus("Working set ready", false);
        }

        private void HandleClearContext()
        {
            if (_busy)
            {
                return;
            }

            _workingMessages.Clear();
            _externalContext.Clear();
            _externalImages.Clear();
            _selectedMessage = null;
            RefreshContextLayer("External files");
            SetScopeUnavailable(
                "No context - use /search or select emails");
            SetStatus("Context cleared", false);
        }

        // ------------------------------------------------------------------
        // Suite exchange: deliberate, bounded hand-off between the
        // Scribble panes in Outlook, Excel, and PowerPoint.
        // ------------------------------------------------------------------

        private void HandleShareAnswer()
        {
            if (_lastAssistantText.Length == 0)
            {
                SetStatus(
                    "Nothing to share yet - ask something first",
                    true);
                return;
            }

            SuiteExchange.Save(
                "Outlook",
                "Answer from Outlook",
                _lastAssistantText);
            SetStatus(
                "Shared - open the Scribble pane in another app " +
                "and choose Add shared Scribble context",
                false);
        }

        private void HandleAddShared()
        {
            if (_busy)
            {
                return;
            }

            var entry = SuiteExchange.TryLoad();
            if (entry == null)
            {
                SetStatus(
                    "No shared Scribble context found",
                    true);
                return;
            }

            if (_externalContext.Count >=
                ExternalContextDocument.MaxDocuments)
            {
                SetStatus(
                    "Document limit reached (" +
                    ExternalContextDocument.MaxDocuments +
                    " files, bounded text)",
                    true);
                return;
            }

            var usedCharacters = 0;
            foreach (var existing in _externalContext)
            {
                usedCharacters +=
                    existing.Document.Content.Length;
            }

            var remaining =
                ContextScale.Scaled(
                    ExternalContextDocument.MaxTotalCharacters) -
                usedCharacters;
            if (remaining <= 0)
            {
                SetStatus(
                    "Context text budget reached",
                    true);
                return;
            }

            var name = "Shared from " + entry.Source +
                (entry.Title.Length > 0
                    ? ": " + entry.Title
                    : string.Empty);
            var document = new ExternalContextDocument(
                name,
                entry.Content);
            var warn = false;
            if (document.Content.Length > remaining)
            {
                document = new ExternalContextDocument(
                    name,
                    document.Content.Substring(0, remaining));
                warn = true;
            }

            if (document.Content.Length == 0)
            {
                return;
            }

            _externalContext.Add(new ExternalDocumentContext(
                document,
                warn,
                warn
                    ? "shared - clipped to the budget"
                    : "shared " + entry.SavedAt));
            AppendContext("Added " + name);
            RefreshContextLayer("External files");
            SetStatus("Shared context added", false);
        }

        private void HandleWebFileDrop(
            CoreWebView2WebMessageReceivedEventArgs eventArgs)
        {
            if (_busy)
            {
                return;
            }

            var paths = new List<string>();
            var objects = eventArgs.AdditionalObjects;
            if (objects == null)
            {
                return;
            }

            foreach (var item in objects)
            {
                var file = item as CoreWebView2File;
                if (file != null &&
                    !string.IsNullOrEmpty(file.Path))
                {
                    paths.Add(file.Path);
                }
            }

            if (paths.Count > 0)
            {
                AddExternalFiles(paths);
            }
        }

        private void HandleAddFiles()
        {
            if (_busy)
            {
                return;
            }

            using (var dialog = new OpenFileDialog
            {
                Title = "Add bounded text context to Scribble",
                Multiselect = true,
                CheckFileExists = true,
                Filter =
                    "Supported files|*.txt;*.md;*.csv;*.tsv;*.json;*.xml;*.yaml;*.yml;*.ini;*.html;*.htm;*.log;*.pdf;*.docx;*.docm;*.dotx;*.dotm;*.pptx;*.pptm;*.ppsx;*.ppsm;*.potx;*.xlsx;*.xlsm;*.xlsb;*.xltx;*.xltm;*.xls;*.doc;*.ppt;*.rtf;*.odt;*.ods;*.odp;*.msg;*.oft;*.eml;*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.webp;*.tif;*.tiff|" +
                    "All files|*.*"
            })
            {
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    AddExternalFiles(dialog.FileNames);
                }
            }
        }

        // Any supported file type is accepted: documents run through the
        // same bounded extractors as email attachments, and images become
        // vision input with a tray thumbnail. ExternalContextLoader
        // remains the strict text-only path for programmatic use.
        private void AddExternalFiles(IEnumerable<string> paths)
        {
            try
            {
                var added = 0;
                foreach (var path in
                    paths ?? new string[0])
                {
                    if (string.IsNullOrWhiteSpace(path))
                    {
                        continue;
                    }

                    var content =
                        EmailAttachmentReader.LoadLocalFile(path);
                    if (content == null)
                    {
                        continue;
                    }

                    if (content.ImageDataUrl.Length > 0)
                    {
                        if (_externalImages.Count >=
                            MaxExternalImages)
                        {
                            SetStatus(
                                "Image limit reached (" +
                                MaxExternalImages + ")",
                                true);
                            continue;
                        }

                        _externalImages.Add(
                            new ExternalImageContext(
                                new VisionImagePayload(
                                    content.FileName,
                                    content.ImageDataUrl),
                                EmailAttachmentReader
                                    .BuildThumbnailDataUrl(path)));
                        AppendContext(
                            "Added image " + content.FileName);
                        added++;
                        continue;
                    }

                    if (content.Text.Length == 0)
                    {
                        continue;
                    }

                    if (_externalContext.Count >=
                        ExternalContextDocument.MaxDocuments)
                    {
                        SetStatus(
                            "Document limit reached (" +
                            ExternalContextDocument.MaxDocuments +
                            " files, bounded text)",
                            true);
                        continue;
                    }

                    var usedCharacters = 0;
                    foreach (var existing in _externalContext)
                    {
                        usedCharacters +=
                            existing.Document.Content.Length;
                    }

                    var remaining =
                        ContextScale.Scaled(
                            ExternalContextDocument
                                .MaxTotalCharacters) -
                        usedCharacters;
                    if (remaining <= 0)
                    {
                        SetStatus(
                            "Context text budget reached (" +
                            ContextScale.Scaled(
                                ExternalContextDocument
                                    .MaxTotalCharacters) +
                            " characters across files)",
                            true);
                        continue;
                    }

                    var warn = false;
                    var subtitle = string.Empty;
                    if (content.Kind == "unreadable")
                    {
                        warn = true;
                        subtitle = content.Text.IndexOf(
                            "Too large",
                            StringComparison.Ordinal) >= 0
                            ? "Over " +
                              (EmailAttachmentReader
                                  .MaxBytesPerAttachment /
                                  (1024 * 1024)) +
                              " MB cap - content not read"
                            : "Unsupported type - noted for the model";
                    }
                    else if (content.Truncated)
                    {
                        warn = true;
                        subtitle = "Over the text cap - first " +
                            ContextScale.Scaled(
                                ExternalContextDocument
                                    .MaxCharactersPerDocument) +
                            " characters kept";
                    }

                    // Re-bounding to the document cap would clip the
                    // reader's trailing truncation notice, so rebuild
                    // the text with the notice inside the cap.
                    var documentText = content.Text;
                    if (content.Truncated)
                    {
                        var marker = "\n[Truncated: more content " +
                            "follows in the original file.]";
                        documentText = TextBoundary.PlainText(
                            content.Text,
                            ContextScale.Scaled(
                                ExternalContextDocument
                                    .MaxCharactersPerDocument) -
                            marker.Length) + marker;
                    }

                    var document = new ExternalContextDocument(
                        content.FileName,
                        documentText);
                    if (document.Content.Length > remaining)
                    {
                        document = new ExternalContextDocument(
                            content.FileName,
                            document.Content.Substring(0, remaining));
                        warn = true;
                        if (subtitle.Length == 0)
                        {
                            subtitle = "Clipped to the shared " +
                                "context budget";
                        }
                    }

                    if (document.Content.Length == 0)
                    {
                        continue;
                    }

                    var duplicate = false;
                    foreach (var existing in _externalContext)
                    {
                        if (string.Equals(
                                existing.Document.Name,
                                document.Name,
                                StringComparison.OrdinalIgnoreCase) &&
                            string.Equals(
                                existing.Document.Content,
                                document.Content,
                                StringComparison.Ordinal))
                        {
                            duplicate = true;
                            break;
                        }
                    }

                    if (duplicate)
                    {
                        continue;
                    }

                    _externalContext.Add(new ExternalDocumentContext(
                        document,
                        warn,
                        subtitle));
                    AppendContext(
                        "Added " + content.FileName +
                        (warn ? " (" + subtitle + ")" : string.Empty));
                    added++;
                }

                RefreshContextLayer("External files");
                if (added > 0)
                {
                    SetStatus(
                        added +
                        (added == 1
                            ? " item added"
                            : " items added"),
                        false);
                }
            }
            catch (Exception exception)
            {
                var details = DiagnosticDetails.ForException(
                    exception,
                    "EXTERNAL_CONTEXT_FAILED");
                SetStatus(FirstLine(details), true);
                Log.Error("AddExternalContext", exception);
            }
        }

        private void ChatPaneDragEnter(
            object sender,
            DragEventArgs eventArgs)
        {
            if (_busy || eventArgs.Data == null)
            {
                eventArgs.Effect = DragDropEffects.None;
                return;
            }

            if (HasOutlookDragFormat(eventArgs.Data))
            {
                eventArgs.Effect = DragDropEffects.Link;
                return;
            }

            if (eventArgs.Data.GetDataPresent(DataFormats.FileDrop))
            {
                eventArgs.Effect = DragDropEffects.Copy;
                return;
            }

            eventArgs.Effect = DragDropEffects.None;
        }

        private void ChatPaneDragDrop(
            object sender,
            DragEventArgs eventArgs)
        {
            if (_busy || eventArgs.Data == null)
            {
                return;
            }

            try
            {
                if (HasOutlookDragFormat(eventArgs.Data))
                {
                    AddActiveSelection();
                    return;
                }

                if (eventArgs.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    var paths = eventArgs.Data.GetData(
                        DataFormats.FileDrop) as string[];
                    AddExternalFiles(paths);
                }
            }
            catch (Exception exception)
            {
                var details = DiagnosticDetails.ForException(
                    exception,
                    "CONTEXT_DROP_FAILED");
                SetStatus(FirstLine(details), true);
                Log.Error("DropContext", exception);
            }
        }

        private static bool HasOutlookDragFormat(IDataObject data)
        {
            foreach (var format in data.GetFormats())
            {
                if (format.Equals(
                        "RenPrivateMessages",
                        StringComparison.OrdinalIgnoreCase) ||
                    format.Equals(
                        "FileGroupDescriptor",
                        StringComparison.OrdinalIgnoreCase) ||
                    format.Equals(
                        "FileGroupDescriptorW",
                        StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        // ------------------------------------------------------------------
        // Local /search command.
        // ------------------------------------------------------------------

        private void HandleLocalSearchCommand(
            string prompt,
            LocalSearchCommand command)
        {
            AppendUserTurn(prompt);
            switch (command.Kind)
            {
                case LocalSearchCommandKind.Help:
                    AppendContext(
                        "/search <person or topic> builds a ten-email " +
                        "working set; /search clear removes it");
                    SetStatus("Search help shown", false);
                    return;
                case LocalSearchCommandKind.Clear:
                    _workingMessages.Clear();
                    RefreshContextLayer("External files");
                    _selectedMessage = null;
                    SetScopeUnavailable(
                        "No context - use /search or select emails");
                    AppendContext("Working set cleared");
                    SetStatus("Working set cleared", false);
                    return;
                case LocalSearchCommandKind.Search:
                    SearchWorkingMessages(command.Query);
                    return;
                default:
                    return;
            }
        }

        private void SearchWorkingMessages(string query)
        {
            if (_outlookApplication == null)
            {
                SetStatus(
                    "[OUTLOOK_NOT_READY] Outlook is still initializing",
                    true);
                return;
            }

            SetStatus("Searching mailbox...", false);
            try
            {
                var hits = new MailboxContextService(
                    _outlookApplication)
                    .Search(
                        query,
                        "all",
                        3650,
                        MailboxWorkingSet.MaxMessages);
                var messages = new List<MessageSnapshot>();
                foreach (var hit in hits)
                {
                    messages.Add(hit.Message);
                }

                if (messages.Count == 0)
                {
                    AppendContext(
                        "No emails matched '" + query + "'" +
                        (_workingMessages.Count > 0
                            ? " - previous working set kept"
                            : ""));
                    SetStatus("No matches - refine /search", true);
                    return;
                }

                SetWorkingMessages(
                    messages,
                    "Search: " + query);
            }
            catch (Exception exception)
            {
                Log.Error("LocalMailboxSearch", exception);
                var details = DiagnosticDetails.ForException(
                    exception,
                    "LOCAL_SEARCH_FAILED");
                AppendError(details);
                SetStatus(FirstLine(details), true);
            }
        }

        // ------------------------------------------------------------------
        // Chat request flow.
        // ------------------------------------------------------------------

        // Stop must be 100% responsive: bump the generation so every
        // in-flight continuation becomes stale and discards itself,
        // cancel the HTTP call, and release the UI right now rather
        // than waiting for the network stack to notice.
        private void HandleStop()
        {
            _requestGeneration++;
            try
            {
                _requestCancellation?.Cancel();
            }
            catch (ObjectDisposedException)
            {
            }

            PostStreamEnd();
            SetBusy(false);
            SetStatus("Stopped", false);
        }

        private async void HandleSendMessage(string rawText)
        {
            if (_busy)
            {
                return;
            }

            var prompt = TextBoundary.PlainText(
                rawText,
                TextBoundary.MaxUserPromptCharacters);
            if (prompt.Length == 0)
            {
                SetStatus("Type a message first", true);
                return;
            }

            var localCommand = LocalSearchCommand.Parse(prompt);
            if (localCommand.Kind != LocalSearchCommandKind.None)
            {
                HandleLocalSearchCommand(prompt, localCommand);
                return;
            }

            if (_outlookApplication == null)
            {
                SetStatus(
                    "[OUTLOOK_NOT_READY] Outlook is still initializing",
                    true);
                return;
            }

            if (!_settings.IsConfigured)
            {
                OpenSettings();
                if (!_settings.IsConfigured)
                {
                    PostToWeb(new Dictionary<string, object>
                    {
                        { "type", "restorePrompt" },
                        { "text", prompt }
                    });
                    return;
                }
            }

            var requestSelectedMessage = _selectedMessage;
            var requestWorkingMessages =
                new List<MessageSnapshot>(_workingMessages);
            var requestExternalContext =
                new List<ExternalContextDocument>();
            foreach (var entry in _externalContext)
            {
                requestExternalContext.Add(entry.Document);
            }
            var requestExternalImages =
                new List<VisionImagePayload>();
            foreach (var image in _externalImages)
            {
                requestExternalImages.Add(image.Payload);
            }

            var hasLinkedDraft =
                _draftTools != null &&
                _draftTools.HasActiveDraft;
            // Document-production phrasing ("build a slide of my
            // day", "create an excel") also unlocks the one-shot
            // draft so mailbox content can be handed to Excel/
            // PowerPoint/Word - even while an email draft is
            // linked, since the cross-app tools write into sibling
            // apps, never into the linked draft.
            var draftAuthorization =
                new OneShotDraftAuthorization(
                    (!hasLinkedDraft &&
                     DraftIntentPolicy.AllowsCreate(prompt)) ||
                    DocumentDraftIntentPolicy.AllowsDraft(
                        prompt),
                    hasLinkedDraft &&
                    DraftIntentPolicy.AllowsUpdate(prompt));
            // Separate permission for the sibling-app tools: one
            // deliverable, up to four bounded calls, so a dense deck
            // is not squeezed into a single JSON payload. The email
            // permission above stays strictly single-shot.
            var crossAppAuthorization =
                new OneShotDraftAuthorization(
                    draftAuthorization.CanCreate,
                    false,
                    4);
            _diagnostics.BeginRequest(
                "Outlook",
                _settings.Model,
                draftAuthorization.CanCreate ||
                draftAuthorization.CanUpdate);
            AppendUserTurn(prompt);
            SetBusy(true);
            _requestStartedAt = DateTime.UtcNow;
            var generation = ++_requestGeneration;
            var cancellation = new CancellationTokenSource();
            _requestCancellation = cancellation;

            try
            {
                var response = await CompleteMailboxChatAsync(
                    requestSelectedMessage,
                    requestWorkingMessages,
                    requestExternalContext,
                    requestExternalImages,
                    prompt,
                    draftAuthorization,
                    crossAppAuthorization,
                    cancellation.Token);
                if (generation != _requestGeneration)
                {
                    // The user stopped this request; the reply
                    // arrived anyway and is discarded.
                    return;
                }

                _history.Add(new ChatTurn("user", prompt));
                _history.Add(
                    new ChatTurn("assistant", response));
                _lastAssistantText = response;
                _memory.LastAnswer = response;
                AppendFormattedAssistantText(response);
                if (crossAppAuthorization.IsCreated)
                {
                    // Slides, a sheet, or a document were drafted
                    // into a sibling app rather than an email.
                    SetStatus(
                        "Draft created - unsaved, open for review",
                        false);
                    _diagnostics.CompleteRequest(
                        "Done - sibling app draft created");
                }
                else if (draftAuthorization.IsCreated)
                {
                    SetStatus(
                        "Draft created - unsent, open for review",
                        false);
                    _diagnostics.CompleteRequest(
                        "Done - draft created");
                }
                else if (draftAuthorization.IsUpdated)
                {
                    SetStatus("Draft updated", false);
                    _diagnostics.CompleteRequest(
                        "Done - draft updated");
                }
                else if (draftAuthorization.IsConsumed ||
                         crossAppAuthorization.IsConsumed)
                {
                    SetStatus(
                        draftAuthorization.CanUpdate
                            ? "Draft update did not complete"
                            : "Draft creation did not complete",
                        true);
                    _diagnostics.CompleteRequest(
                        "Draft attempt consumed but not completed");
                }
                else if (draftAuthorization.CanCreate)
                {
                    SetStatus("No draft was created", false);
                    _diagnostics.CompleteRequest(
                        "Done - no draft was created");
                }
                else
                {
                    SetStatus(
                        hasLinkedDraft
                            ? "Done - draft unchanged"
                            : "Done",
                        false);
                    _diagnostics.CompleteRequest("Done");
                }
            }
            catch (OperationCanceledException)
            {
                _diagnostics.CompleteRequest("Stopped by user");
                PostToWeb(new Dictionary<string, object>
                {
                    { "type", "restorePrompt" },
                    { "text", prompt }
                });
                if (generation == _requestGeneration)
                {
                    SetStatus("Stopped - prompt restored", false);
                }
            }
            catch (Exception exception)
            {
                Log.Error("CompleteMailboxChat", exception);
                _diagnostics.CompleteRequest(
                    "Failed: " + FirstLine(exception.Message));
                if (generation == _requestGeneration)
                {
                    var details = DiagnosticDetails.ForException(
                        exception,
                        "AI_REQUEST_FAILED");
                    AppendError(details);
                    PostToWeb(new Dictionary<string, object>
                    {
                        { "type", "restorePrompt" },
                        { "text", prompt }
                    });
                    SetStatus(FirstLine(details), true);
                }
            }
            finally
            {
                PostStreamEnd();
                if (ReferenceEquals(
                    _requestCancellation,
                    cancellation))
                {
                    _requestCancellation = null;
                }

                cancellation.Dispose();
                if (generation == _requestGeneration)
                {
                    SetBusy(false);
                }

                UpdateDraftState();
            }
        }

        private async Task<string> CompleteMailboxChatAsync(
            MessageSnapshot selectedMessage,
            IReadOnlyList<MessageSnapshot> workingMessages,
            IReadOnlyList<ExternalContextDocument> externalContext,
            IReadOnlyList<VisionImagePayload> externalImages,
            string prompt,
            OneShotDraftAuthorization draftAuthorization,
            OneShotDraftAuthorization crossAppAuthorization,
            CancellationToken cancellationToken)
        {
            var activeDraft = draftAuthorization.CanUpdate
                ? _draftTools?.ActiveDraft
                : null;
            var imagesExpected = ModelRouting.ContextMayIncludeImages(
                selectedMessage,
                workingMessages) ||
                externalImages.Count > 0;
            var activeModel = ModelRouting.ResolveForRequest(
                _settings,
                imagesExpected);
            // Reading budgets follow the model this request will
            // actually use, including a temporary vision switch.
            ContextScale.Apply(
                GeminiCodeAssistGateway.IsGeminiModel(activeModel));
            if (ModelRouting.IsTemporaryVisionSwitch(
                    _settings,
                    activeModel))
            {
                SetStatus(
                    "Using " + activeModel + " for images",
                    false);
            }
            else if (imagesExpected &&
                     !ModelCatalog.IsVisionCapable(activeModel))
            {
                SetStatus(
                    activeModel + " is text-only - images will " +
                    "not be read",
                    false);
            }

            // MCP definitions come from user-configured servers;
            // connecting can spawn processes or hit the network, so
            // it happens off the UI thread and one failed server is
            // skipped inside the host.
            IReadOnlyList<ChatToolDefinition> mcpTools = null;
            var mcpHost = _mcpTools;
            if (mcpHost != null && mcpHost.HasServers)
            {
                SetStatus("Connecting MCP tools...", false);
                mcpTools = await Task.Run(
                    () => mcpHost.GetDefinitions(),
                    cancellationToken);
                SetStatus("Thinking...", false);
            }

            var request = ChatRequestFactory.Create(
                activeModel,
                selectedMessage,
                _history,
                prompt,
                draftAuthorization.CanCreate,
                activeDraft,
                draftAuthorization.CanUpdate,
                workingMessages,
                externalContext,
                _settings.UseToneProfile
                    ? _settings.ToneProfile
                    : null,
                _settings.ToneStrength,
                _settings.DraftRules,
                mcpTools);
            var exposedNames = new List<string>();
            foreach (var tool in request.tools)
            {
                exposedNames.Add(tool.function.name);
            }

            _diagnostics.SetExposedTools(exposedNames);
            _diagnostics.RecordEvent(
                "resolved model: " + activeModel);
            var mailboxTools = new MailboxToolHost(
                _outlookApplication,
                selectedMessage,
                workingMessages);
            if (VisionImagePrefetch.TryInject(
                    request,
                    activeModel,
                    mailboxTools,
                    selectedMessage,
                    workingMessages))
            {
                SetStatus("Images attached for vision", false);
            }

            if (externalImages.Count > 0)
            {
                VisionAttachmentExchange.AppendVisionContext(
                    request,
                    activeModel,
                    new[]
                    {
                        new MailboxToolResult(
                            "external_files",
                            string.Empty,
                            string.Empty,
                            externalImages)
                    });
            }

            for (var round = 0;
                 round <= TextBoundary.MaxToolRounds;
                 round++)
            {
                var response =
                    await _client.CompleteStreamingAsync(
                        _settings,
                        request,
                        PostStreamDelta,
                        cancellationToken);
                var toolCalls = response.tool_calls;
                if (toolCalls == null || toolCalls.Count == 0)
                {
                    if (string.IsNullOrWhiteSpace(response.content))
                    {
                        throw new AiEndpointException(
                            "RESPONSE_MISSING_CONTENT",
                            "The model stopped without returning text.");
                    }

                    return response.content;
                }

                // A round that continues into tool calls clears any
                // preamble text that streamed to the page.
                PostStreamEnd();

                if (round == TextBoundary.MaxToolRounds)
                {
                    throw new AiEndpointException(
                        "TOOL_ROUND_LIMIT",
                        "The model exceeded the maximum number of bounded tool rounds.");
                }

                if (toolCalls.Count >
                    TextBoundary.MaxToolCallsPerRound)
                {
                    throw new AiEndpointException(
                        "TOOL_CALL_LIMIT",
                        "The model requested too many tools in one round.");
                }

                var results = new List<MailboxToolResult>();
                foreach (var toolCall in toolCalls)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var isDraftCall =
                        DraftToolCatalog.IsDraftTool(
                            toolCall?.function?.name);
                    var isCrossAppCall =
                        CrossAppToolCatalog.IsCrossAppTool(
                            toolCall?.function?.name);
                    MailboxToolResult result;
                    if (isDraftCall)
                    {
                        result = _draftTools.Execute(
                            toolCall,
                            mailboxTools.ResolveHandle,
                            draftAuthorization,
                            toolCalls.Count == 1);
                    }
                    else if (isCrossAppCall &&
                             _crossAppTools != null)
                    {
                        // Mailbox content handed to Excel/
                        // PowerPoint/Word as a clearly marked
                        // draft, on the same one-shot permission.
                        result = _crossAppTools.Execute(
                            toolCall,
                            crossAppAuthorization,
                            toolCalls.Count == 1,
                            prompt);
                    }
                    else if (McpToolHost.IsMcpTool(
                        toolCall?.function?.name))
                    {
                        // MCP calls run off the UI thread; the
                        // host bounds the result and marks it
                        // untrusted.
                        result = await Task.Run(
                            () => mcpHost != null
                                ? mcpHost.Execute(toolCall)
                                : mailboxTools.Execute(toolCall),
                            cancellationToken);
                    }
                    else
                    {
                        result = mailboxTools.Execute(toolCall);
                    }
                    results.Add(result);
                    _diagnostics.RecordEvent(
                        "tool " +
                        (toolCall?.function?.name ?? "(null)") +
                        " -> " + result.StatusText);
                    if (isDraftCall || isCrossAppCall)
                    {
                        AppendDraftAction(result.StatusText);
                    }
                    else
                    {
                        AppendContext(result.StatusText);
                    }

                    SetStatus(result.StatusText, false);
                }

                activeModel = ModelRouting.ResolveForRequest(
                    _settings,
                    imagesExpected,
                    results);
                request.model = TextBoundary.PlainText(
                    activeModel,
                    200);
                if (ModelRouting.IsTemporaryVisionSwitch(
                        _settings,
                        activeModel))
                {
                    SetStatus(
                        "Using " + activeModel + " for images",
                        false);
                }

                ChatRequestFactory.AppendToolExchange(
                    request,
                    response,
                    results,
                    activeModel);
            }

            throw new AiEndpointException(
                "TOOL_ROUND_LIMIT",
                "The model did not finish after bounded tool use.");
        }

        // ------------------------------------------------------------------
        // Chat lifecycle and settings.
        // ------------------------------------------------------------------

        private void HandleNewChat()
        {
            if (_busy)
            {
                return;
            }

            _history.Clear();
            _workingMessages.Clear();
            _externalContext.Clear();
            _externalImages.Clear();
            _transcriptEvents.Clear();
            _lastAssistantText = string.Empty;
            _memory.LastAnswer = string.Empty;
            _draftTools?.Dispose();
            _draftTools = _outlookApplication == null
                ? null
                : new DraftToolHost(_outlookApplication);
            _crossAppTools?.Dispose();
            _crossAppTools = _outlookApplication == null
                ? null
                : new Scribble.Office.DocumentDraftHost(
                    "outlook",
                    _outlookApplication);
            PostToWeb(new Dictionary<string, object>
            {
                { "type", "clear" }
            });
            PushContextToWeb();
            SetScopeUnavailable(
                "No context - drag emails here, or use Add email " +
                "or /search");
            UpdateDraftState();
            SetStatus("Chat and context cleared", false);
        }

        private void HandleSetModel(string model)
        {
            if (_busy ||
                !ModelSelectionPolicy.IsGenerativeModel(model))
            {
                return;
            }

            if (string.Equals(
                model,
                _settings.Model,
                StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            try
            {
                _settings.Model = model;
                _settingsStore.Save(_settings);
                SetStatus("Model: " + model, false);
            }
            catch (Exception exception)
            {
                Log.Error("SwitchModel", exception);
                SetStatus("The model change was not saved", true);
            }
        }

        private void OpenSettings()
        {
            if (_busy)
            {
                return;
            }

            using (var settingsWindow =
                new SettingsWindow(
                    _settingsStore,
                    _settings,
                    _outlookApplication))
            {
                if (settingsWindow.ShowDialog(this) ==
                    DialogResult.OK)
                {
                    _settings =
                        settingsWindow.SavedSettings;
                    ContextScale.Apply(
                        GeminiCodeAssistGateway.IsGeminiModel(
                            _settings.Model));
                    _settings.ApplyLimits();
                    _mcpTools?.Dispose();
                    _mcpTools = new McpToolHost(
                        _settings.McpServers);
                    RefreshModelPicker();
                    SetStatus(
                        "Settings saved - " + _settings.Model,
                        false);
                }

                // The support report opens only after the modal
                // settings dialog is gone - Outlook refuses to
                // display mail while a dialog box is open.
                if (settingsWindow.SupportReportRequested)
                {
                    var reportError =
                        SettingsWindow.OpenSupportReport(
                            _outlookApplication,
                            settingsWindow
                                .SupportReportDescription);
                    SetStatus(
                        reportError == null
                            ? "Report email opened - review and " +
                              "send it yourself"
                            : FirstLine(reportError),
                        reportError != null);
                }
            }
        }

        // Copies the bounded per-request diagnostics record to the
        // clipboard so a misbehaving request can be reported
        // precisely. Contains no keys, settings, or message bodies.
        private void HandleCopyDiagnostics()
        {
            try
            {
                Clipboard.SetText(_diagnostics.BuildReport(
                    "Scribble Outlook pane"));
                SetStatus(
                    "Diagnostics copied to the clipboard",
                    false);
            }
            catch (Exception exception)
            {
                Log.Error("CopyDiagnostics", exception);
                SetStatus(
                    "Could not copy diagnostics",
                    true);
            }
        }

        internal void Shutdown()
        {
            if (_shutdown)
            {
                return;
            }

            _shutdown = true;
            _elapsedTimer.Stop();
            _elapsedTimer.Dispose();
            _requestCancellation?.Cancel();
            _requestCancellation?.Dispose();
            _requestCancellation = null;
            _client.Dispose();
            _draftTools?.Dispose();
            _draftTools = null;
            _crossAppTools?.Dispose();
            _crossAppTools = null;
            _mcpTools?.Dispose();
            _mcpTools = null;
            _outlookApplication = null;
            try
            {
                _webView.Dispose();
            }
            catch (Exception exception)
            {
                Log.Error("WebViewDispose", exception);
            }

            if (ReferenceEquals(LastCreated, this))
            {
                LastCreated = null;
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                Shutdown();
            }

            base.Dispose(disposing);
        }

        private static string FirstLine(string value)
        {
            var text = value ?? string.Empty;
            var index = text.IndexOfAny(
                new[] { '\r', '\n' });
            return index >= 0
                ? text.Substring(0, index)
                : text;
        }
    }
}

/*
The Scribble pane for Excel and PowerPoint. It reuses the embedded chat
page, the chat client, and every text boundary of the Outlook pane,
while its capability surface is document-shaped: bounded read-only
workbook/presentation tools, and one-shot clearly marked draft
writes (Scribble Draft sheet, [Scribble draft] slides, unsent Outlook
email drafts, or source-preserving Excel selection output) that only
unlock from the user's own prompt.
*/

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Scribble.Chat;
using Scribble.Configuration;
using Scribble.Office;
using Scribble.Outlook;
using Scribble.Security;
using Scribble.Utilities;

namespace Scribble.UI
{
    [ComVisible(true)]
    [Guid("BC9047E7-9AFE-4F75-BBBC-27241B1DE2FA")]
    [ProgId("Scribble.OfficePane")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class OfficeChatPane : UserControl
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

        private sealed class ExternalDocumentContext
        {
            public ExternalDocumentContext(
                ExternalContextDocument document,
                bool warn,
                string subtitle,
                ExcelSelectionSnapshot excelSelection = null)
            {
                Document = document;
                Warn = warn;
                Subtitle = subtitle ?? string.Empty;
                ExcelSelection = excelSelection;
            }

            public ExternalContextDocument Document { get; }

            public bool Warn { get; }

            public string Subtitle { get; }

            public ExcelSelectionSnapshot ExcelSelection { get; }
        }

        private readonly SettingsStore _settingsStore =
            new SettingsStore();
        private readonly SkillStore _skillStore = new SkillStore();
        private readonly OpenAiCompatibleClient _client =
            new OpenAiCompatibleClient();
        private readonly JavaScriptSerializer _serializer =
            new JavaScriptSerializer();
        private List<ChatTurn> _history =
            new List<ChatTurn>();
        private readonly List<ExternalDocumentContext> _externalContext =
            new List<ExternalDocumentContext>();
        private readonly List<ExternalImageContext> _externalImages =
            new List<ExternalImageContext>();
        private List<string> _transcriptEvents =
            new List<string>();
        private PaneMemory.Slot _memory;
        private readonly DiagnosticsRecorder _diagnostics =
            new DiagnosticsRecorder();
        private readonly System.Windows.Forms.Timer _elapsedTimer =
            new System.Windows.Forms.Timer
            {
                Interval = 5000
            };
        private DateTime _requestStartedAt = DateTime.UtcNow;
        private DateTime _statusChangedAt = DateTime.UtcNow;
        private readonly WebView2 _webView = new WebView2();
        private readonly PromptHelperSession _promptHelper;

        private object _hostApplication;
        private string _hostKind = string.Empty;
        private AppSettings _settings;
        private DocumentDraftHost _draftHost;
        private McpToolHost _mcpTools;
        private CancellationTokenSource _requestCancellation;
        private int _requestGeneration;
        private string _lastAssistantText = string.Empty;
        private bool _busy;
        private bool _shutdown;
        private bool _webReady;
        private bool _focusComposerWhenReady;
        private string _statusText = "Ready";
        private bool _statusError;

        public OfficeChatPane()
        {
            LastCreated = this;
            _promptHelper = new PromptHelperSession(PostToWeb);
            _settings = _settingsStore.Load();
            ContextScale.Apply(
                GeminiCodeAssistGateway.IsGeminiModel(
                    _settings.Model));
            _settings.ApplyLimits();
            _mcpTools = new McpToolHost(_settings.McpServers);
            _client.GeminiGateway.StatusListener =
                message => SetStatus(message, false);
            _elapsedTimer.Tick += ElapsedTick;
            _elapsedTimer.Start();

            Dock = DockStyle.Fill;
            BackColor = Color.FromArgb(26, 27, 30);
            MinimumSize = new Size(300, 480);
            AllowDrop = true;
            DragEnter += PaneDragEnter;
            DragDrop += PaneDragDrop;

            _webView.Dock = DockStyle.Fill;
            _webView.DefaultBackgroundColor =
                Color.FromArgb(26, 27, 30);
            Controls.Add(_webView);
            InitializeWebView();
        }

        internal static OfficeChatPane LastCreated
        {
            get;
            private set;
        }

        internal void Initialize(
            string hostKind,
            object hostApplication)
        {
            if (_hostApplication != null)
            {
                return;
            }

            _hostKind = hostKind == "excel" || hostKind == "word"
                ? hostKind
                : "powerpoint";
            _hostApplication = hostApplication ??
                throw new ArgumentNullException(
                    nameof(hostApplication));
            _draftHost = new DocumentDraftHost(
                _hostKind,
                _hostApplication);
            // Reopening the pane in the same Office session picks
            // the conversation back up from process memory.
            _memory = PaneMemory.For(_hostKind);
            _history = _memory.History;
            _transcriptEvents = _memory.Transcript;
            _lastAssistantText = _memory.LastAnswer;
            PostMode();
            if (_webReady)
            {
                PushSkillsToWeb();
                ReplayTranscript();
            }
        }

        private string HostName
        {
            get
            {
                return _hostKind == "excel"
                    ? "Excel"
                    : (_hostKind == "word"
                        ? "Word"
                        : "PowerPoint");
            }
        }

        // ------------------------------------------------------------------
        // WebView2 hosting, identical restrictions to the Outlook pane.
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
                Log.Error("OfficeWebViewInit", exception);
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
                    "restart " + HostName + ".\r\n\r\nDetails: " +
                    TextBoundary.SingleLine(
                        exception?.Message,
                        300)
            };
            Controls.Add(notice);
        }

        private static string LoadChatPage()
        {
            using (var stream = typeof(OfficeChatPane).Assembly
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
                    case "runSkill":
                        object skillOriginValue;
                        object skillIdValue;
                        message.TryGetValue(
                            "origin",
                            out skillOriginValue);
                        message.TryGetValue("id", out skillIdValue);
                        HandleRunSkill(
                            Convert.ToString(skillOriginValue) ??
                            string.Empty,
                            Convert.ToString(skillIdValue) ??
                            string.Empty);
                        break;
                    case "stop":
                        HandleStop();
                        break;
                    case "newChat":
                        HandleNewChat();
                        break;
                    case "addEmail":
                    case "emailDrop":
                        AddCurrentSelection();
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
                    case "setTopic":
                        object topicValue;
                        message.TryGetValue("topicId", out topicValue);
                        HandleSetTopic(
                            Convert.ToString(topicValue) ??
                            string.Empty);
                        break;
                    case "askUserAnswer":
                        object promptAnswerValue;
                        message.TryGetValue(
                            "answer",
                            out promptAnswerValue);
                        _promptHelper.HandleAnswer(
                            promptAnswerValue);
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
                Log.Error("OfficeWebMessage", exception);
            }
        }

        private void HandleWebReady()
        {
            _webReady = true;
            PostMode();
            RefreshModelPicker();
            PushSkillsToWeb();
            PushTopicsToWeb(false);
            PushContextToWeb();
            ReplayTranscript();
            _promptHelper.RestoreIfPending();
            if (_focusComposerWhenReady)
            {
                FocusComposer();
            }
        }

        private void ReplayTranscript()
        {
            PostToWeb(new Dictionary<string, object>
            {
                { "type", "clear" }
            });
            foreach (var recorded in _transcriptEvents.ToArray())
            {
                PostRawToWeb(recorded);
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

        private void PostMode()
        {
            if (_hostKind.Length == 0)
            {
                return;
            }

            PostToWeb(new Dictionary<string, object>
            {
                { "type", "mode" },
                { "host", _hostKind }
            });
            PostToWeb(new Dictionary<string, object>
            {
                { "type", "draft" },
                {
                    "text",
                    "Drafts stay unsaved and unsent - Scribble " +
                    "never saves or sends anything."
                },
                { "linked", false }
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
                Log.Error("OfficePostToWeb", exception);
            }
        }

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

        private void PushSkillsToWeb()
        {
            if (_hostKind.Length == 0)
            {
                return;
            }

            var publicSkills = _skillStore.LoadPublic()
                .Where(skill => string.Equals(
                    skill.Host,
                    _hostKind,
                    StringComparison.Ordinal))
                .OrderBy(skill => skill.DisplayOrder)
                .ThenBy(skill => skill.Name)
                .Select(BuildSkillButton)
                .ToArray();
            var localSkills = _skillStore.LoadLocal()
                .Where(skill => string.Equals(
                    skill.Host,
                    _hostKind,
                    StringComparison.Ordinal))
                .OrderBy(skill => skill.Name)
                .Select(BuildSkillButton)
                .ToArray();
            PostToWeb(new Dictionary<string, object>
            {
                { "type", "skills" },
                { "publicItems", publicSkills },
                { "localItems", localSkills }
            });
        }

        private static object BuildSkillButton(SkillDefinition skill)
        {
            return new Dictionary<string, object>
            {
                { "id", skill.Id },
                { "name", skill.Name },
                { "description", skill.Description },
                { "origin", skill.Origin }
            };
        }

        private void HandleRunSkill(string origin, string id)
        {
            if (_busy || _hostKind.Length == 0)
            {
                return;
            }

            var skill = _skillStore.Resolve(origin, id, _hostKind);
            if (skill == null)
            {
                SetStatus("That skill is no longer available", true);
                PushSkillsToWeb();
                return;
            }

            if (skill.StartFresh)
            {
                HandleNewChat();
            }

            HandleSendMessage(SkillStore.ExpandPrompt(skill.Prompt));
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

        // ------------------------------------------------------------------
        // Context tray: document selection snapshots, files, images,
        // and shared suite snippets all ride the same bounded
        // external-document budget.
        // ------------------------------------------------------------------

        private void PushContextToWeb()
        {
            var items = new List<object>();
            for (var index = 0;
                 index < _externalContext.Count;
                 index++)
            {
                var entry = _externalContext[index];
                var subtitle = entry.Subtitle.Length > 0
                    ? entry.Subtitle
                    : entry.Document.Content.Length +
                      " text characters";
                var card = new Dictionary<string, object>
                {
                    { "kind", "file" },
                    { "index", index },
                    { "badge", entry.Warn ? "!" : "F" },
                    {
                        "title",
                        TextBoundary.SingleLine(
                            entry.Document.Name,
                            180)
                    },
                    { "subtitle", subtitle }
                };
                if (entry.Warn)
                {
                    card["warn"] = true;
                }

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

            PushContextToWeb();
            SetStatus("Removed from context", false);
        }

        private void HandleClearContext()
        {
            if (_busy)
            {
                return;
            }

            _externalContext.Clear();
            _externalImages.Clear();
            PushContextToWeb();
            SetStatus("Context cleared", false);
        }

        // Captures the current Excel selection or PowerPoint slide
        // as a bounded untrusted context document.
        private void AddCurrentSelection()
        {
            if (_busy)
            {
                SetStatus(
                    "Scribble is working\u2014stop or wait before " +
                    "sending another selection",
                    true);
                return;
            }

            if (_hostApplication == null)
            {
                return;
            }

            try
            {
                string title;
                string content;
                if (_hostKind == "excel")
                {
                    AddExcelSelection(
                        new WorkbookToolHost(
                            _hostApplication).CaptureSelection());
                    return;
                }
                else if (_hostKind == "word")
                {
                    content = new WordToolHost(
                        _hostApplication).DescribeSelection(
                        out title);
                }
                else
                {
                    content = new PresentationToolHost(
                        _hostApplication).DescribeCurrentSlide(
                        out title);
                }

                AddContextDocument(
                    title,
                    content,
                    "from " + HostName);
                SetStatus("Selection added to context", false);
            }
            catch (Exception exception)
            {
                Log.Error("OfficeAddSelection", exception);
                var details = DiagnosticDetails.ForException(
                    exception,
                    "SELECTION_CONTEXT_FAILED");
                SetStatus(FirstLine(details), true);
            }
        }

        // Ribbon callbacks capture before opening the task pane and
        // hand the immutable snapshot in here. The click attaches
        // context and focuses the composer; it never submits a prompt.
        internal void AddExcelSelection(
            ExcelSelectionSnapshot snapshot)
        {
            if (_busy)
            {
                SetStatus(
                    "Scribble is working\u2014stop or wait before " +
                    "sending another selection",
                    true);
                return;
            }

            if (snapshot == null)
            {
                SetStatus("Select cells in Excel first", true);
                return;
            }

            var added = AddContextDocument(
                "Excel " + snapshot.WorksheetName + "!" +
                    snapshot.Address,
                snapshot.BuildContextText(string.Empty),
                "from Excel" +
                    (snapshot.PreviewTruncated
                        ? " - preview truncated"
                        : string.Empty),
                snapshot);
            if (!added)
            {
                return;
            }

            FocusComposer();
            if (snapshot.ColumnCount != 1 ||
                snapshot.RowCount >
                    ExcelSelectionOutputPolicy.MaxSelectedCells ||
                snapshot.PreviewTruncated)
            {
                SetStatus(
                    "Selection added. For adjacent output, select one " +
                    "contiguous column in chunks of at most " +
                    ExcelSelectionOutputPolicy.MaxSelectedCells +
                    " fully captured cells",
                    false);
            }
            else
            {
                SetStatus("Selection added to context", false);
            }
        }

        private void FocusComposer()
        {
            if (!_webReady)
            {
                _focusComposerWhenReady = true;
                return;
            }

            _focusComposerWhenReady = false;
            PostToWeb(new Dictionary<string, object>
            {
                { "type", "focusComposer" }
            });
        }

        private bool AddContextDocument(
            string name,
            string content,
            string subtitle,
            ExcelSelectionSnapshot excelSelection = null)
        {
            if (_externalContext.Count >=
                ExternalContextDocument.MaxDocuments)
            {
                SetStatus(
                    "Document limit reached (" +
                    ExternalContextDocument.MaxDocuments +
                    " items, bounded text)",
                    true);
                return false;
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
                return false;
            }

            var document = new ExternalContextDocument(
                name,
                content);
            var warn = document.Content.Length <
                (content ?? string.Empty).Length;
            if (document.Content.Length > remaining)
            {
                document = new ExternalContextDocument(
                    name,
                    document.Content.Substring(0, remaining));
                warn = true;
            }

            if (document.Content.Length == 0)
            {
                return false;
            }

            _externalContext.Add(new ExternalDocumentContext(
                document,
                warn,
                subtitle +
                (warn ? " - clipped to the budget" : string.Empty),
                excelSelection));
            AppendContext("Added " + document.Name);
            PushContextToWeb();
            return true;
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

        private async void AddExternalFiles(IEnumerable<string> paths)
        {
            if (_busy)
            {
                return;
            }

            var selectedPaths = (paths ?? new string[0])
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .ToArray();
            if (selectedPaths.Length == 0)
            {
                return;
            }

            var cancellation = new CancellationTokenSource();
            _requestCancellation = cancellation;
            SetBusy(true);
            try
            {
                var loaded = await Task.Run(
                    () => EmailAttachmentReader.LoadLocalFiles(
                        selectedPaths,
                        cancellation.Token,
                        (current, total, name) =>
                        {
                            try
                            {
                                if (!IsDisposed && IsHandleCreated)
                                {
                                    BeginInvoke(new Action(() => SetStatus(
                                        "Reading " + current + " of " +
                                        total + ": " + name,
                                        false)));
                                }
                            }
                            catch (InvalidOperationException)
                            {
                            }
                        }),
                    cancellation.Token);
                cancellation.Token.ThrowIfCancellationRequested();
                var added = 0;
                var resourceLimited = loaded.Count(item =>
                    item.Content != null &&
                    item.Content.Kind == "resource-limited");
                foreach (var loadedFile in loaded)
                {
                    var content = loadedFile.Content;
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
                                loadedFile.Thumbnail));
                        AppendContext(
                            "Added image " + content.FileName);
                        added++;
                        continue;
                    }

                    if (content.Text.Length == 0)
                    {
                        continue;
                    }

                    AddContextDocument(
                        content.FileName,
                        content.Text,
                        content.Kind == "resource-limited"
                            ? "attachment resource limit - content not read"
                            : content.Truncated
                            ? "truncated to the text cap"
                            : content.Text.Length +
                              " text characters");
                    added++;
                }

                PushContextToWeb();
                if (added > 0)
                {
                    SetStatus(
                        added +
                        (added == 1
                            ? " item added"
                            : " items added") +
                        (resourceLimited > 0
                            ? "; " + resourceLimited +
                              " skipped by attachment limits"
                            : string.Empty),
                        resourceLimited > 0);
                }
                else if (resourceLimited > 0)
                {
                    SetStatus(
                        resourceLimited +
                        " file" +
                        (resourceLimited == 1 ? "" : "s") +
                        " skipped by attachment limits",
                        true);
                }
            }
            catch (OperationCanceledException)
            {
                SetStatus("Attachment reading stopped", false);
            }
            catch (Exception exception)
            {
                var details = DiagnosticDetails.ForException(
                    exception,
                    "EXTERNAL_CONTEXT_FAILED");
                SetStatus(FirstLine(details), true);
                Log.Error("OfficeAddExternalContext", exception);
            }
            finally
            {
                if (ReferenceEquals(_requestCancellation, cancellation))
                {
                    _requestCancellation = null;
                    SetBusy(false);
                }

                cancellation.Dispose();
            }
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

        private void PaneDragEnter(
            object sender,
            DragEventArgs eventArgs)
        {
            if (_busy || eventArgs.Data == null)
            {
                eventArgs.Effect = DragDropEffects.None;
                return;
            }

            eventArgs.Effect =
                eventArgs.Data.GetDataPresent(DataFormats.FileDrop)
                    ? DragDropEffects.Copy
                    : DragDropEffects.None;
        }

        private void PaneDragDrop(
            object sender,
            DragEventArgs eventArgs)
        {
            if (_busy || eventArgs.Data == null)
            {
                return;
            }

            try
            {
                if (eventArgs.Data.GetDataPresent(
                    DataFormats.FileDrop))
                {
                    var paths = eventArgs.Data.GetData(
                        DataFormats.FileDrop) as string[];
                    AddExternalFiles(paths);
                }
            }
            catch (Exception exception)
            {
                Log.Error("OfficeDropContext", exception);
            }
        }

        // ------------------------------------------------------------------
        // Suite exchange: deliberate, bounded hand-off between the
        // Scribble panes.
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
                HostName,
                "Answer from " + HostName,
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

            AddContextDocument(
                "Shared from " + entry.Source +
                (entry.Title.Length > 0
                    ? ": " + entry.Title
                    : string.Empty),
                entry.Content,
                "shared " + entry.SavedAt);
            SetStatus("Shared context added", false);
        }

        // ------------------------------------------------------------------
        // Chat request flow.
        // ------------------------------------------------------------------

        private void HandleStop()
        {
            _requestGeneration++;
            _promptHelper.Cancel();
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

            if (_hostApplication == null)
            {
                SetStatus(
                    "[HOST_NOT_READY] " + HostName +
                    " is still initializing",
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

            TopicConfig activeTopic;
            string topicError;
            if (!TryResolveActiveTopic(
                    out activeTopic,
                    out topicError))
            {
                SetStatus(topicError, true);
                PostToWeb(new Dictionary<string, object>
                {
                    { "type", "restorePrompt" },
                    { "text", prompt }
                });
                return;
            }

            ExternalDocumentContext eligibleSelection = null;
            var selectionCount = 0;
            foreach (var entry in _externalContext)
            {
                if (entry.ExcelSelection != null)
                {
                    selectionCount++;
                    eligibleSelection = entry;
                }
            }

            // Any locally attached Excel snapshot is an explicit
            // document-reference gesture. Eligibility for the
            // selection-bound output tool is deliberately narrower.
            var hasAttachedExcelSelection = selectionCount > 0;
            ExcelSelectionRequestContext selectionRequest = null;
            if (selectionCount == 1 &&
                eligibleSelection != null &&
                !eligibleSelection.Warn &&
                !eligibleSelection.ExcelSelection.PreviewTruncated &&
                eligibleSelection.ExcelSelection.ColumnCount == 1 &&
                eligibleSelection.ExcelSelection.RowCount <=
                    ExcelSelectionOutputPolicy.MaxSelectedCells)
            {
                selectionRequest = new ExcelSelectionRequestContext(
                    "excel_selection_" +
                        Guid.NewGuid().ToString("N"),
                    eligibleSelection.ExcelSelection);
            }

            var requestExternalContext =
                new List<ExternalContextDocument>();
            foreach (var entry in _externalContext)
            {
                if (selectionRequest != null &&
                    ReferenceEquals(entry, eligibleSelection))
                {
                    requestExternalContext.Add(
                        new ExternalContextDocument(
                            entry.Document.Name,
                            selectionRequest.Snapshot.BuildContextText(
                                selectionRequest.Handle)));
                }
                else
                {
                    requestExternalContext.Add(entry.Document);
                }
            }

            var requestExternalImages =
                new List<VisionImagePayload>();
            foreach (var image in _externalImages)
            {
                requestExternalImages.Add(image.Payload);
            }

            // Document drafts unlock only from the user's own
            // prompt, exactly like the Outlook pane's policy gate.
            // One deliverable, built over up to four bounded
            // calls: a small local model cannot emit a whole dense
            // deck or workbook in a single JSON payload, so it adds
            // it in batches instead of thinning it out.
            var draftAuthorization =
                new OneShotDraftAuthorization(
                    DocumentDraftIntentPolicy.AllowsDraft(
                        prompt,
                        hasAttachedExcelSelection),
                    false,
                    4);
            _diagnostics.BeginRequest(
                HostName,
                _settings.Model,
                draftAuthorization.CanCreate);
            _memory.TopicLocked = true;
            PushTopicLock(false);
            var turnId = Guid.NewGuid().ToString("N");
            AppendUserTurn(prompt);
            SetBusy(true);
            _requestStartedAt = DateTime.UtcNow;
            var generation = ++_requestGeneration;
            var cancellation = new CancellationTokenSource();
            _requestCancellation = cancellation;
            _draftHost?.BeginExcelSelectionRequest(
                selectionRequest);

            try
            {
                var response = await CompleteDocumentChatAsync(
                    requestExternalContext,
                    requestExternalImages,
                    prompt,
                    draftAuthorization,
                    activeTopic,
                    _memory.ChatId,
                    turnId,
                    selectionRequest,
                    cancellation.Token);
                if (generation != _requestGeneration)
                {
                    return;
                }

                _history.Add(new ChatTurn("user", prompt));
                _history.Add(
                    new ChatTurn("assistant", response));
                _lastAssistantText = response;
                if (_memory != null)
                {
                    _memory.LastAnswer = response;
                }

                AppendFormattedAssistantText(response);
                if (draftAuthorization.IsCreated)
                {
                    SetStatus(
                        "Draft ready - unsaved and unsent, " +
                        "open for review",
                        false);
                    _diagnostics.CompleteRequest(
                        "Done - draft created");
                }
                else if (draftAuthorization.IsConsumed)
                {
                    SetStatus(
                        "The draft attempt did not complete",
                        true);
                    _diagnostics.CompleteRequest(
                        "Draft attempt consumed but not completed");
                }
                else
                {
                    SetStatus("Done", false);
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
                Log.Error("CompleteDocumentChat", exception);
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
                _draftHost?.EndExcelSelectionRequest();
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
            }
        }

        private async Task<string> CompleteDocumentChatAsync(
            IReadOnlyList<ExternalContextDocument> externalContext,
            IReadOnlyList<VisionImagePayload> externalImages,
            string prompt,
            OneShotDraftAuthorization draftAuthorization,
            TopicConfig activeTopic,
            string chatId,
            string turnId,
            ExcelSelectionRequestContext selectionRequest,
            CancellationToken cancellationToken)
        {
            var imagesExpected = externalImages.Count > 0;
            var activeModel = ModelRouting.ResolveForRequest(
                _settings,
                imagesExpected);
            // Reading budgets follow the model this request will
            // actually use, including a temporary vision switch.
            ContextScale.Apply(
                GeminiCodeAssistGateway.IsGeminiModel(activeModel));
            WorkbookToolHost workbookTools = null;
            PresentationToolHost presentationTools = null;
            WordToolHost wordTools = null;
            string activeContext;
            if (_hostKind == "excel")
            {
                workbookTools = new WorkbookToolHost(
                    _hostApplication);
                activeContext =
                    workbookTools.DescribeActiveContext();
            }
            else if (_hostKind == "word")
            {
                wordTools = new WordToolHost(
                    _hostApplication);
                activeContext =
                    wordTools.DescribeActiveContext();
            }
            else
            {
                presentationTools = new PresentationToolHost(
                    _hostApplication);
                activeContext =
                    presentationTools.DescribeActiveContext();
            }

            // MCP definitions come from user-configured servers;
            // connecting can spawn processes or hit the network, so
            // it happens off the UI thread.
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

            var request = DocumentChatRequestFactory.Create(
                activeModel,
                _hostKind,
                activeContext,
                _history,
                prompt,
                draftAuthorization.CanCreate,
                externalContext,
                mcpTools,
                activeTopic,
                selectionRequest != null);
            var topicTools = activeTopic == null
                ? null
                : new TopicToolHost(
                    activeTopic,
                    chatId,
                    turnId,
                    false);
            var exposedNames = new List<string>();
            foreach (var tool in request.tools)
            {
                exposedNames.Add(tool.function.name);
            }

            _diagnostics.SetExposedTools(exposedNames);
            _diagnostics.RecordEvent(
                "resolved model: " + activeModel);
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

                if (PromptHelperTool.Contains(toolCalls) &&
                    toolCalls.Count != 1)
                {
                    var rejected = new List<MailboxToolResult>();
                    foreach (var rejectedCall in toolCalls)
                    {
                        rejected.Add(
                            PromptHelperTool.MixedCallResult(
                                rejectedCall));
                    }

                    ChatRequestFactory.AppendToolExchange(
                        request,
                        response,
                        rejected,
                        activeModel);
                    request.tool_choice =
                        PromptHelperTool.CreateRequiredChoice();
                    SetStatus(
                        "Clarification must come before other work",
                        false);
                    continue;
                }

                var results = new List<MailboxToolResult>();
                foreach (var toolCall in toolCalls)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var name = toolCall?.function?.name;
                    var isDraftCall = DocumentDraftHost.IsDraftTool(
                        _hostKind,
                        name);
                    MailboxToolResult result;
                    if (PromptHelperTool.IsTool(name))
                    {
                        result = await _promptHelper.AskAsync(
                            toolCall,
                            cancellationToken);
                        request.tool_choice = "auto";
                    }
                    else if (isDraftCall)
                    {
                        result = _draftHost.Execute(
                            toolCall,
                            draftAuthorization,
                            toolCalls.Count == 1,
                            prompt);
                    }
                    else if (McpToolHost.IsMcpTool(name) &&
                             mcpHost != null)
                    {
                        // MCP calls run off the UI thread; the
                        // host bounds the result and marks it
                        // untrusted.
                        result = await Task.Run(
                            () => mcpHost.Execute(toolCall),
                            cancellationToken);
                    }
                    else if (TopicToolCatalog.IsTopicTool(name) &&
                             topicTools != null)
                    {
                        result = await Task.Run(
                            () => topicTools.Execute(
                                toolCall,
                                cancellationToken),
                            cancellationToken);
                    }
                    else if (WebReadTool.IsWebReadTool(name))
                    {
                        // Network reads run off the UI thread; the
                        // tool bounds the result and marks it
                        // untrusted.
                        result = await Task.Run(
                            () => WebReadTool.Execute(toolCall),
                            cancellationToken);
                    }
                    else if (workbookTools != null)
                    {
                        result = workbookTools.Execute(toolCall);
                    }
                    else if (wordTools != null)
                    {
                        result = wordTools.Execute(toolCall);
                    }
                    else
                    {
                        result = presentationTools.Execute(
                            toolCall);
                    }

                    results.Add(result);
                    _diagnostics.RecordEvent(
                        "tool " + (name ?? "(null)") + " -> " +
                        result.StatusText);
                    if (isDraftCall)
                    {
                        AppendDraftAction(result.StatusText);
                    }
                    else
                    {
                        AppendContext(result.StatusText);
                    }

                    SetStatus(result.StatusText, false);
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
            _externalContext.Clear();
            _externalImages.Clear();
            _transcriptEvents.Clear();
            _lastAssistantText = string.Empty;
            if (_memory != null)
            {
                _memory.LastAnswer = string.Empty;
                _memory.ChatId = Guid.NewGuid().ToString("N");
                _memory.ActiveTopicId = string.Empty;
                _memory.ActiveTopicRoot = string.Empty;
                _memory.TopicLocked = false;
            }
            _draftHost?.Dispose();
            _draftHost = _hostApplication == null
                ? null
                : new DocumentDraftHost(
                    _hostKind,
                    _hostApplication);
            PostToWeb(new Dictionary<string, object>
            {
                { "type", "clear" }
            });
            PushContextToWeb();
            PushTopicsToWeb(false);
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
                Log.Error("OfficeSwitchModel", exception);
                SetStatus("The model change was not saved", true);
            }
        }

        private void HandleSetTopic(string topicId)
        {
            if (_busy || _memory == null || _memory.TopicLocked)
            {
                return;
            }

            var bounded = TextBoundary.SingleLine(topicId, 40);
            var topic = _settings.Topics.Find(entry =>
                string.Equals(
                    entry.Id,
                    bounded,
                    StringComparison.OrdinalIgnoreCase));
            if (bounded.Length > 0 && topic == null)
            {
                SetStatus("The selected Topic is unavailable", true);
                PushTopicsToWeb(false);
                return;
            }

            _memory.ActiveTopicId = topic?.Id ?? string.Empty;
            _memory.ActiveTopicRoot = topic?.FolderPath ?? string.Empty;
            PushTopicsToWeb(false);
            SetStatus(
                topic == null
                    ? "Topic: None"
                    : "Topic: " + topic.Name,
                false);
        }

        private bool TryResolveActiveTopic(
            out TopicConfig topic,
            out string error)
        {
            topic = null;
            error = string.Empty;
            if (_memory == null ||
                _memory.ActiveTopicId.Length == 0)
            {
                return true;
            }

            var latest = _settingsStore.Load();
            _settings.Topics = latest.Topics;
            topic = latest.Topics.Find(entry => string.Equals(
                entry.Id,
                _memory.ActiveTopicId,
                StringComparison.OrdinalIgnoreCase));
            if (topic == null || !string.Equals(
                    topic.FolderPath,
                    _memory.ActiveTopicRoot,
                    StringComparison.OrdinalIgnoreCase))
            {
                topic = null;
                error =
                    "The active Topic was removed or its folder changed. Clear chat before continuing.";
                PushTopicsToWeb(true);
                return false;
            }

            string resolvedRoot;
            string validationError;
            if (!TopicConfig.TryValidateLocalFolder(
                    topic.FolderPath,
                    out resolvedRoot,
                    out validationError))
            {
                error = "The active Topic is unavailable: " +
                    validationError;
                PushTopicsToWeb(false);
                return false;
            }

            PushTopicsToWeb(false);

            return true;
        }

        private void PushTopicsToWeb(bool unavailable)
        {
            var items = new List<object>();
            foreach (var topic in _settings.Topics)
            {
                items.Add(new Dictionary<string, object>
                {
                    { "id", topic.Id },
                    { "name", topic.Name },
                    { "available", Directory.Exists(topic.FolderPath) }
                });
            }

            PostToWeb(new Dictionary<string, object>
            {
                { "type", "topics" },
                { "items", items },
                {
                    "current",
                    _memory?.ActiveTopicId ?? string.Empty
                },
                { "unavailable", unavailable }
            });
            PushTopicLock(unavailable);
        }

        private void PushTopicLock(bool unavailable)
        {
            PostToWeb(new Dictionary<string, object>
            {
                { "type", "topicLock" },
                { "locked", _memory?.TopicLocked ?? false },
                { "unavailable", unavailable }
            });
        }

        // Copies the bounded per-request diagnostics record to the
        // clipboard so a misbehaving request can be reported
        // precisely. Contains no keys, settings, or message bodies.
        private void HandleCopyDiagnostics()
        {
            try
            {
                Clipboard.SetText(_diagnostics.BuildReport(
                    "Scribble " + HostName + " pane"));
                SetStatus(
                    "Diagnostics copied to the clipboard",
                    false);
            }
            catch (Exception exception)
            {
                Log.Error("OfficeCopyDiagnostics", exception);
                SetStatus(
                    "Could not copy diagnostics",
                    true);
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
                    null,
                    _hostKind))
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
                    PushSkillsToWeb();
                    PushTopicsToWeb(
                        _memory != null &&
                        _memory.TopicLocked &&
                        _memory.ActiveTopicId.Length > 0 &&
                        !_settings.Topics.Exists(topic =>
                            string.Equals(
                                topic.Id,
                                _memory.ActiveTopicId,
                                StringComparison.OrdinalIgnoreCase) &&
                            string.Equals(
                                topic.FolderPath,
                                _memory.ActiveTopicRoot,
                                StringComparison.OrdinalIgnoreCase)));
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
                            null,
                            settingsWindow
                                .SupportReportDescription);
                    SetStatus(
                        reportError == null
                            ? "Report email opened in Outlook - " +
                              "review and send it yourself"
                            : FirstLine(reportError),
                        reportError != null);
                }
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
            _draftHost?.Dispose();
            _draftHost = null;
            _mcpTools?.Dispose();
            _mcpTools = null;
            _hostApplication = null;
            try
            {
                _webView.Dispose();
            }
            catch (Exception exception)
            {
                Log.Error("OfficeWebViewDispose", exception);
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

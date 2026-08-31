using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using Scribble.Chat;
using Scribble.Configuration;
using Scribble.Outlook;
using Scribble.Security;
using Scribble.Utilities;

namespace Scribble.UI
{
    public sealed class SettingsWindow : Form
    {
        private const int ModelDiscoveryTimeoutSeconds = 15;
        private const int EndpointProbeTimeoutSeconds = 90;

        private readonly TextBox _endpoint = new TextBox();
        private readonly ComboBox _model = new ComboBox();
        private readonly TextBox _apiKey = new TextBox();
        private readonly Label _transportWarning = new Label();
        private readonly CheckBox _switchVisionForImages = new CheckBox();
        private readonly Label _modelGuidance = new Label();
        private readonly Label _testStatus = new Label();
        private readonly CheckBox _useToneProfile = new CheckBox();
        private readonly RichTextBox _toneProfile = new RichTextBox();
        private readonly Label _toneStatus = new Label();
        private readonly Label _error = new Label();
        private readonly Button _checkEndpoint =
            MakeButton("Test selected model", false, 148);
        private readonly Button _refreshModels =
            MakeButton("Connect & load models", true, 168);
        private readonly Button _analyzeTone =
            MakeButton("Analyze 15 sent emails", false, 176);
        private readonly Button _updateButton =
            MakeButton("Update Scribble", false, 128);
        private readonly Label _updateStatus = new Label();
        private readonly CheckBox _useGeminiSignIn = new CheckBox();
        private readonly Button _googleSignIn =
            MakeButton("Sign in with Google", false, 160);
        private readonly Label _googleStatus = new Label();
        private readonly TextBox _geminiProject = new TextBox();
        private readonly TrackBar _toneStrength = new TrackBar();
        private readonly Label _toneStrengthValue = new Label();
        private readonly NumericUpDown _workingSetSize =
            new NumericUpDown();
        private readonly RichTextBox _draftRules =
            new RichTextBox();
        private readonly RichTextBox _supportText =
            new RichTextBox();
        private readonly Button _supportButton =
            MakeButton("Create report email", false, 160);
        private readonly Label _supportStatus = new Label();
        private readonly ListBox _mcpList = new ListBox();
        private readonly TextBox _mcpName = new TextBox();
        private readonly TextBox _mcpTarget = new TextBox();
        private readonly TextBox _mcpArguments = new TextBox();
        private readonly TextBox _mcpHeaders = new TextBox();
        private readonly TextBox _mcpBrowserTools = new TextBox();
        private readonly CheckBox _mcpEnabled = new CheckBox();
        private readonly CheckBox _mcpBrowserToolsApproved =
            new CheckBox();
        private readonly Button _mcpSave =
            MakeButton("Add / update server", false, 150);
        private readonly Button _mcpRemove =
            MakeButton("Remove selected", false, 140);
        private readonly Label _mcpStatus = new Label();
        private readonly List<McpServerConfig> _mcpServers =
            new List<McpServerConfig>();
        private readonly Button _save =
            MakeButton("Save", true, 96);
        private readonly OpenAiCompatibleClient _client =
            new OpenAiCompatibleClient();
        private readonly SettingsStore _store;
        private readonly object _outlookApplication;
        private CancellationTokenSource _checkCancellation;
        private CancellationTokenSource _modelRefreshCancellation;
        private CancellationTokenSource _toneCancellation;
        private CancellationTokenSource _updateCancellation;
        private bool _checking;
        private bool _analyzingTone;
        private bool _refreshingModels;
        private bool _updating;
        private bool _signingIn;
        private bool _commonControlsEnabled = true;
        private string _geminiRefreshToken = string.Empty;

        public SettingsWindow(
            SettingsStore store,
            AppSettings current)
            : this(store, current, null)
        {
        }

        public SettingsWindow(
            SettingsStore store,
            AppSettings current,
            object outlookApplication)
        {
            _store = store ??
                throw new ArgumentNullException(nameof(store));
            _outlookApplication = outlookApplication;

            Text = "Scribble settings";
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(700, 620);
            MinimumSize = new Size(620, 560);
            MaximizeBox = false;
            MinimizeBox = false;
            ShowIcon = false;
            Font = SystemFonts.MessageBoxFont;
            BackColor = SystemColors.Control;
            ForeColor = SystemColors.ControlText;
            AutoScaleMode = AutoScaleMode.Font;

            ConfigureFields();
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(18, 16, 18, 14)
            };
            root.RowStyles.Add(
                new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(
                new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 46));

            var tabs = new TabControl
            {
                Dock = DockStyle.Fill,
                AccessibleName = "Scribble settings sections"
            };
            tabs.TabPages.Add(BuildConnectionPage());
            if (!AdminPolicy.GeminiDisabled)
            {
                tabs.TabPages.Add(BuildGeminiPage());
            }
            tabs.TabPages.Add(BuildMcpPage());
            tabs.TabPages.Add(BuildWritingStylePage());
            tabs.TabPages.Add(BuildLimitsPage());
            tabs.TabPages.Add(BuildSupportPage());
            root.Controls.Add(tabs, 0, 0);

            ConfigureSupportingLabel(_error);
            _error.ForeColor = ErrorText;
            _error.AccessibleName = "Settings error";
            _error.AccessibleRole = AccessibleRole.Alert;
            root.Controls.Add(_error, 0, 1);
            var buttons = BuildButtons();
            root.Controls.Add(buttons, 0, 2);
            Controls.Add(root);
            FormClosing += SettingsWindowFormClosing;

            AcceptButton = _save;
            CancelButton = GetCancelButton(buttons);

            _endpoint.Text = current?.BaseUrl ?? string.Empty;
            _model.Text = ModelSelectionPolicy.IsGenerativeModel(
                current?.Model)
                    ? current.Model
                    : string.Empty;
            _apiKey.Text = current?.ApiKey ?? string.Empty;
            _toneProfile.Text = TextBoundary.PlainText(
                current?.ToneProfile,
                TextBoundary.MaxToneProfileCharacters);
            _useToneProfile.Checked =
                (current?.UseToneProfile ?? false) &&
                _toneProfile.TextLength > 0;
            _switchVisionForImages.Checked =
                current?.SwitchToVisionModelForImages ?? false;
            RestoreDiscoveredModels(current?.DiscoveredModels);
            _geminiRefreshToken =
                current?.GeminiRefreshToken ?? string.Empty;
            _geminiProject.Text =
                current?.GeminiProject ?? string.Empty;
            _useGeminiSignIn.Checked =
                current?.UseGeminiSignIn ?? false;
            _toneStrength.Value = Math.Max(
                10,
                Math.Min(100, current?.ToneStrength ?? 60));
            _draftRules.Text = TextBoundary.PlainText(
                current?.DraftRules,
                2000);
            var storedWorkingSet =
                current?.LimitWorkingSetMessages ?? 0;
            _workingSetSize.Value = Math.Max(
                LimitOverrides.MinWorkingSetMessages,
                Math.Min(
                    LimitOverrides.MaxWorkingSetMessages,
                    storedWorkingSet > 0
                        ? storedWorkingSet
                        : LimitOverrides
                            .RecommendedWorkingSetMessages));
            foreach (var server in current?.McpServers ??
                new List<McpServerConfig>())
            {
                if (server == null)
                {
                    continue;
                }

                var sanitized = server.Sanitized();
                if (sanitized.Target.Length > 0)
                {
                    _mcpServers.Add(sanitized);
                }
            }

            RefreshMcpList();
            UpdateToneStrengthLabel();
            UpdateModelGuidance();
            UpdateTransportWarning();
            UpdateGeminiModeUi();
        }

        public AppSettings SavedSettings { get; private set; }

        protected override void OnFormClosed(
            FormClosedEventArgs eventArgs)
        {
            _checkCancellation?.Cancel();
            _modelRefreshCancellation?.Cancel();
            _toneCancellation?.Cancel();
            _updateCancellation?.Cancel();
            _checkCancellation?.Dispose();
            _modelRefreshCancellation?.Dispose();
            _toneCancellation?.Dispose();
            _updateCancellation?.Dispose();
            _client.Dispose();
            base.OnFormClosed(eventArgs);
        }

        private static Color SecondaryText
        {
            get
            {
                return SystemInformation.HighContrast
                    ? SystemColors.GrayText
                    : Color.FromArgb(80, 80, 80);
            }
        }

        private static Color ErrorText
        {
            get
            {
                return SystemInformation.HighContrast
                    ? SystemColors.HotTrack
                    : Color.FromArgb(163, 38, 38);
            }
        }

        private static Color SuccessText
        {
            get
            {
                return SystemInformation.HighContrast
                    ? SystemColors.Highlight
                    : Color.FromArgb(20, 112, 70);
            }
        }

        private void ConfigureFields()
        {
            ConfigureField(
                _endpoint,
                "AI endpoint",
                "HTTP or HTTPS URL for an OpenAI-compatible endpoint.");
            _endpoint.TextChanged +=
                (sender, args) => UpdateTransportWarning();
            ConfigureModelField();
            ConfigureField(
                _apiKey,
                "API key",
                "Encrypted for the current Windows user.");
            _apiKey.UseSystemPasswordChar = true;

            _useGeminiSignIn.AutoSize = true;
            _useGeminiSignIn.Text =
                "Use Google Gemini with browser sign-in " +
                "(no endpoint or API key needed)";
            _useGeminiSignIn.AccessibleName =
                "Use Gemini with Google sign-in";
            _useGeminiSignIn.AccessibleDescription =
                "Signs in with your Google account in the browser " +
                "and uses Gemini models. Email context is sent to " +
                "Google instead of a local endpoint.";
            _useGeminiSignIn.CheckedChanged += GeminiModeChanged;

            ConfigureField(
                _geminiProject,
                "Google Cloud project id",
                "Leave empty unless sign-in reports that your " +
                "organization requires a designated project. The " +
                "same id works for everyone in the organization.");

            _switchVisionForImages.AutoSize = true;
            _switchVisionForImages.Text =
                "Auto-switch to vision for images";
            _switchVisionForImages.AccessibleName =
                "Switch to vision model for images";
            _switchVisionForImages.AccessibleDescription =
                "Uses your saved model list to pick a vision model for this request only. " +
                "Save settings after Connect and load models so Scribble knows which vision models are available.";

            _useToneProfile.AutoSize = true;
            _useToneProfile.Text =
                "Use this writing profile for drafts";
            _useToneProfile.AccessibleDescription =
                "Applies the editable writing profile only when creating or updating drafts.";

            _toneProfile.Dock = DockStyle.Fill;
            _toneProfile.BorderStyle = BorderStyle.FixedSingle;
            _toneProfile.Font = new Font(
                SystemFonts.MessageBoxFont.FontFamily,
                SystemFonts.MessageBoxFont.Size + 1F,
                FontStyle.Regular);
            _toneProfile.MaxLength =
                TextBoundary.MaxToneProfileCharacters;
            _toneProfile.ScrollBars =
                RichTextBoxScrollBars.Vertical;
            _toneProfile.DetectUrls = false;
            _toneProfile.AccessibleName =
                "Editable email writing profile";
        }

        private TabPage BuildConnectionPage()
        {
            var page = new TabPage("Connection")
            {
                AutoScroll = true
            };
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 1,
                RowCount = 15,
                Padding = new Padding(18, 16, 18, 12),
                Width = 640
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var intro = SupportingText(
                "Enter the endpoint URL and API key, then load the " +
                "models it serves. The model field stays editable for " +
                "servers that do not expose a model list.");
            intro.ForeColor = SystemColors.ControlText;
            layout.Controls.Add(intro, 0, 0);
            layout.Controls.Add(FieldLabel("Endpoint URL"), 0, 1);
            layout.Controls.Add(_endpoint, 0, 2);
            layout.Controls.Add(FieldLabel("API key"), 0, 3);
            layout.Controls.Add(_apiKey, 0, 4);

            ConfigureSupportingLabel(_transportWarning);
            _transportWarning.AccessibleRole = AccessibleRole.Alert;
            layout.Controls.Add(_transportWarning, 0, 5);

            _checkEndpoint.Click += CheckEndpointClick;
            _refreshModels.Click += RefreshModelsClick;
            var checkRow = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(0, 4, 0, 0),
                Margin = new Padding(0)
            };
            checkRow.Controls.Add(_refreshModels);
            layout.Controls.Add(checkRow, 0, 6);

            ConfigureSupportingLabel(_testStatus);
            _testStatus.Text =
                "Connect loads the model list. Test selected model also " +
                "verifies mailbox tool-call compatibility.";
            _testStatus.AccessibleRole = AccessibleRole.StatusBar;
            layout.Controls.Add(_testStatus, 0, 7);

            layout.Controls.Add(FieldLabel("Model"), 0, 8);
            layout.Controls.Add(_model, 0, 9);
            ConfigureSupportingLabel(_modelGuidance);
            layout.Controls.Add(_modelGuidance, 0, 10);

            var testRow = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(0, 4, 0, 0),
                Margin = new Padding(0)
            };
            testRow.Controls.Add(_checkEndpoint);
            layout.Controls.Add(testRow, 0, 11);
            layout.Controls.Add(_switchVisionForImages, 0, 12);

            _updateButton.Click += UpdateClick;
            var updateRow = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(0, 14, 0, 0),
                Margin = new Padding(0)
            };
            updateRow.Controls.Add(_updateButton);
            var versionLabel = SupportingText(
                "Installed version " +
                SelfUpdater.InstalledVersion() + ".");
            versionLabel.Padding = new Padding(8, 8, 0, 0);
            updateRow.Controls.Add(versionLabel);
            layout.Controls.Add(updateRow, 0, 13);

            ConfigureSupportingLabel(_updateStatus);
            _updateStatus.Text =
                "Update downloads the latest Scribble release and installs " +
                "silently once Outlook, Excel, PowerPoint, and Word are closed. " +
                "One update refreshes all four add-ins.";
            _updateStatus.AccessibleRole = AccessibleRole.StatusBar;
            layout.Controls.Add(_updateStatus, 0, 14);
            page.Controls.Add(layout);
            return page;
        }

        private async void UpdateClick(
            object sender,
            EventArgs eventArgs)
        {
            if (_updating ||
                _checking ||
                _refreshingModels ||
                _analyzingTone)
            {
                return;
            }

            _error.Text = string.Empty;
            var confirm = MessageBox.Show(
                this,
                "Scribble will download the latest version and install it " +
                "silently. Outlook, Excel, PowerPoint, and Word are " +
                "closed automatically so the update can finish" +
                (_outlookApplication != null
                    ? ", and Outlook reopens with the new version"
                    : string.Empty) +
                ". One update refreshes all four Scribble add-ins. " +
                "Continue?",
                "Update Scribble",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);
            if (confirm != DialogResult.Yes)
            {
                return;
            }

            // Second, deliberate warning: the update closes the
            // Office apps itself, so unsaved work must be saved
            // first. Waiting for the user to close them by hand is
            // what left installs silently unfinished.
            var closeConfirm = MessageBox.Show(
                this,
                "This will close all Office apps (Outlook, Excel, " +
                "PowerPoint, and Word).\r\n\r\n" +
                "Please save any unsaved work before continuing.",
                "Scribble will close your Office apps",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Warning);
            if (closeConfirm != DialogResult.OK)
            {
                return;
            }

            _updating = true;
            SetCommonControlsEnabled(false);
            _updateButton.Enabled = false;
            _updateCancellation = new CancellationTokenSource();
            try
            {
                _updateStatus.Text =
                    "Downloading the latest installer (up to five minutes)...";
                var installerPath =
                    await SelfUpdater.DownloadInstallerAsync(
                        _updateCancellation.Token);
                _updateStatus.ForeColor = SuccessText;
                _updateStatus.Text = _outlookApplication != null
                    ? "Update downloaded. The Office apps close now " +
                      "and Outlook reopens with the new version."
                    : "Update downloaded. The Office apps close now " +
                      "and the update installs automatically.";
                SelfUpdater.LaunchUpdateAndQuitHost(
                    _outlookApplication,
                    installerPath,
                    _outlookApplication != null
                        ? "outlook.exe"
                        : string.Empty);
                _updating = false;
                Close();
            }
            catch (OperationCanceledException)
            {
                _updating = false;
                _updateStatus.ForeColor = SecondaryText;
                _updateStatus.Text =
                    "The update was cancelled. Scribble is unchanged.";
                SetCommonControlsEnabled(true);
            }
            catch (Exception exception)
            {
                _updating = false;
                _updateStatus.ForeColor = SecondaryText;
                _updateStatus.Text =
                    "The update did not start. Scribble is unchanged.";
                _error.Text = DiagnosticDetails.ForException(
                    exception,
                    "UPDATE_FAILED");
                SetCommonControlsEnabled(true);
            }
            finally
            {
                _updateCancellation?.Dispose();
                _updateCancellation = null;
            }
        }

        private TabPage BuildGeminiPage()
        {
            var page = new TabPage("Gemini")
            {
                AutoScroll = true
            };
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 1,
                RowCount = 8,
                Padding = new Padding(18, 16, 18, 12),
                Width = 640
            };
            for (var index = 0; index < 8; index++)
            {
                layout.RowStyles.Add(
                    new RowStyle(SizeType.AutoSize));
            }

            var logoRow = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0, 0, 0, 6)
            };
            var logo = new Panel
            {
                Width = 40,
                Height = 40,
                Margin = new Padding(0, 0, 8, 0)
            };
            logo.Paint += PaintGeminiLogo;
            logoRow.Controls.Add(logo);
            var logoText = new Label
            {
                Text = "Google Gemini",
                AutoSize = true,
                Font = new Font(
                    SystemFonts.MessageBoxFont.FontFamily,
                    SystemFonts.MessageBoxFont.Size + 4F,
                    FontStyle.Bold),
                Padding = new Padding(0, 8, 0, 0)
            };
            logoRow.Controls.Add(logoText);
            layout.Controls.Add(logoRow, 0, 0);

            var responsibility = SupportingText(
                "Before you enable this: anything you send from " +
                "this pane - email text, attachments, images, and " +
                "prompts - leaves this machine and is processed by " +
                "Google Gemini under your own Google account and " +
                "your organization's Google agreement. Treat it " +
                "like any other cloud service. Do not submit " +
                "confidential, personal, or otherwise regulated " +
                "information unless your organization's data " +
                "policies allow it for this service. You are the " +
                "operator of this tool: using it in line with " +
                "those policies is your responsibility.");
            responsibility.ForeColor = SystemColors.ControlText;
            responsibility.Padding = new Padding(0, 0, 0, 8);
            layout.Controls.Add(responsibility, 0, 1);

            layout.Controls.Add(_useGeminiSignIn, 0, 2);

            _googleSignIn.Click += GoogleSignInClick;
            var geminiRow = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(0, 4, 0, 4),
                Margin = new Padding(0)
            };
            geminiRow.Controls.Add(_googleSignIn);
            ConfigureSupportingLabel(_googleStatus);
            _googleStatus.Padding = new Padding(8, 6, 0, 0);
            _googleStatus.AccessibleRole =
                AccessibleRole.StatusBar;
            geminiRow.Controls.Add(_googleStatus);
            layout.Controls.Add(geminiRow, 0, 3);

            layout.Controls.Add(
                FieldLabel(
                    "Google Cloud project (only if your " +
                    "organization's Gemini license requires one)"),
                0,
                4);
            _geminiProject.Width = 360;
            layout.Controls.Add(_geminiProject, 0, 5);

            layout.Controls.Add(
                SupportingText(
                    "Pick a gemini model on the Connection tab " +
                    "after signing in (gemini-2.5-flash is a good " +
                    "default). If a model is at capacity, Scribble " +
                    "hops to the next available one and says so in " +
                    "the status line."),
                0,
                6);
            layout.Controls.Add(
                SupportingText(
                    "Sign-in uses your browser and Google's own " +
                    "pages; Scribble never sees your password and " +
                    "stores only an encrypted sign-in token for " +
                    "this Windows user."),
                0,
                7);
            page.Controls.Add(layout);
            return page;
        }

        // A simple four-point spark next to the Gemini name; drawn
        // locally so no external asset ships with the add-in.
        private static void PaintGeminiLogo(
            object sender,
            PaintEventArgs eventArgs)
        {
            var panel = (Panel)sender;
            var graphics = eventArgs.Graphics;
            graphics.SmoothingMode =
                System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            var w = panel.Width;
            var h = panel.Height;
            var cx = w / 2f;
            var cy = h / 2f;
            var points = new[]
            {
                new PointF(cx, 1),
                new PointF(cx + w * 0.14f, cy - h * 0.14f),
                new PointF(w - 1, cy),
                new PointF(cx + w * 0.14f, cy + h * 0.14f),
                new PointF(cx, h - 1),
                new PointF(cx - w * 0.14f, cy + h * 0.14f),
                new PointF(1, cy),
                new PointF(cx - w * 0.14f, cy - h * 0.14f)
            };
            using (var brush =
                new System.Drawing.Drawing2D.LinearGradientBrush(
                    new Point(0, 0),
                    new Point(w, h),
                    Color.FromArgb(66, 133, 244),
                    Color.FromArgb(155, 114, 203)))
            {
                graphics.FillPolygon(brush, points);
            }
        }

        private TabPage BuildMcpPage()
        {
            var page = new TabPage("MCP")
            {
                AutoScroll = true
            };
            var layout = new TableLayoutPanel
            {
                // Scrolls when the rows do not fit instead of
                // squeezing them.
                AutoScroll = true,
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 16,
                Padding = new Padding(18, 16, 18, 12)
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(
                new RowStyle(SizeType.Percent, 100));
            for (var row = 2; row < 16; row++)
            {
                layout.RowStyles.Add(
                    new RowStyle(SizeType.AutoSize));
            }

            var intro = SupportingText(
                "MCP (Model Context Protocol) servers add extra " +
                "tools to the chat - local commands or HTTP " +
                "endpoints that you connect here yourself. They " +
                "run with your Windows account's own permissions, " +
                "outside this add-in's guardrails: Scribble itself " +
                "still cannot send email or save or delete " +
                "documents, but a server you add acts with " +
                "whatever powers it has. Only add servers you " +
                "trust, prefer read-only ones, and remember their " +
                "results are treated as untrusted data. Browser " +
                "chat gets no MCP tools by default; it can use " +
                "only exact tool names that you separately approve " +
                "as read-only below.");
            intro.ForeColor = SystemColors.ControlText;
            layout.Controls.Add(intro, 0, 0);

            _mcpList.Dock = DockStyle.Fill;
            _mcpList.MinimumSize = new Size(0, 110);
            _mcpList.AccessibleName = "Configured MCP servers";
            _mcpList.SelectedIndexChanged += McpSelectionChanged;
            layout.Controls.Add(_mcpList, 0, 1);

            layout.Controls.Add(
                FieldLabel("Server name (short, letters and digits)"),
                0,
                2);
            _mcpName.Dock = DockStyle.Fill;
            _mcpName.MaxLength = 24;
            _mcpName.AccessibleName = "MCP server name";
            layout.Controls.Add(_mcpName, 0, 3);

            layout.Controls.Add(
                FieldLabel(
                    "Command path or HTTP(S) endpoint URL"),
                0,
                4);
            _mcpTarget.Dock = DockStyle.Fill;
            _mcpTarget.MaxLength = 400;
            _mcpTarget.AccessibleName =
                "MCP server command or URL";
            layout.Controls.Add(_mcpTarget, 0, 5);

            layout.Controls.Add(
                FieldLabel(
                    "Command-line arguments (local commands only)"),
                0,
                6);
            _mcpArguments.Dock = DockStyle.Fill;
            _mcpArguments.MaxLength = 1000;
            _mcpArguments.AccessibleName =
                "MCP server arguments";
            layout.Controls.Add(_mcpArguments, 0, 7);

            layout.Controls.Add(
                FieldLabel(
                    "HTTP headers, one per line as Name: value " +
                    "(HTTP servers only, e.g. Authorization)"),
                0,
                8);
            _mcpHeaders.Dock = DockStyle.Fill;
            _mcpHeaders.Multiline = true;
            _mcpHeaders.ScrollBars = ScrollBars.Vertical;
            _mcpHeaders.MinimumSize = new Size(0, 48);
            _mcpHeaders.MaxLength = 2000;
            _mcpHeaders.AccessibleName =
                "MCP server HTTP headers";
            layout.Controls.Add(_mcpHeaders, 0, 9);

            layout.Controls.Add(
                FieldLabel(
                    "Edge/Chrome tool allowlist (exact MCP names, " +
                    "one per line; leave blank to disable)"),
                0,
                10);
            _mcpBrowserTools.Dock = DockStyle.Fill;
            _mcpBrowserTools.Multiline = true;
            _mcpBrowserTools.ScrollBars = ScrollBars.Vertical;
            _mcpBrowserTools.MinimumSize = new Size(0, 48);
            _mcpBrowserTools.MaxLength = 2000;
            _mcpBrowserTools.AccessibleName =
                "Read-only MCP tools allowed in browser chat";
            layout.Controls.Add(_mcpBrowserTools, 0, 11);

            _mcpBrowserToolsApproved.AutoSize = true;
            _mcpBrowserToolsApproved.Text =
                "Allow only these tools in Edge/Chrome; I verified " +
                "that they are read-only";
            _mcpBrowserToolsApproved.AccessibleName =
                "Approve listed read-only MCP tools for browser chat";
            layout.Controls.Add(
                _mcpBrowserToolsApproved,
                0,
                12);

            _mcpEnabled.AutoSize = true;
            _mcpEnabled.Checked = true;
            _mcpEnabled.Text = "Enabled";
            _mcpEnabled.AccessibleName =
                "MCP server enabled";
            layout.Controls.Add(_mcpEnabled, 0, 13);

            var buttons = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(0, 8, 0, 0),
                Margin = new Padding(0)
            };
            _mcpSave.Click += McpSaveClick;
            _mcpRemove.Click += McpRemoveClick;
            buttons.Controls.Add(_mcpSave);
            buttons.Controls.Add(_mcpRemove);
            layout.Controls.Add(buttons, 0, 14);

            ConfigureSupportingLabel(_mcpStatus);
            _mcpStatus.Text =
                "Changes apply after you press Save. Up to " +
                McpServerConfig.MaxServers +
                " servers; their tools appear to the model as " +
                "mcp_<server>_<tool>. Browser chat uses at most " +
                "one approved server and one tool call per request.";
            layout.Controls.Add(_mcpStatus, 0, 15);
            page.Controls.Add(layout);
            return page;
        }

        private void RefreshMcpList()
        {
            _mcpList.Items.Clear();
            foreach (var server in _mcpServers)
            {
                _mcpList.Items.Add(
                    server.Name +
                    (server.Enabled ? string.Empty : " (off)") +
                    (server.BrowserToolsApproved &&
                     server.ParsedBrowserTools().Count > 0
                        ? " (browser: " +
                          server.ParsedBrowserTools().Count + ")"
                        : string.Empty) +
                    "  -  " +
                    TextBoundary.SingleLine(server.Target, 80));
            }
        }

        private void McpSelectionChanged(
            object sender,
            EventArgs eventArgs)
        {
            var index = _mcpList.SelectedIndex;
            if (index < 0 || index >= _mcpServers.Count)
            {
                return;
            }

            var server = _mcpServers[index];
            _mcpName.Text = server.Name;
            _mcpTarget.Text = server.Target;
            _mcpArguments.Text = server.Arguments;
            _mcpHeaders.Text = server.Headers;
            _mcpBrowserTools.Text = server.BrowserTools;
            _mcpBrowserToolsApproved.Checked =
                server.BrowserToolsApproved;
            _mcpEnabled.Checked = server.Enabled;
        }

        private void McpSaveClick(
            object sender,
            EventArgs eventArgs)
        {
            var server = new McpServerConfig
            {
                Name = _mcpName.Text,
                Target = _mcpTarget.Text,
                Arguments = _mcpArguments.Text,
                Headers = _mcpHeaders.Text,
                BrowserTools = _mcpBrowserTools.Text,
                BrowserToolsApproved =
                    _mcpBrowserToolsApproved.Checked,
                Enabled = _mcpEnabled.Checked
            }.Sanitized();
            if (server.Target.Length == 0)
            {
                _mcpStatus.Text =
                    "Enter a command path or an HTTP(S) URL " +
                    "for the server.";
                return;
            }

            if (server.BrowserToolsApproved &&
                server.ParsedBrowserTools().Count == 0)
            {
                _mcpStatus.Text =
                    "List at least one exact read-only tool name, " +
                    "or untick browser approval.";
                return;
            }

            var existing = _mcpServers.FindIndex(entry =>
                string.Equals(
                    entry.Name,
                    server.Name,
                    StringComparison.Ordinal));
            if (existing >= 0)
            {
                _mcpServers[existing] = server;
                _mcpStatus.Text =
                    "Updated " + server.Name +
                    ". Press Save to apply.";
            }
            else if (_mcpServers.Count >=
                     McpServerConfig.MaxServers)
            {
                _mcpStatus.Text =
                    "Server limit reached (" +
                    McpServerConfig.MaxServers +
                    "). Remove one first.";
                return;
            }
            else
            {
                _mcpServers.Add(server);
                _mcpStatus.Text =
                    "Added " + server.Name +
                    ". Press Save to apply.";
            }

            if (server.BrowserToolsApproved)
            {
                foreach (var entry in _mcpServers)
                {
                    if (!ReferenceEquals(entry, server) &&
                        !string.Equals(
                            entry.Name,
                            server.Name,
                            StringComparison.Ordinal))
                    {
                        entry.BrowserToolsApproved = false;
                    }
                }

                _mcpStatus.Text =
                    "Approved " + server.Name +
                    " for its listed read-only browser tools. " +
                    "Other browser MCP servers were disabled. " +
                    "Press Save to apply.";
            }

            RefreshMcpList();
        }

        private void McpRemoveClick(
            object sender,
            EventArgs eventArgs)
        {
            var index = _mcpList.SelectedIndex;
            if (index < 0 || index >= _mcpServers.Count)
            {
                _mcpStatus.Text =
                    "Select a server in the list first.";
                return;
            }

            var name = _mcpServers[index].Name;
            _mcpServers.RemoveAt(index);
            RefreshMcpList();
            _mcpStatus.Text =
                "Removed " + name + ". Press Save to apply.";
        }

        private TabPage BuildLimitsPage()
        {
            var page = new TabPage("Limits")
            {
                AutoScroll = true
            };
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 2,
                Padding = new Padding(18, 16, 18, 12)
            };
            layout.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 100));
            layout.ColumnStyles.Add(
                new ColumnStyle(SizeType.AutoSize));

            var intro = SupportingText(
                "How many emails Scribble can work with in a " +
                "single request: the /search and Outlook " +
                "multi-selection working set, mailbox search " +
                "results, and the request-wide cap on message " +
                "bodies. The default of " +
                LimitOverrides.RecommendedWorkingSetMessages +
                " suits local models with modest context " +
                "windows. Raising it sends more mailbox text to " +
                "the model, which can slow requests down or " +
                "overflow a small model's context window. " +
                "Drafting guardrails are unaffected: Scribble " +
                "still never sends email and every draft stays " +
                "yours to review.");
            intro.ForeColor = SystemColors.ControlText;
            layout.Controls.Add(intro, 0, 0);
            layout.SetColumnSpan(intro, 2);

            layout.Controls.Add(
                FieldLabel("Emails per request (working set)"),
                0,
                1);
            _workingSetSize.Minimum =
                LimitOverrides.MinWorkingSetMessages;
            _workingSetSize.Maximum =
                LimitOverrides.MaxWorkingSetMessages;
            _workingSetSize.Value =
                LimitOverrides.RecommendedWorkingSetMessages;
            _workingSetSize.Increment = 10;
            _workingSetSize.Width = 90;
            _workingSetSize.Margin =
                new Padding(0, 6, 0, 0);
            _workingSetSize.AccessibleName =
                "Emails per request (working set)";
            layout.Controls.Add(_workingSetSize, 1, 1);

            var note = SupportingText(
                "Applies after Save, from the next request " +
                "onward.");
            layout.Controls.Add(note, 0, 2);
            layout.SetColumnSpan(note, 2);

            page.Controls.Add(layout);
            return page;
        }

        private TabPage BuildSupportPage()
        {
            var page = new TabPage("Support")
            {
                AutoScroll = true
            };
            var layout = new TableLayoutPanel
            {
                // Scrolls when the rows do not fit instead of
                // squeezing them.
                AutoScroll = true,
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 5,
                Padding = new Padding(18, 16, 18, 12)
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(
                new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var intro = SupportingText(
                "Found a problem or have an idea? Describe it " +
                "below and click Create report email. Settings " +
                "closes and Scribble opens a pre-filled email to the " +
                "creator (r.cunha@samsung.com) with your notes " +
                "and the recent diagnostic log so you can review " +
                "everything and send it yourself from Outlook. " +
                "Scribble never sends anything on its own.");
            intro.ForeColor = SystemColors.ControlText;
            layout.Controls.Add(intro, 0, 0);
            layout.Controls.Add(
                FieldLabel("What happened, or what would help"),
                0,
                1);
            _supportText.Dock = DockStyle.Fill;
            _supportText.BorderStyle = BorderStyle.FixedSingle;
            _supportText.Font = new Font(
                SystemFonts.MessageBoxFont.FontFamily,
                SystemFonts.MessageBoxFont.Size + 1F,
                FontStyle.Regular);
            _supportText.DetectUrls = false;
            _supportText.AccessibleName =
                "Problem or feedback description";
            layout.Controls.Add(_supportText, 0, 2);
            _supportButton.Click += SupportClick;
            var supportRow = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(0, 8, 0, 0),
                Margin = new Padding(0)
            };
            supportRow.Controls.Add(_supportButton);
            layout.Controls.Add(supportRow, 0, 3);
            ConfigureSupportingLabel(_supportStatus);
            _supportStatus.Text =
                "The diagnostic log contains timestamps, " +
                "operation names, and error codes only - never " +
                "email content. Review the draft before sending.";
            layout.Controls.Add(_supportStatus, 0, 4);
            page.Controls.Add(layout);
            return page;
        }

        // Outlook refuses to display a mail item while a modal
        // dialog is open in its process ("Outlook cannot do this
        // because a dialog box is open"), and this settings window
        // IS that modal dialog when opened from the Outlook pane.
        // The click therefore only records the request and closes
        // the window; the pane opens the report email afterwards.
        private void SupportClick(
            object sender,
            EventArgs eventArgs)
        {
            SupportReportRequested = true;
            SupportReportDescription = TextBoundary.PlainText(
                _supportText.Text,
                4000);
            Close();
        }

        public bool SupportReportRequested { get; private set; }

        public string SupportReportDescription
        {
            get;
            private set;
        }

        // Opens the pre-filled, UNSENT report email in Outlook.
        // Called by the panes after this dialog has closed. Returns
        // null on success or a diagnostic message on failure.
        public static string OpenSupportReport(
            object preferredOutlook,
            string description)
        {
            // From the Excel/PowerPoint/Word panes there is no
            // Outlook host object, so the running Outlook (or a
            // fresh instance) is used to open the report draft.
            var outlookApplication = preferredOutlook;
            if (outlookApplication == null)
            {
                try
                {
                    outlookApplication =
                        System.Runtime.InteropServices.Marshal
                            .GetActiveObject(
                                "Outlook.Application");
                }
                catch
                {
                    var outlookType = Type.GetTypeFromProgID(
                        "Outlook.Application");
                    if (outlookType != null)
                    {
                        try
                        {
                            outlookApplication =
                                Activator.CreateInstance(
                                    outlookType);
                        }
                        catch
                        {
                        }
                    }
                }
            }

            if (outlookApplication == null)
            {
                return "[OUTLOOK_NOT_READY] Outlook is not " +
                    "available to open the report email.";
            }

            try
            {
                var bounded = TextBoundary.PlainText(
                    description,
                    4000);
                dynamic application = outlookApplication;
                dynamic mail = application.CreateItem(0);
                mail.To = "r.cunha@samsung.com";
                mail.Subject =
                    "Scribble report - version " +
                    SelfUpdater.InstalledVersion();
                mail.Body =
                    "What happened / feedback:\n" +
                    (bounded.Length > 0
                        ? bounded
                        : "(no description entered)") +
                    "\n\n--- Recent Scribble diagnostic log " +
                    "(review before sending) ---\n" +
                    Log.Tail(120);
                mail.Display(false);
                return null;
            }
            catch (Exception exception)
            {
                return DiagnosticDetails.ForException(
                    exception,
                    "SUPPORT_REPORT_FAILED");
            }
        }

        private void UpdateToneStrengthLabel()
        {
            var value = _toneStrength.Value;
            var flavor = value <= 25
                ? "a light touch"
                : (value <= 55
                    ? "a balanced influence"
                    : (value <= 80
                        ? "a strong voice match"
                        : "mirrors you closely"));
            _toneStrengthValue.Text =
                value + " / 100 - " + flavor;
        }

        private TabPage BuildWritingStylePage()
        {
            var page = new TabPage("Writing soul")
            {
                AutoScroll = true
            };
            var layout = new TableLayoutPanel
            {
                // Scrolls when the rows do not fit instead of
                // squeezing them.
                AutoScroll = true,
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 12,
                Padding = new Padding(18, 16, 18, 12)
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 55));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 52));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 45));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 56));

            var disclosure = SupportingText(
                "Your writing soul is a small portrait of how you " +
                "write - formality, warmth, sentence rhythm, how " +
                "you open and how you sign off. Drafts are shaped " +
                "by it so they sound like you, not like a bot. " +
                "Nothing is analyzed automatically: click Analyze " +
                "and Scribble reads up to 15 recent Sent Items " +
                "messages, removes obvious quoted history, and " +
                "sends bounded samples to your configured model " +
                "to distill the portrait. Review and edit the " +
                "result before saving - it is yours.");
            disclosure.ForeColor = SystemColors.ControlText;
            layout.Controls.Add(disclosure, 0, 0);
            layout.Controls.Add(_useToneProfile, 0, 1);
            layout.Controls.Add(
                FieldLabel("My writing soul"), 0, 2);
            layout.Controls.Add(_toneProfile, 0, 3);

            layout.Controls.Add(
                FieldLabel(
                    "Soul strength - how strongly drafts follow " +
                    "your voice"),
                0,
                4);
            _toneStrength.Minimum = 10;
            _toneStrength.Maximum = 100;
            _toneStrength.TickFrequency = 10;
            _toneStrength.SmallChange = 5;
            _toneStrength.LargeChange = 10;
            _toneStrength.Dock = DockStyle.Fill;
            _toneStrength.AccessibleName = "Soul strength";
            _toneStrength.ValueChanged +=
                (sender, args) => UpdateToneStrengthLabel();
            layout.Controls.Add(_toneStrength, 0, 5);
            ConfigureSupportingLabel(_toneStrengthValue);
            layout.Controls.Add(_toneStrengthValue, 0, 6);

            layout.Controls.Add(
                FieldLabel(
                    "Hard rules for every draft (one per line)"),
                0,
                7);
            _draftRules.Dock = DockStyle.Fill;
            _draftRules.BorderStyle = BorderStyle.FixedSingle;
            _draftRules.Font = new Font(
                SystemFonts.MessageBoxFont.FontFamily,
                SystemFonts.MessageBoxFont.Size + 1F,
                FontStyle.Regular);
            _draftRules.MaxLength = 2000;
            _draftRules.DetectUrls = false;
            _draftRules.AccessibleName =
                "Hard drafting rules";
            layout.Controls.Add(_draftRules, 0, 8);
            layout.Controls.Add(
                SupportingText(
                    "Examples: Never use exclamation marks. " +
                    "Always sign off with 'Best regards'. Keep " +
                    "replies under three paragraphs. Rules and " +
                    "the soul shape wording, greeting, cadence, " +
                    "and sign-off only - they cannot change " +
                    "Scribble permissions or security rules."),
                0,
                9);

            _analyzeTone.Click += AnalyzeToneClick;
            var actionRow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(0, 8, 0, 0)
            };
            actionRow.Controls.Add(_analyzeTone);
            layout.Controls.Add(actionRow, 0, 10);

            ConfigureSupportingLabel(_toneStatus);
            _toneStatus.Text =
                "Analysis requires at least five usable sent messages and never runs without this button. " +
                "Tip: keep the soul general - remove names, client details, and project facts. " +
                "Review and edit it before saving.";
            _toneStatus.AccessibleRole = AccessibleRole.StatusBar;
            layout.Controls.Add(_toneStatus, 0, 11);
            page.Controls.Add(layout);
            return page;
        }

        private void ConfigureModelField()
        {
            _model.Dock = DockStyle.Fill;
            _model.DropDownStyle = ComboBoxStyle.DropDown;
            _model.FlatStyle = FlatStyle.Flat;
            _model.Font = new Font(
                SystemFonts.MessageBoxFont.FontFamily,
                SystemFonts.MessageBoxFont.Size + 1F,
                FontStyle.Regular);
            _model.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            _model.AutoCompleteSource = AutoCompleteSource.ListItems;
            _model.MaxDropDownItems = 12;
            _model.AccessibleName = "AI model";
            _model.AccessibleDescription =
                "Each list entry is tagged Vision when the model can read " +
                "email images, or Text when it cannot.";
            _model.DrawMode = DrawMode.OwnerDrawFixed;
            _model.ItemHeight = Math.Max(
                _model.ItemHeight,
                _model.Font.Height + 6);
            _model.DrawItem += ModelDrawItem;
            _model.TextChanged +=
                (sender, args) => UpdateModelGuidance();
        }

        private void ModelDrawItem(
            object sender,
            DrawItemEventArgs eventArgs)
        {
            eventArgs.DrawBackground();
            if (eventArgs.Index < 0)
            {
                eventArgs.DrawFocusRectangle();
                return;
            }

            var modelId = Convert.ToString(
                _model.Items[eventArgs.Index]) ?? string.Empty;
            var isVision = ModelCatalog.IsVisionCapable(modelId);
            var tag = isVision ? "Vision" : "Text";
            var selected =
                (eventArgs.State & DrawItemState.Selected) ==
                DrawItemState.Selected;
            var bounds = eventArgs.Bounds;
            var tagWidth = TextRenderer.MeasureText(
                eventArgs.Graphics,
                tag,
                eventArgs.Font).Width + 8;
            var idBounds = new Rectangle(
                bounds.Left + 2,
                bounds.Top,
                Math.Max(16, bounds.Width - tagWidth - 8),
                bounds.Height);
            var tagBounds = new Rectangle(
                bounds.Right - tagWidth - 4,
                bounds.Top,
                tagWidth,
                bounds.Height);
            TextRenderer.DrawText(
                eventArgs.Graphics,
                modelId,
                eventArgs.Font,
                idBounds,
                selected
                    ? SystemColors.HighlightText
                    : SystemColors.WindowText,
                TextFormatFlags.Left |
                TextFormatFlags.VerticalCenter |
                TextFormatFlags.EndEllipsis |
                TextFormatFlags.NoPrefix);
            TextRenderer.DrawText(
                eventArgs.Graphics,
                tag,
                eventArgs.Font,
                tagBounds,
                selected
                    ? SystemColors.HighlightText
                    : (isVision ? SuccessText : SecondaryText),
                TextFormatFlags.Right |
                TextFormatFlags.VerticalCenter |
                TextFormatFlags.NoPrefix);
            eventArgs.DrawFocusRectangle();
        }

        private Control BuildButtons()
        {
            var panel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Padding = new Padding(0, 6, 0, 0)
            };
            _save.Click += SaveClick;
            var cancel = MakeButton("Cancel", false, 96);
            cancel.DialogResult = DialogResult.Cancel;
            cancel.Name = "CancelSettings";
            panel.Controls.Add(_save);
            panel.Controls.Add(cancel);
            return panel;
        }

        private static Button GetCancelButton(Control root)
        {
            var matches = root.Controls.Find("CancelSettings", true);
            return matches.Length > 0 ? (Button)matches[0] : null;
        }

        private static void ConfigureField(
            TextBox field,
            string accessibleName,
            string accessibleDescription)
        {
            field.Dock = DockStyle.Fill;
            field.BorderStyle = BorderStyle.FixedSingle;
            field.Font = new Font(
                SystemFonts.MessageBoxFont.FontFamily,
                SystemFonts.MessageBoxFont.Size + 1F,
                FontStyle.Regular);
            field.AccessibleName = accessibleName;
            field.AccessibleDescription = accessibleDescription;
        }

        private static void ConfigureSupportingLabel(Label label)
        {
            label.AutoSize = true;
            label.MaximumSize = new Size(620, 0);
            label.ForeColor = SecondaryText;
        }

        private static Label SupportingText(string text)
        {
            var label = new Label { Text = text };
            ConfigureSupportingLabel(label);
            return label;
        }

        private static Label FieldLabel(string text)
        {
            return new Label
            {
                AutoSize = true,
                Font = new Font(
                    SystemFonts.MessageBoxFont.FontFamily,
                    SystemFonts.MessageBoxFont.Size,
                    FontStyle.Bold),
                ForeColor = SystemColors.ControlText,
                Padding = new Padding(0, 0, 0, 4),
                Text = text
            };
        }

        private static Button MakeButton(
            string text,
            bool primary,
            int width)
        {
            var button = new Button
            {
                Text = text,
                Width = width,
                Height = 32,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(8, 0, 0, 0),
                UseVisualStyleBackColor = false
            };
            button.FlatAppearance.BorderSize = 1;
            button.BackColor = primary
                ? (SystemInformation.HighContrast
                    ? SystemColors.Highlight
                    : Color.FromArgb(0, 95, 184))
                : SystemColors.Window;
            button.ForeColor = primary
                ? SystemColors.HighlightText
                : SystemColors.WindowText;
            button.FlatAppearance.BorderColor = primary
                ? button.BackColor
                : SystemColors.ControlDark;
            button.AccessibleName = text;
            return button;
        }

        private AppSettings ReadFormSettings()
        {
            var profile = TextBoundary.PlainText(
                _toneProfile.Text,
                TextBoundary.MaxToneProfileCharacters);
            return new AppSettings
            {
                BaseUrl = _endpoint.Text,
                Model = _model.Text,
                ApiKey = _apiKey.Text,
                AllowInsecureHttp = true,
                UseGeminiSignIn =
                    !AdminPolicy.GeminiDisabled &&
                    _useGeminiSignIn.Checked,
                GeminiRefreshToken = AdminPolicy.GeminiDisabled
                    ? string.Empty
                    : _geminiRefreshToken,
                GeminiProject = AdminPolicy.GeminiDisabled
                    ? string.Empty
                    : _geminiProject.Text,
                ToneProfile = profile,
                UseToneProfile =
                    _useToneProfile.Checked && profile.Length > 0,
                ToneStrength = _toneStrength.Value,
                DraftRules = TextBoundary.PlainText(
                    _draftRules.Text,
                    2000),
                SwitchToVisionModelForImages =
                    _switchVisionForImages.Checked,
                DiscoveredModels = CollectDiscoveredModels(),
                McpServers = _mcpServers
                    .Select(server => server.Sanitized())
                    .ToList(),
                UseRecommendedLimits = true,
                LimitContextMultiplier = 1,
                LimitPromptCharacters = TextBoundary
                    .RecommendedUserPromptCharacters,
                LimitAssistantCharacters = TextBoundary
                    .RecommendedAssistantCharacters,
                LimitHistoryTurns = TextBoundary
                    .RecommendedConversationTurns,
                LimitToolRounds = TextBoundary
                    .RecommendedToolRounds,
                LimitToolCallsPerRound = TextBoundary
                    .RecommendedToolCallsPerRound,
                LimitWorkingSetMessages =
                    (int)_workingSetSize.Value
            };
        }

        private List<string> CollectDiscoveredModels()
        {
            return _model.Items
                .Cast<object>()
                .Select(item =>
                    TextBoundary.PlainText(
                        Convert.ToString(item),
                        200))
                .Where(ModelSelectionPolicy.IsGenerativeModel)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(item => item, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private void RestoreDiscoveredModels(
            IEnumerable<string> models)
        {
            foreach (var model in models ?? Enumerable.Empty<string>())
            {
                if (!ModelSelectionPolicy.IsGenerativeModel(model))
                {
                    continue;
                }

                if (!_model.Items.Cast<object>().Any(item =>
                    string.Equals(
                        item.ToString(),
                        model,
                        StringComparison.OrdinalIgnoreCase)))
                {
                    _model.Items.Add(model);
                }
            }

            SelectPreferredModelIfEmpty();
        }

        private async void AnalyzeToneClick(
            object sender,
            EventArgs eventArgs)
        {
            if (_analyzingTone)
            {
                _toneCancellation?.Cancel();
                return;
            }

            _error.Text = string.Empty;
            if (_outlookApplication == null)
            {
                _error.Text =
                    "[OUTLOOK_NOT_READY] Open Scribble from Outlook before analyzing your writing style.";
                return;
            }

            var settings = ReadFormSettings();
            if (!settings.IsConfigured)
            {
                _error.Text =
                    "[CONFIGURATION_INCOMPLETE] Configure the endpoint, model, and API key first.";
                return;
            }

            SetToneAnalyzing(true);
            _toneCancellation = new CancellationTokenSource();
            try
            {
                _toneStatus.Text =
                    "Reading and cleaning up to 15 recent Sent Items messages...";
                var samples = new SentMailToneSampler(
                    _outlookApplication).CaptureRecent();
                if (samples.Count < 5)
                {
                    throw new AiEndpointException(
                        "TONE_SAMPLES_INSUFFICIENT",
                        "At least five usable sent messages are required. Found " +
                        samples.Count + ".");
                }

                _toneStatus.Text =
                    "Analyzing " + samples.Count +
                    " bounded sent-email samples with " + settings.Model + "...";
                var request = ToneProfileRequestFactory.Create(
                    settings.Model,
                    samples);
                var response = await _client.CompleteAsync(
                    settings,
                    request,
                    _toneCancellation.Token);
                if (response.tool_calls != null &&
                    response.tool_calls.Count > 0)
                {
                    throw new AiEndpointException(
                        "TONE_RESPONSE_INVALID",
                        "The model returned tool calls during style analysis.");
                }

                var profile = SafeModelText.Format(
                    response.content,
                    TextBoundary.MaxToneProfileCharacters).PlainText;
                if (profile.Length == 0)
                {
                    throw new AiEndpointException(
                        "TONE_RESPONSE_EMPTY",
                        "The model returned an empty writing profile.");
                }

                _toneProfile.Text = profile;
                _useToneProfile.Checked = true;
                _toneStatus.ForeColor = SuccessText;
                _toneStatus.Text =
                    "Writing profile generated from " + samples.Count +
                    " sent messages. Review and edit it, then click Save.";
            }
            catch (OperationCanceledException)
            {
                _toneStatus.Text =
                    "Writing-style analysis was cancelled.";
            }
            catch (Exception exception)
            {
                _error.Text = DiagnosticDetails.ForException(
                    exception,
                    "TONE_ANALYSIS_FAILED");
            }
            finally
            {
                _toneCancellation?.Dispose();
                _toneCancellation = null;
                SetToneAnalyzing(false);
            }
        }

        private async void CheckEndpointClick(
            object sender,
            EventArgs eventArgs)
        {
            if (_checking)
            {
                _checkCancellation?.Cancel();
                return;
            }

            if (_refreshingModels)
            {
                return;
            }

            _error.Text = string.Empty;
            var settings = ReadFormSettings();
            if (!settings.HasConnectionSettings)
            {
                _error.Text =
                    "[CONFIGURATION_INCOMPLETE] Enter a valid endpoint and API key first.";
                return;
            }

            SetChecking(true);
            _checkCancellation = new CancellationTokenSource();
            try
            {
                IReadOnlyList<string> models = null;
                var discoveryNote = string.Empty;
                try
                {
                    _testStatus.Text =
                        "Checking authentication and available models " +
                        "(up to " + ModelDiscoveryTimeoutSeconds +
                        " seconds)...";
                    using (var modelsTimeout =
                        CancellationTokenSource.CreateLinkedTokenSource(
                            _checkCancellation.Token))
                    {
                        modelsTimeout.CancelAfter(
                            TimeSpan.FromSeconds(
                                ModelDiscoveryTimeoutSeconds));
                        models = await _client.GetModelsAsync(
                            settings,
                            modelsTimeout.Token);
                    }

                    AddDiscoveredModels(models);
                }
                catch (OperationCanceledException)
                {
                    if (_checkCancellation.IsCancellationRequested)
                    {
                        throw;
                    }

                    discoveryNote =
                        " Model discovery timed out after " +
                        ModelDiscoveryTimeoutSeconds +
                        " seconds, so the entered model was tested directly.";
                }
                catch (AiEndpointException exception)
                {
                    discoveryNote =
                        " Model discovery was unavailable [" + exception.Code +
                        "], so the entered model was tested directly.";
                }

                settings = ReadFormSettings();
                if (settings.Model.Trim().Length == 0)
                {
                    throw new AiEndpointException(
                        "MODEL_REQUIRED",
                        "Choose a model or use Connect & load models after discovery returns at least one generative model.");
                }

                _testStatus.Text =
                    "Testing mailbox tool calls with " + settings.Model +
                    " (up to " + EndpointProbeTimeoutSeconds +
                    " seconds)...";
                using (var probeTimeout =
                    CancellationTokenSource.CreateLinkedTokenSource(
                        _checkCancellation.Token))
                {
                    probeTimeout.CancelAfter(
                        TimeSpan.FromSeconds(
                            EndpointProbeTimeoutSeconds));
                    var probe = ChatRequestFactory.CreateEndpointCheck(
                        settings.Model);
                    var response = await _client.CompleteAsync(
                        settings,
                        probe,
                        probeTimeout.Token);
                    var validCall = response.tool_calls != null &&
                        response.tool_calls.Any(call =>
                            call?.function != null &&
                            MailboxToolCatalog.IsApproved(
                                call.function.name));
                    if (!validCall)
                    {
                        throw new AiEndpointException(
                            "MODEL_TOOL_CALL_UNSUPPORTED",
                            "The endpoint answered, but this model did not return a compatible mailbox tool call.");
                    }
                }

                _testStatus.ForeColor = SuccessText;
                _testStatus.Text =
                    "Endpoint verified. Authentication, model, and mailbox tool calling passed." +
                    discoveryNote;
            }
            catch (OperationCanceledException)
            {
                _error.Text =
                    "[ENDPOINT_CHECK_CANCELLED] The endpoint check was cancelled or timed out. " +
                    "Model discovery allows up to " +
                    ModelDiscoveryTimeoutSeconds +
                    " seconds. The tool-call probe allows up to " +
                    EndpointProbeTimeoutSeconds +
                    " seconds.";
            }
            catch (Exception exception)
            {
                _error.Text = DiagnosticDetails.ForException(
                    exception,
                    "ENDPOINT_CHECK_FAILED");
            }
            finally
            {
                _checkCancellation?.Dispose();
                _checkCancellation = null;
                SetChecking(false);
            }
        }

        private async void RefreshModelsClick(
            object sender,
            EventArgs eventArgs)
        {
            if (_refreshingModels)
            {
                _modelRefreshCancellation?.Cancel();
                return;
            }

            if (_checking ||
                _analyzingTone ||
                _updating)
            {
                return;
            }

            _error.Text = string.Empty;
            var settings = ReadFormSettings();
            if (!settings.HasConnectionSettings)
            {
                _error.Text =
                    "[CONFIGURATION_INCOMPLETE] Enter a valid endpoint and API key first.";
                return;
            }

            _refreshingModels = true;
            _refreshModels.Text = "Cancel connection";
            SetCommonControlsEnabled(false);
            _modelRefreshCancellation =
                new CancellationTokenSource(
                    TimeSpan.FromSeconds(
                        ModelDiscoveryTimeoutSeconds));
            try
            {
                _testStatus.ForeColor = SecondaryText;
                _testStatus.Text =
                    "Loading models from the endpoint (up to " +
                    ModelDiscoveryTimeoutSeconds +
                    " seconds)...";
                var models = await _client.GetModelsAsync(
                    settings,
                    _modelRefreshCancellation.Token);
                var available = ReplaceDiscoveredModels(models);
                _testStatus.ForeColor = SuccessText;
                _testStatus.Text =
                    "Model list loaded: " +
                    available.ToString() +
                    " models available.";
            }
            catch (OperationCanceledException)
            {
                _error.Text =
                    "[MODEL_REFRESH_CANCELLED] Model discovery was cancelled " +
                    "or timed out after " +
                    ModelDiscoveryTimeoutSeconds +
                    " seconds.";
            }
            catch (Exception exception)
            {
                _error.Text = DiagnosticDetails.ForException(
                    exception,
                    "MODEL_REFRESH_FAILED");
            }
            finally
            {
                _modelRefreshCancellation?.Dispose();
                _modelRefreshCancellation = null;
                _refreshingModels = false;
                _refreshModels.Text = "Connect & load models";
                SetCommonControlsEnabled(
                    !_checking &&
                    !_analyzingTone &&
                    !_updating);
            }
        }

        private int AddDiscoveredModels(IEnumerable<string> models)
        {
            var added = 0;
            foreach (var model in models ?? Enumerable.Empty<string>())
            {
                if (!ModelSelectionPolicy.IsGenerativeModel(model))
                {
                    continue;
                }

                if (!_model.Items.Cast<object>().Any(item =>
                    string.Equals(
                        item.ToString(),
                        model,
                        StringComparison.OrdinalIgnoreCase)))
                {
                    _model.Items.Add(model);
                    added++;
                }
            }

            SelectPreferredModelIfEmpty();
            UpdateModelGuidance();
            return added;
        }

        private int ReplaceDiscoveredModels(
            IEnumerable<string> models)
        {
            var discovered = (models ?? Enumerable.Empty<string>())
                .Where(ModelSelectionPolicy.IsGenerativeModel)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (discovered.Count == 0)
            {
                return 0;
            }

            var selected = _model.Text.Trim();
            _model.BeginUpdate();
            try
            {
                _model.Items.Clear();
                foreach (var model in discovered)
                {
                    _model.Items.Add(model);
                }
            }
            finally
            {
                _model.EndUpdate();
            }

            _model.Text = discovered.Any(model =>
                string.Equals(
                    model,
                    selected,
                    StringComparison.OrdinalIgnoreCase))
                        ? selected
                        : string.Empty;
            SelectPreferredModelIfEmpty();
            UpdateModelGuidance();
            return discovered.Count;
        }

        private void SelectPreferredModelIfEmpty()
        {
            if (_model.Text.Trim().Length > 0 ||
                _model.Items.Count == 0)
            {
                return;
            }

            var models = _model.Items
                .Cast<object>()
                .Select(item => Convert.ToString(item))
                .Where(ModelSelectionPolicy.IsGenerativeModel)
                .ToList();
            var preferred = ModelSelectionPolicy.PreferredModel(models);
            if (!string.IsNullOrWhiteSpace(preferred))
            {
                _model.Text = preferred;
            }
        }

        private void SetChecking(bool checking)
        {
            _checking = checking;
            _checkEndpoint.Text =
                checking ? "Cancel test" : "Test selected model";
            SetCommonControlsEnabled(!checking && !_analyzingTone);
            if (checking)
            {
                _testStatus.ForeColor = SecondaryText;
            }
        }

        private void SetToneAnalyzing(bool analyzing)
        {
            _analyzingTone = analyzing;
            _analyzeTone.Text =
                analyzing ? "Cancel analysis" : "Analyze 15 sent emails";
            SetCommonControlsEnabled(!analyzing && !_checking);
            _analyzeTone.Enabled = analyzing || !_checking;
            if (analyzing)
            {
                _toneStatus.ForeColor = SecondaryText;
            }
        }

        private void SetCommonControlsEnabled(bool enabled)
        {
            _save.Enabled = enabled;
            _endpoint.Enabled = enabled;
            _model.Enabled = enabled;
            _apiKey.Enabled = enabled;
            _switchVisionForImages.Enabled = enabled;
            _useToneProfile.Enabled = enabled;
            _toneProfile.Enabled = enabled;
            _checkEndpoint.Enabled = enabled || _checking;
            _refreshModels.Enabled = _refreshingModels ||
                (enabled && !_checking);
            _analyzeTone.Enabled = enabled || _analyzingTone;
            _updateButton.Enabled = enabled && !_updating;
            _commonControlsEnabled = enabled;
            UpdateGeminiModeUi();
        }

        private void SaveClick(object sender, EventArgs eventArgs)
        {
            try
            {
                _error.Text = string.Empty;
                var settings = ReadFormSettings();
                _store.Save(settings);
                SavedSettings = settings;
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception exception)
            {
                _error.Text = DiagnosticDetails.ForException(
                    exception,
                    "SETTINGS_SAVE_FAILED");
            }
        }

        private void SettingsWindowFormClosing(
            object sender,
            FormClosingEventArgs eventArgs)
        {
            if (!_checking &&
                !_refreshingModels &&
                !_analyzingTone &&
                !_updating)
            {
                return;
            }

            eventArgs.Cancel = true;
            _checkCancellation?.Cancel();
            _modelRefreshCancellation?.Cancel();
            _toneCancellation?.Cancel();
            _updateCancellation?.Cancel();
            _error.Text =
                "Cancelling the active settings operation. Close again when it finishes.";
        }

        private void GeminiModeChanged(
            object sender,
            EventArgs eventArgs)
        {
            if (_useGeminiSignIn.Checked &&
                _model.Text.Trim().Length == 0)
            {
                _model.Text = "gemini-2.5-flash";
            }

            UpdateGeminiModeUi();
        }

        private void UpdateGeminiModeUi()
        {
            if (AdminPolicy.GeminiDisabled)
            {
                _useGeminiSignIn.Checked = false;
                _useGeminiSignIn.Enabled = false;
                _googleSignIn.Enabled = false;
                _geminiProject.Enabled = false;
                _endpoint.Enabled = _commonControlsEnabled;
                _apiKey.Enabled = _commonControlsEnabled;
                _googleStatus.Text =
                    "Disabled by administrator policy " +
                    "(Software\\Policies\\Scribble, DisableGemini).";
                return;
            }

            var gemini = _useGeminiSignIn.Checked;
            var baseline = _commonControlsEnabled;
            // Both transports coexist: the tick only adds Gemini
            // models to the picker, so the endpoint stays editable
            // and usable for local models at the same time.
            _endpoint.Enabled = baseline;
            _apiKey.Enabled = baseline;
            _useGeminiSignIn.Enabled = baseline && !_signingIn;
            _geminiProject.Enabled = baseline && gemini;
            _googleSignIn.Enabled =
                baseline && gemini && !_signingIn;
            if (!gemini)
            {
                _googleStatus.Text =
                    "Off - only the endpoint's own models are " +
                    "offered.";
            }
            else if (_signingIn)
            {
                _googleStatus.Text =
                    "Complete the Google sign-in in the browser " +
                    "window...";
            }
            else if (_geminiRefreshToken.Trim().Length > 0)
            {
                _googleStatus.Text =
                    "Signed in with Google. Click Save to keep " +
                    "this sign-in. Gemini models now appear in " +
                    "the model list alongside your endpoint's " +
                    "own models - the model you pick decides " +
                    "where each request goes.";
            }
            else
            {
                _googleStatus.Text =
                    "Not signed in yet. Email context will go to " +
                    "Google Gemini after sign-in.";
            }

            if (gemini && _model.Items.Count == 0)
            {
                RestoreDiscoveredModels(
                    GeminiCodeAssistGateway.KnownModels);
            }
        }

        private async void GoogleSignInClick(
            object sender,
            EventArgs eventArgs)
        {
            if (_signingIn ||
                _checking ||
                _refreshingModels ||
                _updating)
            {
                return;
            }

            _signingIn = true;
            _error.Text = string.Empty;
            UpdateGeminiModeUi();
            try
            {
                using (var httpClient =
                    new System.Net.Http.HttpClient())
                {
                    var result = await GoogleSignInFlow
                        .SignInAsync(
                            httpClient,
                            TimeSpan.FromMinutes(3));
                    _geminiRefreshToken = result.RefreshToken;
                    _client.GeminiGateway.PrimeAccessToken(
                        result.AccessToken,
                        result.ExpiresInSeconds);
                }

                _googleStatus.Text =
                    "Signed in. Verifying Gemini access...";
                var models = await _client.GetModelsAsync(
                    ReadFormSettings(),
                    CancellationToken.None);
                RestoreDiscoveredModels(models);
                if (_model.Text.Trim().Length == 0)
                {
                    _model.Text = "gemini-2.5-flash";
                }

                _error.Text = string.Empty;
            }
            catch (Exception exception)
            {
                _error.Text = DiagnosticDetails.ForException(
                    exception,
                    "GOOGLE_SIGNIN_FAILED");
            }
            finally
            {
                _signingIn = false;
                UpdateGeminiModeUi();
            }
        }

        private void UpdateModelGuidance()
        {
            var text = _model.Text.Trim();
            if (text.Length == 0)
            {
                _modelGuidance.Text =
                    "Use Connect & load models, then choose a model that supports tool calls. " +
                    "Entries tagged Vision can read email images; Text entries cannot.";
                return;
            }

            var capability = ModelCatalog.IsDisallowedModel(text)
                ? string.Empty
                : (ModelCatalog.IsVisionCapable(text)
                    ? "Vision-capable: reads email image attachments. "
                    : "Text-only: email images stay filename-only. ");
            _modelGuidance.Text = capability +
                ModelSelectionPolicy.DescriptionFor(text);
        }

        private void UpdateTransportWarning()
        {
            Uri endpoint;
            if (Uri.TryCreate(
                    _endpoint.Text.Trim(),
                    UriKind.Absolute,
                    out endpoint) &&
                endpoint.Scheme.Equals(
                    Uri.UriSchemeHttp,
                    StringComparison.OrdinalIgnoreCase) &&
                !endpoint.IsLoopback)
            {
                _transportWarning.ForeColor = ErrorText;
                _transportWarning.Text =
                    "Warning: this remote HTTP endpoint receives the API key, " +
                    "prompts, and retrieved email context without transport encryption.";
                return;
            }

            _transportWarning.ForeColor = SecondaryText;
            _transportWarning.Text = endpoint != null &&
                endpoint.Scheme.Equals(
                    Uri.UriSchemeHttp,
                    StringComparison.OrdinalIgnoreCase)
                    ? "Local HTTP is supported by default. Traffic to a loopback endpoint stays on this computer."
                    : "HTTPS is recommended for endpoints outside this computer.";
        }
    }
}

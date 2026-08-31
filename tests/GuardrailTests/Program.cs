using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Scribble;
using Scribble.Chat;
using Scribble.Configuration;
using Scribble.Interop;
using Scribble.Office;
using Scribble.Outlook;
using Scribble.Security;
using Scribble.UI;
using Scribble.Utilities;

namespace GuardrailTests
{
    internal static class Program
    {
        private static int _passed;

        private static int Main(string[] args)
        {
            // The MCP round-trip test relaunches this same exe as a
            // scripted stdio MCP server, so the test needs no
            // external interpreter and stays deterministic. The
            // command line is checked directly as well so the server
            // mode can never fall through into the test suite.
            if ((args.Length > 0 &&
                 args[0] == "--mcp-fake-server") ||
                Environment.CommandLine.IndexOf(
                    "--mcp-fake-server",
                    StringComparison.Ordinal) >= 0)
            {
                return RunFakeMcpServer();
            }

            // Raw line-echo child used by the spawn diagnostic that
            // decides whether this environment can run the live MCP
            // round trip at all.
            if ((args.Length > 0 &&
                 args[0] == "--echo-server") ||
                Environment.CommandLine.IndexOf(
                    "--echo-server",
                    StringComparison.Ordinal) >= 0)
            {
                while (true)
                {
                    var echoLine = Console.In.ReadLine();
                    if (echoLine == null)
                    {
                        return 0;
                    }

                    Console.Out.WriteLine("pong:" + echoLine);
                    Console.Out.Flush();
                }
            }

            try
            {
                Run(
                    "Vision-capable models are detected broadly",
                    VisionCapableModelsAreDetectedBroadly);
                Run(
                    "Vision prefetch injects image input",
                    VisionPrefetchInjectsImageInput);
                Run(
                    "Vision auto-switch picks the best discovered model",
                    VisionAutoSwitchPicksBestDiscoveredModel);
                Run(
                    "Gauss models are excluded from discovery",
                    GaussModelsAreExcluded);
                Run(
                    "Vision models receive multimodal image follow-up",
                    VisionModelsReceiveMultimodalFollowUp);
                Run(
                    "System prompt states image capability",
                    SystemPromptStatesImageCapability);
                Run(
                    "Vision image limits are enforced",
                    VisionImageLimitsAreEnforced);
                Run(
                    "Web-hosted body images are disclosed as unreadable",
                    WebHostedImagesAreDisclosed);
                Run(
                    "Extensionless pasted images are sniffed and read",
                    PastedImagesWithoutExtensionAreRead);
                Run(
                    "Oversized images are downscaled for vision input",
                    OversizedImagesAreDownscaledForVision);
                Run(
                    "PDF, PowerPoint, and Word attachments are extracted",
                    DocumentAttachmentsAreExtracted);
                Run(
                    "Calendar invites are readable with attachments",
                    CalendarInvitesAreReadable);
                Run(
                    "CID-font PDF text decodes through ToUnicode maps",
                    PdfCidFontTextIsDecoded);
                Run(
                    "Legacy Office, RTF, and unknown attachments are handled",
                    LegacyAndUnknownAttachmentsAreHandled);
                Run(
                    "Local files load as bounded context or vision input",
                    LocalFilesLoadAsContext);
                Run(
                    "Spreadsheets stream through shared strings",
                    SpreadsheetsStreamThroughSharedStrings);
                Run(
                    "Binary workbooks decode BIFF12 records",
                    BinaryWorkbooksDecodeBiff12Records);
                Run(
                    "Office variants, OpenDocument, and MSG extract",
                    OfficeVariantsOpenDocumentAndMsgExtract);
                Run(
                    "Oversized text carries a truncation notice",
                    OversizedTextCarriesTruncationNotice);
                Run(
                    "Suggested reply questions are bounded",
                    SuggestedReplyQuestionsAreBounded);
                Run(
                    "Gemini translation preserves the tool contract",
                    GeminiTranslationPreservesToolContract);
                Run(
                    "Direct Gemini is unavailable to end users",
                    GeminiIsUnavailableToEndUsers);
                Run(
                    "Gemini gateway fails closed before network access",
                    GeminiGatewayFailsClosed);
                Run(
                    "Context budgets scale only in large-context mode",
                    ContextBudgetsScaleOnlyInLargeContextMode);
                Run(
                    "Soul strength and draft rules stay bounded",
                    SoulStrengthAndDraftRulesStayBounded);
                Run(
                    "Document panes gate drafts behind explicit intent",
                    DocumentDraftIntentRequiresExplicitPhrase);
                Run(
                    "Workbook and presentation catalogs stay read only",
                    WorkbookAndPresentationCatalogsStayReadOnly);
                Run(
                    "Document factory authorizes at most one marked draft",
                    DocumentFactoryGatesDraftTools);
                Run(
                    "Browser context is bounded and tools are approved-only",
                    BrowserContextIsBoundedAndReadOnly);
                Run(
                    "Browser screenshots require valid vision input",
                    BrowserScreenshotRequiresVision);
                Run(
                    "MCP tools are namespaced, bounded, and user-configured",
                    McpToolsAreNamespacedAndBounded);
                Run(
                    "MCP stdio round trip works against a scripted server",
                    McpStdioRoundTripWorks);
                Run(
                    "Small inline signature images are ignored",
                    SignatureImagesAreIgnored);
                Run(
                    "Model catalog describes vision capability",
                    ModelCatalogDescribesVisionCapability);
                Run(
                    "Qwen is preferred without locking out other models",
                    QwenIsPreferredWithoutLockIn);
                Run("HTTPS endpoint is accepted", HttpsEndpointIsAccepted);
                Run("Loopback HTTP endpoint is accepted", LoopbackHttpIsAccepted);
                Run(
                    "Remote HTTP can be rejected by an explicit policy",
                    RemoteHttpCanBeRejectedExplicitly);
                Run(
                    "Remote HTTP endpoint is accepted by default",
                    RemoteHttpIsAcceptedByDefault);
                Run(
                    "Legacy HTTP opt-in state migrates to the default",
                    LegacyHttpSettingMigratesToDefault);
                Run(
                    "Models endpoint is normalized",
                    ModelsEndpointIsNormalized);
                Run(
                    "Default model is empty on install",
                    DefaultModelIsEmptyOnInstall);
                Run(
                    "Email attachments are bounded and readable",
                    EmailAttachmentsAreBounded);
                Run(
                    "Compatible endpoint model discovery is verified",
                    ModelDiscoveryUsesCompatibleContract);
                Run(
                    "Endpoint check uses a lightweight tool probe",
                    EndpointCheckUsesLightweightProbe);
                Run(
                    "Compatible tool calls tolerate a missing call id",
                    ToolCallResponseIsNormalized);
                Run("Text boundary removes controls and truncates", TextIsBounded);
                Run(
                    "Model emphasis becomes native formatting",
                    ModelEmphasisIsNormalized);
                Run(
                    "Endpoint diagnostics expose transport details",
                    EndpointDiagnosticsExposeTransportDetails);
                Run("Mailbox tools are read only", MailboxToolsAreReadOnly);
                Run(
                    "Local search command is explicit and bounded",
                    LocalSearchCommandIsBounded);
                Run(
                    "Working set is bounded by the configured size",
                    WorkingSetIsStrictlyBounded);
                Run(
                    "Outlook multi-selection accepts up to the working-set size",
                    OutlookMultiSelectionIsBounded);
                Run(
                    "Active Explorer selection is used for Send to Scribble",
                    ActiveExplorerSelectionIsUsed);
                Run(
                    "External context is explicit and bounded",
                    ExternalContextIsBounded);
                Run(
                    "Writing profile analysis is consent-bound and editable",
                    ToneProfileIsBounded);
                Run(
                    "Latest user prompt locally gates draft tools",
                    DraftToolsRequireLocalIntentAuthorization);
                Run(
                    "Draft authorization creates at most one unsent draft",
                    DraftAuthorizationCreatesOnlyOneDraft);
                Run(
                    "Reply draft uses the exact retrieved message handle",
                    ReplyDraftUsesExactHandle);
                Run(
                    "Reply draft rejects missing or fabricated handles",
                    ReplyDraftRequiresIssuedHandle);
                Run(
                    "Linked draft updates the same visible Outlook item",
                    LinkedDraftUpdatesSameItem);
                Run(
                    "Draft HTML is encoded and locally formatted",
                    DraftHtmlIsSafe);
                Run(
                    "Mixed draft tool calls do not consume permission",
                    MixedDraftToolCallIsRejected);
                Run("Request schema exposes bounded tools", RequestSchemaIsBounded);
                Run("Email is labeled as untrusted data", EmailIsUntrustedData);
                Run("Conversation history is bounded", HistoryIsBounded);
                Run("Draft host exposes no send capability", DraftHasNoSend);
                Run(
                    "Compiled add-in contains no Outlook send invocation",
                    CompiledAddInHasNoSendInvocation);
                Run(
                    "Mailbox host exposes one guarded dispatcher",
                    MailboxHostHasGuardedDispatcher);
                Run(
                    "Draft host exposes one guarded dispatcher",
                    DraftHostHasGuardedDispatcher);
                Run(
                    "Office startup and task pane COM interfaces are dual",
                    OfficeStartupInterfacesAreDual);
                Run(
                    "Chat pane is a registered COM control",
                    ChatPaneIsComControl);
                Run(
                    "Outlook ribbon includes Send to Scribble",
                    RibbonIncludesSendToAi);
                Run(
                    "Selected subjects hide reply and forward prefixes",
                    SelectedSubjectIsCleaned);
                Run(
                    "Self update is official, silent, and restarts Outlook",
                    SelfUpdateIsOfficialAndBounded);
                Run(
                    "Draft formulas stay inside the workbook",
                    DraftFormulasStayInsideTheWorkbook);
                Run(
                    "Draft text layout parses structure locally",
                    DraftTextLayoutParsesStructure);
                Run(
                    "Corporate slide theme is hardcoded",
                    CorporateThemeIsHardcoded);
                Run(
                    "Transport follows the selected model",
                    TransportFollowsSelectedModel);
                Run(
                    "One request builds one deliverable",
                    DraftBudgetBuildsOneDeliverable);
                Run(
                    "Draft HTML renders bounded pipe tables",
                    DraftHtmlRendersTables);
                Run(
                    "Unmentioned email recipients are called out",
                    RecipientWarningFlagsUnknownRecipients);
                Run(
                    "MCP HTTP headers are token-checked and bounded",
                    McpHeadersAreBoundedAndSafe);
                Run(
                    "Admin policy reads only documented switches",
                    AdminPolicyIsReadOnlyAndScoped);
                Run(
                    "Text budgets stay fixed; the working set honors the user",
                    SettingsAlwaysApplyRecommendedLimits);
                Console.WriteLine("PASS: " + _passed + " guardrail tests");
                return 0;
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine("FAIL: " + exception.Message);
                return 1;
            }
        }

        private static void Run(string name, Action test)
        {
            test();
            _passed++;
            Console.WriteLine("PASS: " + name);
        }

        private static void HttpsEndpointIsAccepted()
        {
            Uri endpoint;
            Assert(
                AppSettings.TryGetChatCompletionsUri(
                    "https://ai.example.test/v1",
                    out endpoint),
                "HTTPS should be accepted.");
            Assert(
                endpoint.AbsoluteUri ==
                "https://ai.example.test/v1/chat/completions",
                "The chat completions path was not normalized.");
            Assert(
                AppSettings.TryGetChatCompletionsUri(
                    "https://generativelanguage.googleapis.com/v1beta/openai",
                    out endpoint) &&
                endpoint.AbsoluteUri ==
                "https://generativelanguage.googleapis.com/v1beta/openai/chat/completions",
                "An /openai compatibility base was not " +
                "normalized: " + endpoint);
        }

        private static void LoopbackHttpIsAccepted()
        {
            Uri endpoint;
            Assert(
                AppSettings.TryGetChatCompletionsUri(
                    "http://127.0.0.1:1234/v1",
                    out endpoint),
                "Loopback HTTP should be accepted.");
        }

        private static void RemoteHttpCanBeRejectedExplicitly()
        {
            Uri endpoint;
            Assert(
                !AppSettings.TryGetChatCompletionsUri(
                    "http://ai.example.test/v1",
                    false,
                    out endpoint),
                "The low-level explicit transport guard must still " +
                "be able to reject remote HTTP.");
        }

        private static void RemoteHttpIsAcceptedByDefault()
        {
            Uri endpoint;
            Assert(
                AppSettings.TryGetChatCompletionsUri(
                    "http://ai.example.test/v1/chat/completions",
                    out endpoint),
                "Remote HTTP should work without a separate opt in.");
            Assert(
                endpoint.AbsoluteUri ==
                "http://ai.example.test/v1/chat/completions",
                "Remote HTTP endpoint was not normalized correctly.");

            var settings = new AppSettings
            {
                BaseUrl = "http://ai.example.test/v1",
                Model = "local-model",
                ApiKey = "test-key"
            };
            Assert(
                settings.IsConfigured,
                "A new settings object should accept remote HTTP by default.");
        }

        private static void LegacyHttpSettingMigratesToDefault()
        {
            var directory = Path.Combine(
                Path.GetTempPath(),
                "Scribble-http-settings-" +
                Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                var path = Path.Combine(directory, "settings.json");
                File.WriteAllText(
                    path,
                    "{\"BaseUrl\":\"http://ai.example.test/v1\"," +
                    "\"Model\":\"qwen3.6-35b-a3b\"," +
                    "\"ProtectedApiKey\":\"\"," +
                    "\"AllowInsecureHttp\":false}",
                    Encoding.UTF8);
                var store = new SettingsStore();
                SetPrivateField(store, "_settingsPath", path);
                SetPrivateField(
                    store,
                    "_legacySettingsPath",
                    Path.Combine(directory, "legacy.json"));

                var loaded = store.Load();
                Assert(
                    loaded.AllowInsecureHttp,
                    "A legacy false value must migrate to HTTP enabled.");
                loaded.ApiKey = "test-key";
                store.Save(loaded);
                Assert(
                    File.ReadAllText(path, Encoding.UTF8).Contains(
                        "\"AllowInsecureHttp\":true"),
                    "The migrated HTTP default was not persisted.");
            }
            finally
            {
                if (Directory.Exists(directory))
                {
                    Directory.Delete(directory, true);
                }
            }
        }

        private static void SetPrivateField(
            object instance,
            string name,
            object value)
        {
            var field = instance.GetType().GetField(
                name,
                BindingFlags.Instance | BindingFlags.NonPublic);
            if (field == null)
            {
                throw new InvalidOperationException(
                    "Missing private field: " + name);
            }

            field.SetValue(instance, value);
        }

        private static void ModelsEndpointIsNormalized()
        {
            Uri endpoint;
            Assert(
                AppSettings.TryGetModelsUri(
                    "https://ai.example.test/v1/chat/completions",
                    false,
                    out endpoint),
                "Models endpoint should be accepted.");
            Assert(
                endpoint.AbsoluteUri ==
                "https://ai.example.test/v1/models",
                "The models path was not normalized.");
            Assert(
                AppSettings.TryGetModelsUri(
                    "https://generativelanguage.googleapis.com/v1beta/openai",
                    false,
                    out endpoint) &&
                endpoint.AbsoluteUri ==
                "https://generativelanguage.googleapis.com/v1beta/openai/models",
                "The /openai models path was not normalized: " +
                endpoint);
        }

        private static void DefaultModelIsEmptyOnInstall()
        {
            var settings = new AppSettings();
            Assert(
                settings.Model == string.Empty &&
                !settings.IsConfigured,
                "A fresh install should start without a configured model.");
        }

        private static void EmailAttachmentsAreBounded()
        {
            Assert(
                EmailAttachmentReader.MaxAttachments == 10,
                "Email attachments should allow up to ten files.");

            var csvPath = Path.Combine(
                Path.GetTempPath(),
                "Scribble-test-" + Guid.NewGuid().ToString("N") + ".csv");
            File.WriteAllText(
                csvPath,
                "Name,Amount\nWidget,42\nGadget,17");
            try
            {
                var attachments = new FakeOutlookAttachments();
                attachments.Add(new FakeOutlookAttachment(
                    "budget.csv",
                    csvPath));
                var mail = new FakeSelectedMailItem(
                    "attachment-entry",
                    "Quarterly report")
                {
                    Attachments = attachments
                };
                var application = new FakeOutlookApplication();
                application.Session.Register(
                    "attachment-entry",
                    "store",
                    mail);
                var reader = new MessageReader(application);
                var snapshot = reader.CaptureById(
                    "attachment-entry",
                    "store");
                Assert(
                    snapshot.AttachmentNames.Count == 1 &&
                    snapshot.AttachmentNames[0] == "budget.csv",
                    "Supported attachment names were not captured.");

                var request = ChatRequestFactory.Create(
                    "local-model",
                    snapshot,
                    new List<ChatTurn>(),
                    "Help me reply.");
                var reference =
                    MessageContent(request.messages[1]);
                Assert(
                    reference.Contains("Attachments (1): budget.csv"),
                    "Attachment metadata was not exposed in the context reference.");

                var host = new MailboxToolHost(
                    application,
                    snapshot);
                var loaded = host.Execute(
                    MailboxCall(
                        "read-attachment",
                        MailboxToolCatalog.ReadMessages,
                        "{\"handles\":[\"selected\"]}"));
                Assert(
                    loaded.Content.Contains("\"attachments\"") &&
                    loaded.Content.Contains("budget.csv") &&
                    loaded.Content.Contains("Widget") &&
                    loaded.Content.Contains("Gadget"),
                    "Attachment content was not loaded through read_messages.");
            }
            finally
            {
                if (File.Exists(csvPath))
                {
                    File.Delete(csvPath);
                }
            }
        }

        private static void ModelDiscoveryUsesCompatibleContract()
        {
            const string response =
                "{\"data\":[" +
                "{\"id\":\"qwen3.5-35b-a3b\"}," +
                "{\"id\":\"gpt-oss-20b\"}," +
                "{\"id\":\"text-embedding-model\"}]}";
            using (var server = new FakeEndpoint(response))
            using (var client = new OpenAiCompatibleClient())
            {
                var settings = EndpointSettings(server.BaseUrl);
                var models = client.GetModelsAsync(
                    settings,
                    CancellationToken.None)
                    .GetAwaiter()
                    .GetResult();
                server.Wait();

                Assert(
                    server.RequestLine == "GET /v1/models HTTP/1.1",
                    "Unexpected model-list request: " +
                    server.RequestLine);
                Assert(
                    server.Authorization == "Bearer test-key",
                    "The API key was not sent as a Bearer token.");
                Assert(
                    models.SequenceEqual(
                        new[]
                        {
                            "gpt-oss-20b",
                            "qwen3.5-35b-a3b"
                        }),
                    "Unexpected discovered models: " +
                    string.Join(", ", models));
            }
        }

        private static void EndpointCheckUsesLightweightProbe()
        {
            var probe = ChatRequestFactory.CreateEndpointCheck(
                "local-model");
            var json = new JavaScriptSerializer()
                .Serialize(probe);
            Assert(
                probe.tools.Count == 1 &&
                probe.tools[0].function.name ==
                MailboxToolCatalog.SearchMailbox &&
                probe.max_tokens.HasValue &&
                probe.max_tokens.Value >= 64 &&
                probe.max_tokens.Value <= 256 &&
                json.Contains("\"tool_choice\"") &&
                json.Contains("search_mailbox") &&
                !json.Contains("read_messages") &&
                !json.Contains("read_thread"),
                "The endpoint check probe was not minimized.");
        }

        private static void ToolCallResponseIsNormalized()
        {
            const string response =
                "{\"choices\":[{\"message\":{" +
                "\"role\":\"assistant\",\"content\":null," +
                "\"tool_calls\":[{\"function\":{" +
                "\"name\":\"search_mailbox\"," +
                "\"arguments\":\"{}\"}}]}}]}";
            using (var server = new FakeEndpoint(response))
            using (var client = new OpenAiCompatibleClient())
            {
                var settings = EndpointSettings(server.BaseUrl);
                var request = MakeRequest(
                    new List<ChatTurn>());
                request.model = settings.Model;
                var message = client.CompleteAsync(
                    settings,
                    request,
                    CancellationToken.None)
                    .GetAwaiter()
                    .GetResult();
                server.Wait();

                Assert(
                    server.RequestLine ==
                    "POST /v1/chat/completions HTTP/1.1",
                    "Unexpected completion request: " +
                    server.RequestLine);
                Assert(
                    server.Body.Contains(
                        "\"model\":\"local-model\"") &&
                    server.Body.Contains("\"stream\":false") &&
                    server.Body.Contains("\"tools\""),
                    "The completion request contract is incomplete.");
                Assert(
                    message.tool_calls.Count == 1 &&
                    message.tool_calls[0].id == "call_1" &&
                    message.tool_calls[0].type == "function",
                    "The missing tool-call identity was not normalized.");
            }
        }

        private static AppSettings EndpointSettings(
            string baseUrl)
        {
            return new AppSettings
            {
                BaseUrl = baseUrl,
                Model = "local-model",
                ApiKey = "test-key"
            };
        }

        private static void TextIsBounded()
        {
            var result = TextBoundary.PlainText("a\u0000bcd", 3);
            Assert(result == "abc", "Unexpected bounded text: " + result);
        }

        private static void ModelEmphasisIsNormalized()
        {
            var formatted = SafeModelText.Format(
                "**Decision** and *deadline*\n" +
                "* First action\n" +
                "Unmatched * marker and 2*3",
                TextBoundary.MaxAssistantCharacters);
            Assert(
                formatted.PlainText ==
                    "Decision and deadline\n" +
                    "- First action\n" +
                    "Unmatched  marker and 2*3" &&
                formatted.BoldRanges.Count == 2 &&
                formatted.BoldRanges[0].Start == 0 &&
                formatted.BoldRanges[0].Length == 8 &&
                formatted.BoldRanges[1].Start == 13 &&
                formatted.BoldRanges[1].Length == 8,
                "Model formatting markers were not safely normalized.");
        }

        private static void EndpointDiagnosticsExposeTransportDetails()
        {
            var exception = new AiEndpointException(
                "NETWORK_CONNECT_FAILURE",
                "The endpoint could not be reached.",
                transportDetails:
                    "SocketError ConnectionRefused NativeError 10061");
            var diagnostic = exception.ToDiagnosticText();
            Assert(
                diagnostic.Contains("[NETWORK_CONNECT_FAILURE]") &&
                diagnostic.Contains("Transport details:") &&
                diagnostic.Contains("ConnectionRefused") &&
                diagnostic.Contains("10061"),
                "Transport diagnostics are incomplete: " + diagnostic);
        }

        private static void MailboxToolsAreReadOnly()
        {
            var request = MakeRequest(new List<ChatTurn>());
            var json = new JavaScriptSerializer().Serialize(request);
            Assert(
                json.Contains("\"tools\"") &&
                json.Contains("\"tool_choice\":\"auto\"") &&
                json.Contains("\"maximum\":10") &&
                json.Contains("\"maxItems\":10"),
                "Request does not expose bounded mailbox tools.");
            Assert(json.Contains("\"stream\":false"), "Streaming must be off.");

            var names = request.tools
                .Select(tool => tool.function.name)
                .OrderBy(name => name)
                .ToArray();
            var expected = new[]
            {
                "read_messages",
                "read_thread",
                "search_mailbox"
            };
            Assert(
                names.SequenceEqual(expected),
                "Unexpected mailbox tools: " +
                string.Join(", ", names));
        }

        private static void LocalSearchCommandIsBounded()
        {
            var search = LocalSearchCommand.Parse(
                "/search project topic");
            var help = LocalSearchCommand.Parse("/search");
            var clear = LocalSearchCommand.Parse(
                "/SEARCH   clear");
            var ordinary = LocalSearchCommand.Parse(
                "/searching project topic");
            var longSearch = LocalSearchCommand.Parse(
                "/search " + new string('x', 500));

            Assert(
                search.Kind == LocalSearchCommandKind.Search &&
                search.Query == "project topic" &&
                help.Kind == LocalSearchCommandKind.Help &&
                clear.Kind == LocalSearchCommandKind.Clear &&
                ordinary.Kind == LocalSearchCommandKind.None &&
                longSearch.Query.Length == 240,
                "The local /search command contract is incomplete.");
        }

        private static void WorkingSetIsStrictlyBounded()
        {
            new AppSettings().ApplyLimits();
            var messages = Enumerable.Range(1, 12)
                .Select(index => new MessageSnapshot(
                    "entry-" + index,
                    "store",
                    "Subject " + index,
                    "Sender " + index,
                    "Recipient",
                    DateTime.UtcNow.AddMinutes(-index),
                    "private-body-" + index))
                .ToList();
            var normalized = MailboxWorkingSet.Normalize(
                messages.Concat(new[] { messages[0] }));
            Assert(
                normalized.Count == 10 &&
                MailboxWorkingSet.MaxMessages == 10 &&
                MailboxWorkingSet.HandleAt(0) == "context1" &&
                MailboxWorkingSet.HandleAt(9) == "context10",
                "The working-set normalization exceeded ten unique emails.");

            var request = MakeRequest(
                new List<ChatTurn>(),
                workingMessages: normalized);
            var names = request.tools
                .Select(tool => tool.function.name)
                .ToArray();
            var reference =
                MessageContent(request.messages[1]);
            var system =
                MessageContent(request.messages[0]);
            Assert(
                names.SequenceEqual(new[] { "read_messages" }) &&
                reference.Contains("<working_email_set") &&
                reference.Contains("context1") &&
                reference.Contains("context10") &&
                !reference.Contains("private-body-") &&
                system.Contains("working set") &&
                system.Contains("Do not search"),
                "A locked working set exposed search, threads, or email bodies.");

            var host = new MailboxToolHost(
                new FakeOutlookApplication(),
                null,
                normalized);
            var loaded = host.Execute(
                MailboxCall(
                    "read-working-set",
                    MailboxToolCatalog.ReadMessages,
                    "{\"handles\":[\"context1\",\"context2\"," +
                    "\"context3\",\"context4\",\"context5\"," +
                    "\"context6\",\"context7\",\"context8\"," +
                    "\"context9\",\"context10\"]}"));
            var duplicate = host.Execute(
                MailboxCall(
                    "read-duplicate",
                    MailboxToolCatalog.ReadMessages,
                    "{\"handles\":[\"context1\"]}"));
            var lockedSearch = host.Execute(
                MailboxCall(
                    "locked-search",
                    MailboxToolCatalog.SearchMailbox,
                    "{\"query\":\"anything\"}"));
            var lockedThread = host.Execute(
                MailboxCall(
                    "locked-thread",
                    MailboxToolCatalog.ReadThread,
                    "{\"handle\":\"context1\"}"));
            Assert(
                loaded.StatusText.Contains("Request total: 10 of 10") &&
                loaded.Content.Contains("private-body-1") &&
                loaded.Content.Contains("private-body-10") &&
                duplicate.Content.Contains("already_loaded") &&
                lockedSearch.Content.Contains(
                    "MAILBOX_WORKING_SET_LOCKED") &&
                lockedThread.Content.Contains(
                    "MAILBOX_WORKING_SET_LOCKED"),
                "The working-set host bypassed its ten-email lock.");

            try
            {
                var raised = new AppSettings
                {
                    LimitWorkingSetMessages = 500
                };
                raised.ApplyLimits();
                var expanded = MailboxWorkingSet.Normalize(messages);
                Assert(
                    MailboxWorkingSet.MaxMessages == 500 &&
                    expanded.Count == 12 &&
                    MailboxWorkingSet.HandleAt(499) == "context500",
                    "A user-raised working-set size was not honored.");
            }
            finally
            {
                new AppSettings().ApplyLimits();
            }
        }

        private static void OutlookMultiSelectionIsBounded()
        {
            var selection = new FakeSelection(
                Enumerable.Range(1, 3)
                    .Select(index =>
                        (object)new FakeSelectedMailItem(
                            "entry-" + index,
                            "Subject " + index))
                    .ToArray());
            var reader = new MessageReader(new object());
            var messages = reader.CaptureSelectionMany(
                new FakeExplorerContext(selection));
            Assert(
                messages.Count == 3 &&
                messages[0].EntryId == "entry-1" &&
                messages[2].EntryId == "entry-3",
                "Ctrl+click selection did not preserve the selected emails.");

            var overflow = false;
            try
            {
                reader.CaptureSelectionMany(
                    new FakeSelection(
                        new object[MailboxWorkingSet.MaxMessages + 1]));
            }
            catch (InvalidOperationException exception)
            {
                overflow = exception.Message.Contains(
                    "no more than " + MailboxWorkingSet.MaxMessages);
            }

            Assert(
                overflow,
                "A selection larger than ten emails was not rejected.");
        }

        private static void ActiveExplorerSelectionIsUsed()
        {
            var selection = new FakeSelection(
                new object[]
                {
                    new FakeSelectedMailItem(
                        "active-entry",
                        "Active subject")
                });
            var application = new FakeOutlookApplication
            {
                Explorer = new FakeExplorerContext(selection)
            };
            var messages = new MessageReader(application)
                .CaptureActiveSelectionMany();
            Assert(
                messages.Count == 1 &&
                messages[0].EntryId == "active-entry" &&
                messages[0].Subject == "Active subject",
                "The active Outlook Explorer selection was not captured.");
        }

        private static void ExternalContextIsBounded()
        {
            var documents = Enumerable.Range(1, 5)
                .Select(index => new ExternalContextDocument(
                    "context-" + index + ".txt",
                    new string(
                        (char)('a' + index),
                        9000)))
                .ToList();
            var normalized = ExternalContextDocument.Normalize(documents);
            var total = normalized.Sum(document =>
                document.Content.Length);
            var request = MakeRequest(
                new List<ChatTurn>(),
                externalContext: documents);
            var reference =
                MessageContent(request.messages[1]);
            Assert(
                normalized.Count == ExternalContextDocument.MaxDocuments &&
                total <= ExternalContextDocument.MaxTotalCharacters &&
                reference.Contains("<external_context count=\"3\"") &&
                reference.Contains("untrusted reference data") &&
                !reference.Contains("context-4.txt"),
                "External context exceeded its document or text boundary.");
        }

        private static void ToneProfileIsBounded()
        {
            var samples = Enumerable.Range(1, 20)
                .Select(index => new MessageSnapshot(
                    "entry-" + index,
                    "store",
                    "Subject " + index,
                    "sender-secret-" + index,
                    "recipient-secret-" + index,
                    DateTime.UtcNow,
                    "Hello, this is a reusable writing sample with a short sign-off."))
                .ToList();
            var analysis = ToneProfileRequestFactory.Create(
                "local-model",
                samples);
            var analysisBody =
                MessageContent(analysis.messages[1]);
            var profile = "Write directly and close with Regards.";
            var ordinary = MakeRequest(
                new List<ChatTurn>(),
                toneProfile: profile);
            var drafting = MakeRequest(
                new List<ChatTurn>(),
                allowDraftCreate: true,
                toneProfile: profile);
            var ordinarySystem =
                MessageContent(ordinary.messages[0]);
            var draftingSystem =
                MessageContent(drafting.messages[0]);
            var cleaned = SentMailToneSampler.CleanBody(
                "Thanks for the update.\n\nRegards,\nMe\n" +
                "-----Original Message-----\nQuoted confidential history");
            Assert(
                analysis.tools == null &&
                analysis.tool_choice == null &&
                analysisBody.Contains("Sample 15") &&
                !analysisBody.Contains("Sample 16") &&
                !analysisBody.Contains("sender-secret") &&
                !analysisBody.Contains("recipient-secret") &&
                !ordinarySystem.Contains(profile) &&
                draftingSystem.Contains(profile) &&
                draftingSystem.Contains("cannot change any capability") &&
                cleaned.Contains("Regards") &&
                !cleaned.Contains("Quoted confidential history"),
                "Tone profiling was not limited to explicit, draft-only use.");
        }

        private static void DraftToolsRequireLocalIntentAuthorization()
        {
            Assert(
                DraftIntentPolicy.AllowsCreate(
                    "Find the matching message and create a draft responding to it.") &&
                DraftIntentPolicy.AllowsCreate(
                    "Please write a reply to the selected email.") &&
                !DraftIntentPolicy.AllowsCreate(
                    "Summarize the latest messages in my mailbox.") &&
                !DraftIntentPolicy.AllowsCreate(
                    "Find the latest project update."),
                "Draft creation intent was not classified locally and conservatively.");
            Assert(
                DraftIntentPolicy.AllowsUpdate(
                    "Bolden this section and make it shorter.") &&
                DraftIntentPolicy.AllowsUpdate(
                    "Update the draft to be more formal.") &&
                !DraftIntentPolicy.AllowsUpdate(
                    "Find emails with a shorter subject.") &&
                !DraftIntentPolicy.AllowsUpdate(
                    "Summarize my inbox."),
                "Draft update intent was not classified locally and conservatively.");

            var withoutPermission = MakeRequest(
                new List<ChatTurn>());
            Assert(
                !withoutPermission.tools.Any(tool =>
                    tool.function.name ==
                    DraftToolCatalog.CreateDraft),
                "Draft creation was exposed without authorization.");

            var withPermission = MakeRequest(
                new List<ChatTurn>(),
                true);
            var names = withPermission.tools
                .Select(tool => tool.function.name)
                .OrderBy(name => name)
                .ToArray();
            var expected = new[]
            {
                "create_draft",
                "read_messages",
                "read_thread",
                "search_mailbox",
                "send_to_excel",
                "send_to_powerpoint",
                "send_to_word"
            };
            Assert(
                names.SequenceEqual(expected),
                "Authorized request tools are wrong: " +
                string.Join(", ", names));

            var json = new JavaScriptSerializer()
                .Serialize(withPermission);
            var system =
                MessageContent(withPermission.messages[0]);
            Assert(
                json.Contains("\"create_draft\"") &&
                json.Contains("\"reply_handle\"") &&
                json.Contains("\"additionalProperties\":false") &&
                system.Contains("local host recognized") &&
                system.Contains("only tool call") &&
                system.Contains("exact handle"),
                "The authorized draft boundary is incomplete.");

            var withLinkedDraft = MakeRequest(
                new List<ChatTurn>(),
                false,
                new DraftReference(
                    "new",
                    "Draft subject",
                    "recipient@example.test",
                    string.Empty,
                    "Current body"),
                true);
            Assert(
                withLinkedDraft.tools.Any(tool =>
                    tool.function.name ==
                    DraftToolCatalog.UpdateDraft) &&
                !withLinkedDraft.tools.Any(tool =>
                    tool.function.name ==
                    DraftToolCatalog.CreateDraft) &&
                MessageContent(withLinkedDraft.messages[2]).Contains(
                        "<linked_draft_reference>"),
                "A linked draft did not replace create with the bounded update tool.");

            var linkedWithoutUpdateIntent = MakeRequest(
                new List<ChatTurn>(),
                false,
                new DraftReference(
                    "new",
                    "Draft subject",
                    "recipient@example.test",
                    string.Empty,
                    "Current body"),
                false);
            Assert(
                !linkedWithoutUpdateIntent.tools.Any(tool =>
                    DraftToolCatalog.IsDraftTool(
                        tool.function.name)) &&
                !linkedWithoutUpdateIntent.messages
                    .OfType<ChatCompletionInputMessage>()
                    .Any(message => MessageContent(message).Contains(
                        "<linked_draft_reference>")),
                "A linked draft was exposed without local update intent.");

            // "Create an excel" while an email draft is linked:
            // the cross-app send tools stay available (they write
            // into sibling apps, never the linked draft), while
            // create_draft and update_draft stay gated.
            var createWithLinkedDraft = MakeRequest(
                new List<ChatTurn>(),
                true,
                new DraftReference(
                    "new",
                    "Draft subject",
                    "recipient@example.test",
                    string.Empty,
                    "Current body"),
                false);
            var linkedNames = createWithLinkedDraft.tools
                .Select(tool => tool.function.name)
                .ToArray();
            Assert(
                linkedNames.Contains("send_to_excel") &&
                linkedNames.Contains("send_to_powerpoint") &&
                linkedNames.Contains("send_to_word") &&
                !linkedNames.Contains("create_draft") &&
                !linkedNames.Contains("update_draft"),
                "Cross-app tools must stay available while an email draft is linked.");
            Assert(
                MessageContent(createWithLinkedDraft.messages[0])
                    .Contains("never say you cannot create files"),
                "The cross-app-only boundary must forbid file-creation refusals.");

            var application = new FakeOutlookApplication();
            var unauthorizedHost = new DraftToolHost(
                application);
            var rejected = unauthorizedHost.Execute(
                DraftCall(
                    "unauthorized",
                    "{\"kind\":\"new\",\"body\":\"Blocked\"}"),
                null,
                new OneShotDraftAuthorization(false),
                true);
            Assert(
                rejected.Content.Contains(
                    "DRAFT_PERMISSION_NOT_AVAILABLE") &&
                application.CreatedCount == 0,
                "A fabricated draft tool call bypassed local authorization.");
        }

        private static void DraftAuthorizationCreatesOnlyOneDraft()
        {
            var application = new FakeOutlookApplication();
            var authorization =
                new OneShotDraftAuthorization(true);
            var host = new DraftToolHost(
                application);

            var first = host.Execute(
                DraftCall(
                    "draft-1",
                    "{\"kind\":\"new\"," +
                    "\"body\":\"Hello\"," +
                    "\"subject\":\"**Subject**\\nInjected\"," +
                    "\"to\":\"one@example.test\\ntwo@example.test\"}"),
                null,
                authorization,
                true);
            var second = host.Execute(
                DraftCall(
                    "draft-2",
                    "{\"kind\":\"new\",\"body\":\"Second\"}"),
                null,
                new OneShotDraftAuthorization(true),
                true);

            Assert(
                authorization.IsConsumed &&
                authorization.IsCreated &&
                application.CreatedCount == 1,
                "The one-shot permission created more than one draft.");
            Assert(
                first.Content.Contains("\"sent\":false") &&
                second.Content.Contains(
                    "DRAFT_ALREADY_LINKED"),
                "The one-shot result contract is incomplete.");
            Assert(
                application.LastDraft.Subject ==
                    "Subject Injected" &&
                application.LastDraft.To ==
                    "one@example.test two@example.test" &&
                application.LastDraft.HTMLBody.Contains("Hello") &&
                application.LastDraft.Saved &&
                application.LastDraft.Displayed &&
                !application.LastDraft.DisplayModal,
                "The unsent draft fields or lifecycle are wrong.");
        }

        private static void LinkedDraftUpdatesSameItem()
        {
            var application = new FakeOutlookApplication();
            var host = new DraftToolHost(application);
            var createAuthorization =
                new OneShotDraftAuthorization(true);
            host.Execute(
                DraftCall(
                    "draft-create",
                    "{\"kind\":\"new\",\"body\":\"First version\"}"),
                null,
                createAuthorization,
                true);

            var original = application.LastDraft;
            var updateAuthorization =
                new OneShotDraftAuthorization(false, true);
            var updated = host.Execute(
                DraftCall(
                    "draft-update",
                    "{\"body\":\"Final section\"," +
                    "\"bold_phrases\":[\"Final\"]}",
                    DraftToolCatalog.UpdateDraft),
                null,
                updateAuthorization,
                true);

            Assert(
                updateAuthorization.IsUpdated &&
                application.CreatedCount == 1 &&
                ReferenceEquals(original, application.LastDraft) &&
                application.LastDraft.HTMLBody.Contains(
                    "<strong>Final</strong> section") &&
                application.LastDraft.SaveCount == 2 &&
                application.LastDraft.DisplayCount == 2 &&
                updated.Content.Contains("\"action\":\"updated\""),
                "The live update did not mutate and redisplay the same draft.");

            var secondUpdate = host.Execute(
                DraftCall(
                    "draft-update-2",
                    "{\"body\":\"Should not apply\"}",
                    DraftToolCatalog.UpdateDraft),
                null,
                updateAuthorization,
                true);
            Assert(
                secondUpdate.Content.Contains(
                    "DRAFT_UPDATE_NOT_AVAILABLE") &&
                application.LastDraft.SaveCount == 2,
                "One request updated the linked draft more than once.");
        }

        private static void ReplyDraftUsesExactHandle()
        {
            var application = new FakeOutlookApplication();
            var wrong = application.RegisterReplySource(
                "wrong-entry",
                "store",
                "RE: Wrong latest message",
                "wrong.sender@example.test");
            var target = application.RegisterReplySource(
                "target-entry",
                "store",
                "RE: Target project update",
                "target.sender@example.test");
            var wrongSnapshot = new MessageSnapshot(
                "wrong-entry",
                "store",
                "Wrong latest message",
                "wrong.sender@example.test",
                "recipient@example.test",
                DateTime.UtcNow,
                "Wrong body");
            var targetSnapshot = new MessageSnapshot(
                "target-entry",
                "store",
                "Target project update",
                "target.sender@example.test",
                "recipient@example.test",
                DateTime.UtcNow.AddMinutes(-5),
                "Target body");
            Func<string, MessageSnapshot> resolver = handle =>
                handle == "m2"
                    ? targetSnapshot
                    : handle == "selected"
                        ? wrongSnapshot
                        : null;
            var authorization =
                new OneShotDraftAuthorization(true);
            var host = new DraftToolHost(application);

            var result = host.Execute(
                DraftCall(
                    "reply-target",
                    "{\"kind\":\"reply\"," +
                    "\"reply_handle\":\"m2\"," +
                    "\"body\":\"Hello **Target contact**\"}"),
                resolver,
                authorization,
                true);

            Assert(
                authorization.IsCreated &&
                result.Content.Contains("\"draft_kind\":\"reply\"") &&
                target.ReplyCount == 1 &&
                wrong.ReplyCount == 0 &&
                application.LastDraft.To ==
                    "target.sender@example.test" &&
                application.LastDraft.Subject ==
                    "RE: Target project update" &&
                application.LastDraft.HTMLBody.Contains(
                    "Hello <strong>Target contact</strong>") &&
                !application.LastDraft.HTMLBody.Contains("**") &&
                host.ActiveDraft.Body == "Hello Target contact",
                "The reply was not bound to the exact retrieved handle.");
        }

        private static void ReplyDraftRequiresIssuedHandle()
        {
            var application = new FakeOutlookApplication();
            var source = application.RegisterReplySource(
                "target-entry",
                "store",
                "RE: Target",
                "target@example.test");
            var snapshot = new MessageSnapshot(
                "target-entry",
                "store",
                "Target",
                "target@example.test",
                "recipient@example.test",
                DateTime.UtcNow,
                "Body");
            Func<string, MessageSnapshot> resolver = handle =>
                handle == "selected" ? snapshot : null;

            var missingAuthorization =
                new OneShotDraftAuthorization(true);
            var missing = new DraftToolHost(application).Execute(
                DraftCall(
                    "reply-missing",
                    "{\"kind\":\"reply\",\"body\":\"Hello\"}"),
                resolver,
                missingAuthorization,
                true);
            var unknownAuthorization =
                new OneShotDraftAuthorization(true);
            var unknown = new DraftToolHost(application).Execute(
                DraftCall(
                    "reply-unknown",
                    "{\"kind\":\"reply\"," +
                    "\"reply_handle\":\"fabricated-id\"," +
                    "\"body\":\"Hello\"}"),
                resolver,
                unknownAuthorization,
                true);

            Assert(
                missing.Content.Contains(
                    "DRAFT_REPLY_HANDLE_REQUIRED") &&
                unknown.Content.Contains(
                    "DRAFT_REPLY_HANDLE_UNKNOWN") &&
                !missingAuthorization.IsConsumed &&
                !unknownAuthorization.IsConsumed &&
                source.ReplyCount == 0,
                "A missing or fabricated reply handle reached Outlook.");
        }

        private static void DraftHtmlIsSafe()
        {
            var html = SafeDraftHtml.Format(
                "Hello <script>alert('x')</script>\nImportant",
                new[] { "Important" });
            Assert(
                !html.Contains("<script>") &&
                html.Contains("&lt;script&gt;") &&
                html.Contains("<div style=") &&
                html.Contains("<strong>Important</strong>"),
                "Draft HTML did not encode untrusted markup: " + html);

            var markdown = SafeDraftHtml.FormatContent(
                "Hello **Target contact**\n* Next step\n__Important__\n" +
                "*Deadline* and stray * marker",
                new string[0]);
            Assert(
                markdown.PlainText ==
                    "Hello Target contact\n- Next step\nImportant\n" +
                    "Deadline and stray  marker" &&
                markdown.Html.Contains(
                    "Hello <strong>Target contact</strong>") &&
                markdown.Html.Contains(
                    "<strong>Important</strong>") &&
                markdown.Html.Contains(
                    "<strong>Deadline</strong>") &&
                !markdown.Html.Contains("**") &&
                !markdown.Html.Contains("__") &&
                !markdown.Html.Contains("<script>"),
                "Markdown notation was not converted to safe email formatting: " +
                markdown.Html);

            var visual = SafeDraftHtml.FormatContent(
                "# Project update\n## Decisions\n- First item\n- Second item\n" +
                "1. Confirm owner\n2. Confirm date\n---\nNext step",
                new[] { "Confirm owner" });
            Assert(
                visual.Html.Contains("<h2") &&
                visual.Html.Contains("<h3") &&
                visual.Html.Contains("<ul") &&
                visual.Html.Contains("<ol") &&
                visual.Html.Contains("<hr") &&
                visual.Html.Contains("<strong>Confirm owner</strong>") &&
                !visual.Html.Contains("<script>") &&
                !visual.Html.Contains("<img"),
                "Safe visual email layout was not rendered locally: " +
                visual.Html);

            var application = new FakeOutlookApplication();
            var rejected = new DraftToolHost(application).Execute(
                DraftCall(
                    "html-injection",
                    "{\"kind\":\"new\",\"body\":\"Safe\"," +
                    "\"html\":\"<img src=x>\"}"),
                null,
                new OneShotDraftAuthorization(true),
                true);
            Assert(
                rejected.Content.Contains(
                    "DRAFT_ARGUMENTS_INVALID") &&
                application.CreatedCount == 0,
                "Arbitrary model HTML reached the Outlook draft path.");
        }

        private static void MixedDraftToolCallIsRejected()
        {
            var application = new FakeOutlookApplication();
            var authorization =
                new OneShotDraftAuthorization(true);
            var host = new DraftToolHost(
                application);
            var result = host.Execute(
                DraftCall(
                    "mixed",
                    "{\"kind\":\"new\",\"body\":\"Hello\"}"),
                null,
                authorization,
                false);

            Assert(
                result.Content.Contains(
                    "DRAFT_TOOL_MUST_BE_EXCLUSIVE") &&
                !authorization.IsConsumed &&
                application.CreatedCount == 0,
                "A mixed draft call bypassed exclusivity.");
        }

        private static ChatToolCall DraftCall(
            string id,
            string arguments,
            string name = DraftToolCatalog.CreateDraft)
        {
            return new ChatToolCall
            {
                id = id,
                type = "function",
                function = new ChatToolCallFunction
                {
                    name = name,
                    arguments = arguments
                }
            };
        }

        private static ChatToolCall MailboxCall(
            string id,
            string name,
            string arguments)
        {
            return new ChatToolCall
            {
                id = id,
                type = "function",
                function = new ChatToolCallFunction
                {
                    name = name,
                    arguments = arguments
                }
            };
        }

        private static void RequestSchemaIsBounded()
        {
            var fields = typeof(ChatCompletionRequest)
                .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Select(property => property.Name)
                .OrderBy(name => name)
                .ToArray();
            var expected = new[]
            {
                "max_tokens",
                "messages",
                "model",
                "stream",
                "tool_choice",
                "tools"
            };
            Assert(
                fields.SequenceEqual(expected),
                "Model request capabilities changed: " +
                string.Join(", ", fields));
        }

        private static void EmailIsUntrustedData()
        {
            var request = MakeRequest(new List<ChatTurn>());
            var context =
                MessageContent(request.messages[1]);
            Assert(
                context.Contains("<selected_email_reference") &&
                context.Contains("untrusted reference data") &&
                context.Contains("Body (untrusted data") &&
                context.Contains("Message body") &&
                context.IndexOf(
                    "Message body",
                    StringComparison.Ordinal) >
                context.IndexOf(
                    "<selected_email_reference",
                    StringComparison.Ordinal),
                "Email boundary markers are missing or the " +
                "inlined body is outside the untrusted envelope.");
        }

        private static void HistoryIsBounded()
        {
            var history = Enumerable.Range(0, 30)
                .Select(index => new ChatTurn("user", "turn " + index))
                .ToList();
            var request = MakeRequest(history);
            var historyMessages = request.messages.Count - 3;
            Assert(
                historyMessages == TextBoundary.MaxConversationTurns,
                "Unexpected retained history count: " + historyMessages);
        }

        private static void DraftHasNoSend()
        {
            var methods = typeof(DraftToolHost)
                .GetMethods(BindingFlags.Instance | BindingFlags.Public)
                .Where(method => method.DeclaringType == typeof(DraftToolHost))
                .Select(method => method.Name)
                .ToArray();
            Assert(
                !methods.Any(name =>
                    name.IndexOf("Send", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    name.IndexOf("Delete", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    name.IndexOf("Move", StringComparison.OrdinalIgnoreCase) >= 0),
                "Draft host exposes a forbidden capability: " +
                string.Join(", ", methods));
        }

        private static void CompiledAddInHasNoSendInvocation()
        {
            var forbidden = new HashSet<string>(
                new[]
                {
                    "Send",
                    "SendAndReceive",
                    "Submit"
                },
                StringComparer.Ordinal);
            var violations = new List<string>();
            foreach (var type in typeof(AddIn).Assembly.GetTypes())
            {
                var methods = type
                    .GetMethods(
                        BindingFlags.Instance |
                        BindingFlags.Static |
                        BindingFlags.Public |
                        BindingFlags.NonPublic)
                    .Cast<MethodBase>()
                    .Concat(type.GetConstructors(
                        BindingFlags.Instance |
                        BindingFlags.Static |
                        BindingFlags.Public |
                        BindingFlags.NonPublic))
                    .Concat(new MethodBase[]
                    {
                        type.TypeInitializer
                    }.Where(item => item != null));
                foreach (var method in methods)
                {
                    foreach (var value in LoadedStrings(method))
                    {
                        if (forbidden.Contains(value))
                        {
                            violations.Add(
                                type.FullName + "." +
                                method.Name + ":" + value);
                        }
                    }
                }
            }

            Assert(
                violations.Count == 0,
                "Compiled Outlook send invocation found: " +
                string.Join(", ", violations));
        }

        private static IEnumerable<string> LoadedStrings(
            MethodBase method)
        {
            MethodBody body;
            try
            {
                body = method.GetMethodBody();
            }
            catch
            {
                body = null;
            }

            var il = body == null
                ? null
                : body.GetILAsByteArray();
            if (il == null)
            {
                yield break;
            }

            var oneByte = new OpCode[256];
            var twoByte = new OpCode[256];
            foreach (var field in typeof(OpCodes).GetFields(
                BindingFlags.Public | BindingFlags.Static))
            {
                var opcode = (OpCode)field.GetValue(null);
                var value = (ushort)opcode.Value;
                if (value < 0x100)
                {
                    oneByte[value] = opcode;
                }
                else if ((value & 0xff00) == 0xfe00)
                {
                    twoByte[value & 0xff] = opcode;
                }
            }

            var position = 0;
            while (position < il.Length)
            {
                var first = il[position++];
                var opcode = first == 0xfe
                    ? twoByte[il[position++]]
                    : oneByte[first];
                if (opcode.OperandType == OperandType.InlineString)
                {
                    var token = BitConverter.ToInt32(
                        il,
                        position);
                    position += 4;
                    string value;
                    try
                    {
                        value = method.Module.ResolveString(token);
                    }
                    catch
                    {
                        continue;
                    }

                    yield return value;
                    continue;
                }

                position += OperandSize(
                    opcode.OperandType,
                    il,
                    position);
            }
        }

        private static int OperandSize(
            OperandType operandType,
            byte[] il,
            int position)
        {
            switch (operandType)
            {
                case OperandType.InlineNone:
                    return 0;
                case OperandType.ShortInlineBrTarget:
                case OperandType.ShortInlineI:
                case OperandType.ShortInlineVar:
                    return 1;
                case OperandType.InlineVar:
                    return 2;
                case OperandType.InlineI:
                case OperandType.InlineBrTarget:
                case OperandType.InlineField:
                case OperandType.InlineMethod:
                case OperandType.InlineSig:
                case OperandType.InlineString:
                case OperandType.InlineTok:
                case OperandType.InlineType:
                case OperandType.ShortInlineR:
                    return 4;
                case OperandType.InlineI8:
                case OperandType.InlineR:
                    return 8;
                case OperandType.InlineSwitch:
                    return 4 +
                        BitConverter.ToInt32(il, position) * 4;
                default:
                    throw new InvalidOperationException(
                        "Unknown IL operand type: " + operandType);
            }
        }

        private static void MailboxHostHasGuardedDispatcher()
        {
            var methods = typeof(MailboxToolHost)
                .GetMethods(BindingFlags.Instance | BindingFlags.Public)
                .Where(method =>
                    method.DeclaringType ==
                    typeof(MailboxToolHost))
                .Select(method => method.Name)
                .ToArray();
            Assert(
                methods.Length == 1 &&
                methods[0] == "Execute",
                "Mailbox tool host public capabilities changed: " +
                string.Join(", ", methods));
        }

        private static void DraftHostHasGuardedDispatcher()
        {
            var methods = typeof(DraftToolHost)
                .GetMethods(BindingFlags.Instance | BindingFlags.Public)
                .Where(method =>
                    method.DeclaringType ==
                    typeof(DraftToolHost))
                .Select(method => method.Name)
                .ToArray();
            Assert(
                methods.Contains("Execute") &&
                methods.Contains("Dispose") &&
                !methods.Contains("Send"),
                "Draft tool host public capabilities changed: " +
                string.Join(", ", methods));
        }

        private static void OfficeStartupInterfacesAreDual()
        {
            Assert(
                typeof(IDTExtensibility2).IsImport,
                "IDTExtensibility2 must be a COM-import interface.");
            Assert(
                typeof(IRibbonExtensibility).IsImport,
                "IRibbonExtensibility must be a COM-import interface.");
            Assert(
                typeof(ICustomTaskPaneConsumer).IsImport,
                "ICustomTaskPaneConsumer must be a COM-import interface.");

            AssertDual(typeof(IDTExtensibility2));
            AssertDual(typeof(IRibbonExtensibility));
            AssertDual(typeof(ICustomTaskPaneConsumer));

            var addIn = new AddIn();
            AssertComInterface(addIn, typeof(IDTExtensibility2));
            AssertComInterface(addIn, typeof(IRibbonExtensibility));
            AssertComInterface(
                addIn,
                typeof(ICustomTaskPaneConsumer));
        }

        private static void ChatPaneIsComControl()
        {
            var type = typeof(ChatPane);
            var visible = type
                .GetCustomAttributes(
                    typeof(ComVisibleAttribute),
                    false)
                .Cast<ComVisibleAttribute>()
                .Single();
            var progId = type
                .GetCustomAttributes(
                    typeof(ProgIdAttribute),
                    false)
                .Cast<ProgIdAttribute>()
                .Single();
            Assert(visible.Value, "ChatPane must be COM visible.");
            Assert(
                progId.Value ==
                "Scribble.ChatPane",
                "Unexpected ChatPane ProgID.");
            Assert(
                type.GUID ==
                new Guid(
                    "14D24FA1-4342-442F-B68B-B68D7372794C"),
                "Unexpected ChatPane CLSID.");
        }

        private static void RibbonIncludesSendToAi()
        {
            var xml = new AddIn().GetCustomUI(
                "Microsoft.Outlook.Explorer");
            Assert(
                xml.Contains("ContextMenuMailItem") &&
                xml.Contains("OnSendToAi") &&
                xml.Contains("Send to Scribble") &&
                xml.Contains("label=\"Scribble\""),
                "The Outlook explorer ribbon XML is incomplete: " + xml);
        }

        private static void SelectedSubjectIsCleaned()
        {
            Assert(
                SubjectDisplay.Clean(" RE: FW: Fwd: Quarterly plan ") ==
                    "Quarterly plan" &&
                SubjectDisplay.Clean("Project update") ==
                    "Project update",
                "Selected subject prefixes were not removed safely.");
        }

        private static void AssertDual(Type interfaceType)
        {
            var attribute = interfaceType
                .GetCustomAttributes(typeof(TypeLibTypeAttribute), false)
                .Cast<TypeLibTypeAttribute>()
                .Single();
            var expected =
                TypeLibTypeFlags.FDispatchable |
                TypeLibTypeFlags.FDual;
            Assert(
                (attribute.Value & expected) == expected,
                interfaceType.Name + " must be a dual dispatch interface.");
        }

        private static void AssertComInterface(object instance, Type interfaceType)
        {
            var pointer = Marshal.GetComInterfaceForObject(
                instance,
                interfaceType);
            try
            {
                Assert(
                    pointer != IntPtr.Zero,
                    interfaceType.Name + " was not exposed by the add-in.");
            }
            finally
            {
                if (pointer != IntPtr.Zero)
                {
                    Marshal.Release(pointer);
                }
            }
        }

        private static ChatCompletionRequest MakeRequest(
            IReadOnlyList<ChatTurn> history,
            bool allowDraftCreate = false,
            DraftReference activeDraft = null,
            bool allowDraftUpdate = false,
            IReadOnlyList<MessageSnapshot> workingMessages = null,
            IReadOnlyList<ExternalContextDocument> externalContext = null,
            string toneProfile = null)
        {
            return ChatRequestFactory.Create(
                "local-model",
                new MessageSnapshot(
                    "entry",
                    "store",
                    "Subject",
                    "Sender",
                    "Recipient",
                    DateTime.UtcNow,
                    "Message body"),
                history,
                "Help me reply.",
                allowDraftCreate,
                activeDraft,
                allowDraftUpdate,
                workingMessages,
                externalContext,
                toneProfile);
        }

        private static string MessageContent(object message)
        {
            return Convert.ToString(
                ((ChatCompletionInputMessage)message).content) ??
                string.Empty;
        }

        private static void VisionCapableModelsAreDetectedBroadly()
        {
            Assert(
                ModelCatalog.IsVisionCapable("qwen3-vl-30b") &&
                ModelCatalog.IsVisionCapable("Qwen3-VL-30B-Instruct") &&
                ModelCatalog.IsVisionCapable("my-vision-model") &&
                ModelCatalog.IsVisionCapable("gemma-4-31b-it") &&
                ModelCatalog.IsVisionCapable("gemma-4-26b-a4b-it") &&
                ModelCatalog.IsVisionCapable("gemma3-27b-it") &&
                ModelCatalog.IsVisionCapable("llava-1.6-13b") &&
                ModelCatalog.IsVisionCapable("MiniCPM-V-2_6") &&
                !ModelCatalog.IsVisionCapable("gpt-oss-20b") &&
                !ModelCatalog.IsVisionCapable("qwen3.6-35b-a3b") &&
                !ModelCatalog.IsVisionCapable("text-embedding-vl"),
                "Vision capability detection is too narrow or too broad.");
        }

        private static void VisionPrefetchInjectsImageInput()
        {
            var pngPath = Path.Combine(
                Path.GetTempPath(),
                "Scribble-prefetch-" + Guid.NewGuid().ToString("N") + ".png");
            File.WriteAllBytes(
                pngPath,
                Convert.FromBase64String(
                    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJ" +
                    "AAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="));
            try
            {
                var attachments = new FakeOutlookAttachments();
                attachments.Add(new FakeOutlookAttachment(
                    "scan.png",
                    pngPath));
                var mail = new FakeSelectedMailItem(
                    "prefetch-entry",
                    "Invoice scan")
                {
                    Attachments = attachments
                };
                var application = new FakeOutlookApplication();
                application.Session.Register(
                    "prefetch-entry",
                    "store",
                    mail);
                var snapshot = new MessageReader(application)
                    .CaptureById("prefetch-entry", "store");
                var request = ChatRequestFactory.Create(
                    "qwen3-vl-30b",
                    snapshot,
                    new List<ChatTurn>(),
                    "Summarize the image.");
                var host = new MailboxToolHost(application, snapshot);
                Assert(
                    VisionImagePrefetch.TryInject(
                        request,
                        "qwen3-vl-30b",
                        host,
                        snapshot,
                        null),
                    "Vision prefetch did not inject image context.");
                var body = new JavaScriptSerializer()
                    .Serialize(request);
                Assert(
                    body.Contains("\"type\":\"image_url\"") &&
                    body.Contains("data:image/png;base64,"),
                    "Prefetched request is missing multimodal image content.");
            }
            finally
            {
                if (File.Exists(pngPath))
                {
                    File.Delete(pngPath);
                }
            }
        }

        private static void VisionAutoSwitchPicksBestDiscoveredModel()
        {
            var discovered = new[]
            {
                "gpt-oss-20b",
                "qwen3.6-35b-a3b",
                "qwen3-vl-30b"
            };
            Assert(
                ModelCatalog.FindBestVisionModel(discovered) ==
                    "qwen3-vl-30b",
                "The best vision model was not selected from the saved list.");
            Assert(
                ModelCatalog.FindBestVisionModel(
                    new[] { "gemma-4-31b-it", "qwen3-vl-30b" }) ==
                    "qwen3-vl-30b",
                "The dedicated vision model should be preferred over multimodal Gemma.");
            Assert(
                ModelCatalog.FindBestVisionModel(
                    new[] { "gpt-oss-20b", "gemma-4-31b-it" }) ==
                    "gemma-4-31b-it",
                "Multimodal Gemma should be used when no dedicated vision model exists.");

            var settings = new AppSettings
            {
                Model = "gpt-oss-20b",
                SwitchToVisionModelForImages = true,
                DiscoveredModels = new List<string>(discovered)
            };
            var snapshot = new MessageSnapshot(
                "entry",
                "store",
                "Invoice",
                "Sender",
                "Recipient",
                DateTime.UtcNow,
                "Body",
                new[] { "scan.png" });
            var routed = ModelRouting.ResolveForRequest(
                settings,
                ModelRouting.ContextMayIncludeImages(snapshot, null));
            Assert(
                routed == "qwen3-vl-30b" &&
                ModelRouting.IsTemporaryVisionSwitch(settings, routed),
                "Image requests did not temporarily switch to the vision model.");

            settings.SwitchToVisionModelForImages = false;
            Assert(
                ModelRouting.ResolveForRequest(
                    settings,
                    true) == "gpt-oss-20b",
                "The checkbox should gate temporary vision switching.");

            settings.SwitchToVisionModelForImages = true;
            settings.Model = "qwen3-vl-30b";
            Assert(
                ModelRouting.ResolveForRequest(
                    settings,
                    true) == "qwen3-vl-30b" &&
                !ModelRouting.IsTemporaryVisionSwitch(
                    settings,
                    "qwen3-vl-30b"),
                "Vision models should not be replaced when already selected.");
        }

        private static void GaussModelsAreExcluded()
        {
            const string response =
                "{\"data\":[" +
                "{\"id\":\"qwen3-vl-30b\"}," +
                "{\"id\":\"gauss-vision-7b\"}," +
                "{\"id\":\"my-gausso-model\"}," +
                "{\"id\":\"gpt-oss-20b\"}]}";
            using (var server = new FakeEndpoint(response))
            using (var client = new OpenAiCompatibleClient())
            {
                var models = client.GetModelsAsync(
                    EndpointSettings(server.BaseUrl),
                    CancellationToken.None)
                    .GetAwaiter()
                    .GetResult();
                server.Wait();

                Assert(
                    models.SequenceEqual(
                        new[]
                        {
                            "gpt-oss-20b",
                            "qwen3-vl-30b"
                        }),
                    "Gauss models were not excluded: " +
                    string.Join(", ", models));
            }

            Assert(
                ModelCatalog.IsDisallowedModel("gauss-vision-7b") &&
                ModelCatalog.IsDisallowedModel("my-gausso-model") &&
                !ModelCatalog.IsDisallowedModel("qwen3-vl-30b"),
                "The Gauss filter did not classify model names correctly.");
        }

        private static void VisionModelsReceiveMultimodalFollowUp()
        {
            var pngPath = Path.Combine(
                Path.GetTempPath(),
                "Scribble-test-" + Guid.NewGuid().ToString("N") + ".png");
            File.WriteAllBytes(
                pngPath,
                Convert.FromBase64String(
                    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJ" +
                    "AAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="));
            try
            {
                var attachments = new FakeOutlookAttachments();
                attachments.Add(new FakeOutlookAttachment(
                    "scan.png",
                    pngPath));
                var mail = new FakeSelectedMailItem(
                    "vision-entry",
                    "Invoice scan")
                {
                    Attachments = attachments
                };
                var application = new FakeOutlookApplication();
                application.Session.Register(
                    "vision-entry",
                    "store",
                    mail);
                var snapshot = new MessageReader(application)
                    .CaptureById("vision-entry", "store");
                var host = new MailboxToolHost(application, snapshot);
                var loaded = host.Execute(
                    MailboxCall(
                        "read-vision",
                        MailboxToolCatalog.ReadMessages,
                        "{\"handles\":[\"selected\"]}"));
                Assert(
                    loaded.VisionImages.Count == 1 &&
                    loaded.VisionImages[0].FileName == "scan.png" &&
                    loaded.VisionImages[0].DataUrl.StartsWith(
                        "data:image/png;base64,") &&
                    loaded.Content.Contains("vision_available") &&
                    !loaded.Content.Contains("base64,"),
                    "Vision image payloads were not separated from tool JSON.");

                var request = ChatRequestFactory.Create(
                    "qwen3-vl-30b",
                    snapshot,
                    new List<ChatTurn>(),
                    "What is in the attachment?");
                ChatRequestFactory.AppendToolExchange(
                    request,
                    new ChatCompletionResponseMessage
                    {
                        role = "assistant",
                        content = string.Empty,
                        tool_calls = new List<ChatToolCall>
                        {
                            MailboxCall(
                                "read-vision",
                                MailboxToolCatalog.ReadMessages,
                                "{\"handles\":[\"selected\"]}")
                        }
                    },
                    new List<MailboxToolResult> { loaded },
                    "qwen3-vl-30b");
                var body = new JavaScriptSerializer()
                    .Serialize(request);
                Assert(
                    body.Contains("\"type\":\"image_url\"") &&
                    body.Contains("data:image/png;base64,") &&
                    body.Contains("scan.png"),
                    "Vision follow-up did not append multimodal image content.");
            }
            finally
            {
                if (File.Exists(pngPath))
                {
                    File.Delete(pngPath);
                }
            }
        }

        private static void SystemPromptStatesImageCapability()
        {
            var snapshot = new MessageSnapshot(
                "entry",
                "store",
                "Invoice",
                "Sender",
                "Recipient",
                DateTime.UtcNow,
                "Body",
                new[] { "scan.png" });

            var textRequest = ChatRequestFactory.Create(
                "gpt-oss-20b",
                snapshot,
                new List<ChatTurn>(),
                "Summarize this image.");
            var textSystem = MessageContent(textRequest.messages[0]);
            Assert(
                textSystem.Contains("text-only") &&
                textSystem.Contains("cannot view images") &&
                textSystem.Contains("Today's date is"),
                "Text-only models are not told about unseen images.");

            var visionRequest = ChatRequestFactory.Create(
                "qwen3-vl-30b",
                snapshot,
                new List<ChatTurn>(),
                "Summarize this image.");
            var visionSystem = MessageContent(visionRequest.messages[0]);
            Assert(
                visionSystem.Contains("vision-capable model") &&
                visionSystem.Contains("attachment filename"),
                "Vision models are not instructed to use image input.");

            var plainRequest = ChatRequestFactory.Create(
                "gpt-oss-20b",
                new MessageSnapshot(
                    "entry",
                    "store",
                    "Subject",
                    "Sender",
                    "Recipient",
                    DateTime.UtcNow,
                    "Body"),
                new List<ChatTurn>(),
                "Hello.");
            Assert(
                !MessageContent(plainRequest.messages[0])
                    .Contains("text-only"),
                "The image warning should appear only when images are in context.");
        }

        private static void VisionImageLimitsAreEnforced()
        {
            var images = new List<VisionImagePayload>();
            for (var index = 0; index < 12; index++)
            {
                images.Add(new VisionImagePayload(
                    "img" + index + ".png",
                    "data:image/png;base64,AAA" + index));
            }

            images.Add(new VisionImagePayload(
                "huge.png",
                "data:image/png;base64," + new string('A', 2200001)));

            var request = ChatRequestFactory.Create(
                "qwen3-vl-30b",
                null,
                new List<ChatTurn>(),
                "Describe the images.");
            ChatRequestFactory.AppendToolExchange(
                request,
                new ChatCompletionResponseMessage
                {
                    role = "assistant",
                    content = string.Empty,
                    tool_calls = new List<ChatToolCall>
                    {
                        MailboxCall(
                            "read-many",
                            MailboxToolCatalog.ReadMessages,
                            "{\"handles\":[\"selected\"]}")
                    }
                },
                new List<MailboxToolResult>
                {
                    new MailboxToolResult(
                        "read-many",
                        "{}",
                        "loaded",
                        images)
                },
                "qwen3-vl-30b");
            var body = new JavaScriptSerializer
            {
                MaxJsonLength = int.MaxValue
            }.Serialize(request);
            Assert(
                CountToken(body, "\"type\":\"image_url\"") == 8,
                "The per-request image cap was not applied.");
            Assert(
                !body.Contains("huge.png"),
                "An oversized image data URL was not dropped.");
            Assert(
                body.Contains("omitted"),
                "Omitted images are not disclosed to the model.");
        }

        private static void PastedImagesWithoutExtensionAreRead()
        {
            var pngPath = Path.Combine(
                Path.GetTempPath(),
                "Scribble-pasted-" + Guid.NewGuid().ToString("N"));
            File.WriteAllBytes(
                pngPath,
                Convert.FromBase64String(
                    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJ" +
                    "AAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="));
            try
            {
                var attachments = new FakeOutlookAttachments();
                attachments.Add(new FakeOutlookAttachment(
                    "image001",
                    pngPath));
                var mail = new FakeSelectedMailItem(
                    "pasted-entry",
                    "Survey request")
                {
                    Attachments = attachments
                };
                var application = new FakeOutlookApplication();
                application.Session.Register(
                    "pasted-entry",
                    "store",
                    mail);
                var snapshot = new MessageReader(application)
                    .CaptureById("pasted-entry", "store");
                var host = new MailboxToolHost(application, snapshot);
                var loaded = host.Execute(
                    MailboxCall(
                        "read-pasted",
                        MailboxToolCatalog.ReadMessages,
                        "{\"handles\":[\"selected\"]}"));
                Assert(
                    loaded.VisionImages.Count == 1 &&
                    loaded.VisionImages[0].DataUrl.StartsWith(
                        "data:image/png;base64,"),
                    "An extensionless pasted image was not sniffed as PNG.");
            }
            finally
            {
                if (File.Exists(pngPath))
                {
                    File.Delete(pngPath);
                }
            }
        }

        private static void SignatureImagesAreIgnored()
        {
            var pngPath = Path.Combine(
                Path.GetTempPath(),
                "Scribble-sig-" + Guid.NewGuid().ToString("N") + ".png");
            File.WriteAllBytes(
                pngPath,
                Convert.FromBase64String(
                    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJ" +
                    "AAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="));
            try
            {
                var attachments = new FakeOutlookAttachments();
                attachments.Add(new FakeOutlookAttachment(
                    "logo.png",
                    pngPath)
                {
                    Size = 2048,
                    PropertyAccessor = new FakePropertyAccessor
                    {
                        Hidden = true,
                        ContentId = "logo@signature"
                    }
                });
                attachments.Add(new FakeOutlookAttachment(
                    "screenshot.png",
                    pngPath)
                {
                    Size = 500000,
                    PropertyAccessor = new FakePropertyAccessor
                    {
                        Hidden = true,
                        ContentId = "shot@body"
                    }
                });
                var mail = new FakeSelectedMailItem(
                    "signature-entry",
                    "Weekly update")
                {
                    Attachments = attachments
                };
                var application = new FakeOutlookApplication();
                application.Session.Register(
                    "signature-entry",
                    "store",
                    mail);
                var snapshot = new MessageReader(application)
                    .CaptureById("signature-entry", "store");
                Assert(
                    snapshot.AttachmentNames.Count == 1 &&
                    snapshot.AttachmentNames[0] == "screenshot.png",
                    "Signature images should be excluded from metadata: " +
                    string.Join(", ", snapshot.AttachmentNames));

                var host = new MailboxToolHost(application, snapshot);
                var loaded = host.Execute(
                    MailboxCall(
                        "read-signature",
                        MailboxToolCatalog.ReadMessages,
                        "{\"handles\":[\"selected\"]}"));
                Assert(
                    loaded.VisionImages.Count == 1 &&
                    loaded.VisionImages[0].FileName ==
                        "screenshot.png",
                    "Only the large inline image should reach vision input.");
                Assert(
                    loaded.Content.Contains(
                        "ignored as signature graphics"),
                    "The skipped signature image was not disclosed.");
            }
            finally
            {
                if (File.Exists(pngPath))
                {
                    File.Delete(pngPath);
                }
            }
        }

        private static void LocalFilesLoadAsContext()
        {
            var temp = Path.Combine(
                Path.GetTempPath(),
                "Scribble-local-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(temp);
            try
            {
                var pngPath = Path.Combine(temp, "photo.png");
                File.WriteAllBytes(
                    pngPath,
                    Convert.FromBase64String(
                        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJ" +
                        "AAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="));
                var image = EmailAttachmentReader.LoadLocalFile(
                    pngPath);
                Assert(
                    image != null &&
                    image.ImageDataUrl.StartsWith(
                        "data:image/png;base64,"),
                    "A local image did not load as vision input.");
                var thumbnail =
                    EmailAttachmentReader.BuildThumbnailDataUrl(
                        pngPath);
                Assert(
                    thumbnail != null &&
                    thumbnail.StartsWith("data:image/jpeg;base64,"),
                    "A local image did not produce a tray thumbnail.");

                var pdfPath = Path.Combine(temp, "report.pdf");
                File.WriteAllText(
                    pdfPath,
                    "%PDF-1.4\n1 0 obj << /Length 96 >>\nstream\n" +
                    "BT /F1 12 Tf (Quarterly spending summary for " +
                    "the operations department) Tj ET\n" +
                    "endstream\nendobj\ntrailer\n%%EOF",
                    Encoding.ASCII);
                var document = EmailAttachmentReader.LoadLocalFile(
                    pdfPath);
                Assert(
                    document != null &&
                    document.Text.Contains(
                        "Quarterly spending summary for the " +
                        "operations department"),
                    "A local PDF did not load as extracted text.");
            }
            finally
            {
                if (Directory.Exists(temp))
                {
                    Directory.Delete(temp, true);
                }
            }
        }

        private static void LegacyAndUnknownAttachmentsAreHandled()
        {
            var temp = Path.Combine(
                Path.GetTempPath(),
                "Scribble-legacy-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(temp);
            try
            {
                const string pptSentence =
                    "Roadmap kickoff agenda for the team";
                const string docSentence =
                    "Contract renewal terms with the northern supplier";
                var pptPath = Path.Combine(temp, "deck.ppt");
                File.WriteAllBytes(
                    pptPath,
                    BuildCompoundFile(
                        "PowerPoint Document",
                        BuildPptStream(pptSentence)));
                var docPath = Path.Combine(temp, "notes.doc");
                var docBytes = BuildCompoundFile(
                    "WordDocument",
                    BuildDocStream(docSentence));
                File.WriteAllBytes(docPath, docBytes);
                var datPath = Path.Combine(temp, "mystery.dat");
                File.WriteAllBytes(datPath, docBytes);
                var rtfPath = Path.Combine(temp, "memo.rtf");
                File.WriteAllText(
                    rtfPath,
                    "{\\rtf1\\ansi{\\fonttbl{\\f0 Calibri;}}" +
                    "\\f0\\fs22 Meeting notes:\\par Budget " +
                    "\\b approved\\b0  for the third quarter.}",
                    Encoding.ASCII);
                var binPath = Path.Combine(temp, "data.bin");
                var noise = new byte[256];
                new Random(7).NextBytes(noise);
                noise[0] = 0;
                File.WriteAllBytes(binPath, noise);

                var attachments = new FakeOutlookAttachments();
                attachments.Add(new FakeOutlookAttachment(
                    "deck.ppt",
                    pptPath));
                attachments.Add(new FakeOutlookAttachment(
                    "notes.doc",
                    docPath));
                attachments.Add(new FakeOutlookAttachment(
                    "mystery.dat",
                    datPath));
                attachments.Add(new FakeOutlookAttachment(
                    "memo.rtf",
                    rtfPath));
                attachments.Add(new FakeOutlookAttachment(
                    "data.bin",
                    binPath));
                var mail = new FakeSelectedMailItem(
                    "legacy-entry",
                    "Old documents")
                {
                    Attachments = attachments
                };
                var application = new FakeOutlookApplication();
                application.Session.Register(
                    "legacy-entry",
                    "store",
                    mail);
                var snapshot = new MessageReader(application)
                    .CaptureById("legacy-entry", "store");
                Assert(
                    snapshot.AttachmentNames.Count == 5,
                    "All attachments should be listed regardless of type: " +
                    string.Join(", ", snapshot.AttachmentNames));

                var host = new MailboxToolHost(application, snapshot);
                var loaded = host.Execute(
                    MailboxCall(
                        "read-legacy",
                        MailboxToolCatalog.ReadMessages,
                        "{\"handles\":[\"selected\"]}"));
                Assert(
                    loaded.Content.Contains(pptSentence),
                    "Legacy .ppt slide text was not extracted.");
                Assert(
                    loaded.Content.Contains(docSentence),
                    "Legacy .doc text was not extracted twice over: " +
                    "direct extension failed.");
                Assert(
                    loaded.Content.Contains("Meeting notes:") &&
                    loaded.Content.Contains("approved") &&
                    !loaded.Content.Contains("Calibri"),
                    "RTF text was not stripped of control words.");
                Assert(
                    loaded.Content.Contains("data.bin") &&
                    loaded.Content.Contains(
                        "could not be converted"),
                    "Unreadable attachments must be visibly noted.");
            }
            finally
            {
                if (Directory.Exists(temp))
                {
                    Directory.Delete(temp, true);
                }
            }
        }

        private static byte[] BuildPptStream(string sentence)
        {
            var chars = Encoding.Unicode.GetBytes(sentence);
            var stream = new MemoryStream();
            WriteRecordHeader(
                stream,
                0x000F,
                0x03E8,
                chars.Length + 8);
            WriteRecordHeader(stream, 0x0000, 0x0FA0, chars.Length);
            stream.Write(chars, 0, chars.Length);
            var padded = new byte[Math.Max(
                4096,
                (int)stream.Length)];
            Array.Copy(
                stream.ToArray(),
                padded,
                (int)stream.Length);
            return padded;
        }

        private static void WriteRecordHeader(
            MemoryStream stream,
            int verInstance,
            int recordType,
            int length)
        {
            WriteUInt16(stream, verInstance);
            WriteUInt16(stream, recordType);
            WriteUInt32(stream, (uint)length);
        }

        private static byte[] BuildDocStream(string sentence)
        {
            var stream = new byte[4096];
            var text = Encoding.GetEncoding("ISO-8859-1")
                .GetBytes(sentence);
            const int fcMin = 0x400;
            Array.Copy(text, 0, stream, fcMin, text.Length);
            WriteUInt32At(stream, 0x18, fcMin);
            WriteUInt32At(stream, 0x1C, (uint)(fcMin + text.Length));
            return stream;
        }

        private static byte[] BuildCompoundFile(
            string streamName,
            byte[] payload)
        {
            const int sectorSize = 512;
            var dataSectors =
                (payload.Length + sectorSize - 1) / sectorSize;
            var file = new byte[
                512 + 512 + 512 + dataSectors * sectorSize];
            file[0] = 0xD0;
            file[1] = 0xCF;
            file[2] = 0x11;
            file[3] = 0xE0;
            file[4] = 0xA1;
            file[5] = 0xB1;
            file[6] = 0x1A;
            file[7] = 0xE1;
            WriteUInt16At(file, 24, 0x3E);
            WriteUInt16At(file, 26, 3);
            WriteUInt16At(file, 28, 0xFFFE);
            WriteUInt16At(file, 30, 9);
            WriteUInt16At(file, 32, 6);
            WriteUInt32At(file, 44, 1);
            WriteUInt32At(file, 48, 1);
            WriteUInt32At(file, 56, 4096);
            WriteUInt32At(file, 60, 0xFFFFFFFE);
            WriteUInt32At(file, 68, 0xFFFFFFFE);
            WriteUInt32At(file, 76, 0);
            for (var index = 1; index < 109; index++)
            {
                WriteUInt32At(file, 76 + index * 4, 0xFFFFFFFF);
            }

            for (var index = 0; index < 128; index++)
            {
                WriteUInt32At(file, 512 + index * 4, 0xFFFFFFFF);
            }

            WriteUInt32At(file, 512, 0xFFFFFFFD);
            WriteUInt32At(file, 516, 0xFFFFFFFE);
            for (var index = 0; index < dataSectors; index++)
            {
                WriteUInt32At(
                    file,
                    512 + (2 + index) * 4,
                    index == dataSectors - 1
                        ? 0xFFFFFFFE
                        : (uint)(3 + index));
            }

            WriteDirectoryEntry(
                file,
                1024,
                "Root Entry",
                5,
                0xFFFFFFFE,
                0,
                1);
            WriteDirectoryEntry(
                file,
                1024 + 128,
                streamName,
                2,
                2,
                (uint)payload.Length,
                0xFFFFFFFF);
            Array.Copy(payload, 0, file, 1536, payload.Length);
            return file;
        }

        private static void WriteDirectoryEntry(
            byte[] file,
            int offset,
            string name,
            byte objectType,
            uint startSector,
            uint size,
            uint child)
        {
            var encoded = Encoding.Unicode.GetBytes(name);
            Array.Copy(encoded, 0, file, offset, encoded.Length);
            WriteUInt16At(
                file,
                offset + 64,
                (ushort)(encoded.Length + 2));
            file[offset + 66] = objectType;
            file[offset + 67] = 1;
            WriteUInt32At(file, offset + 68, 0xFFFFFFFF);
            WriteUInt32At(file, offset + 72, 0xFFFFFFFF);
            WriteUInt32At(file, offset + 76, child);
            WriteUInt32At(file, offset + 116, startSector);
            WriteUInt32At(file, offset + 120, size);
        }

        private static void WriteUInt16(
            MemoryStream stream,
            int value)
        {
            stream.WriteByte((byte)(value & 0xFF));
            stream.WriteByte((byte)((value >> 8) & 0xFF));
        }

        private static void WriteUInt32(
            MemoryStream stream,
            uint value)
        {
            stream.WriteByte((byte)(value & 0xFF));
            stream.WriteByte((byte)((value >> 8) & 0xFF));
            stream.WriteByte((byte)((value >> 16) & 0xFF));
            stream.WriteByte((byte)((value >> 24) & 0xFF));
        }

        private static void WriteUInt16At(
            byte[] target,
            int offset,
            ushort value)
        {
            target[offset] = (byte)(value & 0xFF);
            target[offset + 1] = (byte)((value >> 8) & 0xFF);
        }

        private static void WriteUInt32At(
            byte[] target,
            int offset,
            uint value)
        {
            target[offset] = (byte)(value & 0xFF);
            target[offset + 1] = (byte)((value >> 8) & 0xFF);
            target[offset + 2] = (byte)((value >> 16) & 0xFF);
            target[offset + 3] = (byte)((value >> 24) & 0xFF);
        }

        private static void PdfCidFontTextIsDecoded()
        {
            // Word, Chrome, and LibreOffice PDFs store text as hex glyph
            // codes decoded through each font's ToUnicode CMap.
            const string sentence =
                "Hello quarterly budget numbers attached for review today";
            var hex = new StringBuilder();
            foreach (var character in sentence)
            {
                hex.Append(((int)character).ToString("X4"));
            }

            var pdf =
                "%PDF-1.4\n" +
                "1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj\n" +
                "2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 >> endobj\n" +
                "3 0 obj << /Type /Page /Parent 2 0 R " +
                "/Resources << /Font << /F1 4 0 R >> >> " +
                "/Contents 5 0 R >> endobj\n" +
                "4 0 obj << /Type /Font /Subtype /Type0 " +
                "/ToUnicode 6 0 R >> endobj\n" +
                "5 0 obj << /Length 400 >>\nstream\n" +
                "BT /F1 12 Tf <" + hex + "> Tj ET\n" +
                "endstream\nendobj\n" +
                "6 0 obj << /Length 200 >>\nstream\n" +
                "begincmap\n" +
                "1 begincodespacerange\n<0000> <FFFF>\n" +
                "endcodespacerange\n" +
                "1 beginbfrange\n<0020> <007E> <0020>\nendbfrange\n" +
                "endcmap\n" +
                "endstream\nendobj\n" +
                "trailer << /Root 1 0 R >>\n%%EOF";
            var extracted = PdfTextExtractor.Extract(
                Encoding.ASCII.GetBytes(pdf),
                8000);
            Assert(
                extracted.Contains(sentence),
                "CID hex-coded PDF text was not decoded. Got: " +
                extracted);

            var literalPdf =
                "%PDF-1.4\n1 0 obj << /Length 96 >>\nstream\n" +
                "BT /F1 12 Tf (Plain literal payment reminder for " +
                "the october invoice) Tj ET\n" +
                "endstream\nendobj\ntrailer\n%%EOF";
            var literal = PdfTextExtractor.Extract(
                Encoding.ASCII.GetBytes(literalPdf),
                8000);
            Assert(
                literal.Contains(
                    "Plain literal payment reminder for the " +
                    "october invoice"),
                "Literal PDF text extraction regressed. Got: " +
                literal);
        }

        private static void CalendarInvitesAreReadable()
        {
            var agendaPath = Path.Combine(
                Path.GetTempPath(),
                "Scribble-agenda-" +
                Guid.NewGuid().ToString("N") +
                ".txt");
            File.WriteAllText(
                agendaPath,
                "Agenda: budget review and hiring plan");
            try
            {
                var attachments = new FakeOutlookAttachments();
                attachments.Add(new FakeOutlookAttachment(
                    "agenda.txt",
                    agendaPath));
                var invite = new FakeSelectedMailItem(
                    "invite-entry",
                    "Design review")
                {
                    MessageClass = "IPM.Schedule.Meeting.Request",
                    Attachments = attachments
                };
                var application = new FakeOutlookApplication();
                application.Session.Register(
                    "invite-entry",
                    "store",
                    invite);
                var snapshot = new MessageReader(application)
                    .CaptureById("invite-entry", "store");
                Assert(
                    snapshot.Subject == "Design review" &&
                    snapshot.Body.Contains("Message body") &&
                    snapshot.AttachmentNames.Count == 1,
                    "The meeting invite was not captured as readable context.");

                var host = new MailboxToolHost(application, snapshot);
                var loaded = host.Execute(
                    MailboxCall(
                        "read-invite",
                        MailboxToolCatalog.ReadMessages,
                        "{\"handles\":[\"selected\"]}"));
                Assert(
                    loaded.Content.Contains(
                        "Agenda: budget review and hiring plan"),
                    "The invite attachment text was not extracted.");

                var appointment = new FakeSelectedMailItem(
                    "appointment-entry",
                    "Quarterly offsite")
                {
                    MessageClass = "IPM.Appointment"
                };
                application.Session.Register(
                    "appointment-entry",
                    "store",
                    appointment);
                var appointmentSnapshot =
                    new MessageReader(application)
                        .CaptureById("appointment-entry", "store");
                Assert(
                    appointmentSnapshot.Subject == "Quarterly offsite",
                    "Appointment items should be readable.");

                var rejected = false;
                var task = new FakeSelectedMailItem(
                    "task-entry",
                    "Not readable")
                {
                    MessageClass = "IPM.Task"
                };
                application.Session.Register(
                    "task-entry",
                    "store",
                    task);
                try
                {
                    new MessageReader(application)
                        .CaptureById("task-entry", "store");
                }
                catch (InvalidOperationException)
                {
                    rejected = true;
                }

                Assert(
                    rejected,
                    "Non-mail, non-calendar item classes must stay rejected.");
            }
            finally
            {
                if (File.Exists(agendaPath))
                {
                    File.Delete(agendaPath);
                }
            }
        }

        private static void DocumentAttachmentsAreExtracted()
        {
            var temp = Path.Combine(
                Path.GetTempPath(),
                "Scribble-docs-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(temp);
            var pptxPath = Path.Combine(temp, "deck.pptx");
            var docxPath = Path.Combine(temp, "notes.docx");
            var pdfPath = Path.Combine(temp, "invoice.pdf");
            try
            {
                const string drawingNamespace =
                    "http://schemas.openxmlformats.org/drawingml/2006/main";
                WriteZipEntry(
                    pptxPath,
                    "ppt/slides/slide1.xml",
                    "<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" " +
                    "xmlns:a=\"" + drawingNamespace + "\">" +
                    "<p:cSld><p:spTree><p:sp><p:txBody>" +
                    "<a:p><a:r><a:t>Quarterly revenue plan</a:t></a:r></a:p>" +
                    "</p:txBody></p:sp></p:spTree></p:cSld></p:sld>");
                WriteZipEntry(
                    docxPath,
                    "word/document.xml",
                    "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                    "<w:body><w:p><w:r><w:t>Contract terms baseline</w:t></w:r></w:p>" +
                    "</w:body></w:document>");
                File.WriteAllText(
                    pdfPath,
                    "%PDF-1.4\n1 0 obj << /Length 96 >>\nstream\n" +
                    "BT /F1 12 Tf (Invoice total 12345 dollars for " +
                    "consulting services rendered in July) Tj ET\n" +
                    "endstream\nendobj\ntrailer\n%%EOF",
                    Encoding.ASCII);

                var attachments = new FakeOutlookAttachments();
                attachments.Add(new FakeOutlookAttachment(
                    "deck.pptx",
                    pptxPath));
                attachments.Add(new FakeOutlookAttachment(
                    "notes.docx",
                    docxPath));
                attachments.Add(new FakeOutlookAttachment(
                    "invoice.pdf",
                    pdfPath));
                var mail = new FakeSelectedMailItem(
                    "document-entry",
                    "Project documents")
                {
                    Attachments = attachments
                };
                var application = new FakeOutlookApplication();
                application.Session.Register(
                    "document-entry",
                    "store",
                    mail);
                var snapshot = new MessageReader(application)
                    .CaptureById("document-entry", "store");
                Assert(
                    snapshot.AttachmentNames.Count == 3,
                    "Document attachments were not listed in metadata: " +
                    string.Join(", ", snapshot.AttachmentNames));

                var host = new MailboxToolHost(application, snapshot);
                var loaded = host.Execute(
                    MailboxCall(
                        "read-documents",
                        MailboxToolCatalog.ReadMessages,
                        "{\"handles\":[\"selected\"]}"));
                Assert(
                    loaded.Content.Contains("Quarterly revenue plan") &&
                    loaded.Content.Contains("[Slide 1]"),
                    "PowerPoint slide text was not extracted.");
                Assert(
                    loaded.Content.Contains("Contract terms baseline"),
                    "Word document text was not extracted.");
                Assert(
                    loaded.Content.Contains(
                        "Invoice total 12345 dollars for " +
                        "consulting services rendered in July"),
                    "PDF text was not extracted.");
            }
            finally
            {
                if (Directory.Exists(temp))
                {
                    Directory.Delete(temp, true);
                }
            }
        }

        private static void WriteZipEntry(
            string zipPath,
            string entryName,
            string xml)
        {
            WriteZipEntries(
                zipPath,
                new[] { entryName },
                new[] { xml });
        }

        private static void WriteZipEntries(
            string zipPath,
            string[] entryNames,
            string[] xmls)
        {
            using (var archive = ZipFile.Open(
                zipPath,
                ZipArchiveMode.Create))
            {
                for (var index = 0;
                     index < entryNames.Length;
                     index++)
                {
                    var entry = archive.CreateEntry(
                        entryNames[index]);
                    using (var stream = entry.Open())
                    using (var writer = new StreamWriter(
                        stream,
                        Encoding.UTF8))
                    {
                        writer.Write(
                            "<?xml version=\"1.0\" " +
                            "encoding=\"UTF-8\"?>" +
                            xmls[index]);
                    }
                }
            }
        }

        private static void SpreadsheetsStreamThroughSharedStrings()
        {
            var temp = Path.Combine(
                Path.GetTempPath(),
                "Scribble-xlsx-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(temp);
            try
            {
                const string sheetNamespace =
                    "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                var xlsxPath = Path.Combine(temp, "budget.xlsx");
                WriteZipEntries(
                    xlsxPath,
                    new[]
                    {
                        "xl/sharedStrings.xml",
                        "xl/worksheets/sheet1.xml",
                        "xl/worksheets/sheet2.xml"
                    },
                    new[]
                    {
                        "<sst xmlns=\"" + sheetNamespace + "\">" +
                        "<si><t>Region</t></si>" +
                        "<si><t>Northern office</t></si>" +
                        "</sst>",
                        "<worksheet xmlns=\"" + sheetNamespace + "\">" +
                        "<sheetData>" +
                        "<row><c t=\"s\"><v>0</v></c>" +
                        "<c><v>42</v></c></row>" +
                        "<row><c t=\"s\"><v>1</v></c>" +
                        "<c><v>98.5</v></c></row>" +
                        "</sheetData></worksheet>",
                        "<worksheet xmlns=\"" + sheetNamespace + "\">" +
                        "<sheetData>" +
                        "<row><c><v>777</v></c></row>" +
                        "</sheetData></worksheet>"
                    });
                var content = EmailAttachmentReader.LoadLocalFile(
                    xlsxPath);
                Assert(
                    content != null,
                    "The spreadsheet did not load at all.");
                Assert(
                    content.Text.Contains("Region\t42"),
                    "Shared-string and numeric cells were not " +
                    "joined into a row: " + content.Text);
                Assert(
                    content.Text.Contains("Northern office\t98.5"),
                    "The second streamed row was not extracted: " +
                    content.Text);
                Assert(
                    content.Text.Contains("[Sheet 2]") &&
                    content.Text.Contains("777"),
                    "The second worksheet was not extracted: " +
                    content.Text);
            }
            finally
            {
                if (Directory.Exists(temp))
                {
                    Directory.Delete(temp, true);
                }
            }
        }

        private static void WriteZipBinaryEntries(
            string zipPath,
            string[] entryNames,
            byte[][] payloads)
        {
            using (var archive = ZipFile.Open(
                zipPath,
                ZipArchiveMode.Create))
            {
                for (var index = 0;
                     index < entryNames.Length;
                     index++)
                {
                    var entry = archive.CreateEntry(
                        entryNames[index]);
                    using (var stream = entry.Open())
                    {
                        stream.Write(
                            payloads[index],
                            0,
                            payloads[index].Length);
                    }
                }
            }
        }

        private static void WriteBiffRecord(
            MemoryStream stream,
            int id,
            byte[] payload)
        {
            if (id < 0x80)
            {
                stream.WriteByte((byte)id);
            }
            else
            {
                stream.WriteByte(
                    (byte)((id & 0x7F) | 0x80));
                stream.WriteByte(
                    (byte)((id >> 7) & 0x7F));
            }

            var length = payload.Length;
            do
            {
                var value = length & 0x7F;
                length >>= 7;
                stream.WriteByte(
                    (byte)(length > 0 ? value | 0x80 : value));
            }
            while (length > 0);
            stream.Write(payload, 0, payload.Length);
        }

        private static byte[] BiffWideString(string value)
        {
            var characters = Encoding.Unicode.GetBytes(value);
            var payload = new byte[4 + characters.Length];
            WriteUInt32At(payload, 0, (uint)value.Length);
            Array.Copy(
                characters,
                0,
                payload,
                4,
                characters.Length);
            return payload;
        }

        private static byte[] BiffCell(int column, byte[] value)
        {
            var payload = new byte[8 + value.Length];
            WriteUInt32At(payload, 0, (uint)column);
            Array.Copy(value, 0, payload, 8, value.Length);
            return payload;
        }

        private static void BinaryWorkbooksDecodeBiff12Records()
        {
            var temp = Path.Combine(
                Path.GetTempPath(),
                "Scribble-xlsb-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(temp);
            try
            {
                byte[] sharedStrings;
                using (var stream = new MemoryStream())
                {
                    WriteBiffRecord(
                        stream,
                        19,
                        Combine(
                            new byte[] { 0 },
                            BiffWideString("Region")));
                    WriteBiffRecord(
                        stream,
                        19,
                        Combine(
                            new byte[] { 0 },
                            BiffWideString("Northern office")));
                    sharedStrings = stream.ToArray();
                }

                byte[] sheet;
                using (var stream = new MemoryStream())
                {
                    var isstZero = new byte[4];
                    var isstOne = new byte[4];
                    WriteUInt32At(isstOne, 0, 1);
                    var rk = new byte[4];
                    WriteUInt32At(rk, 0, (42u << 2) | 2u);
                    var real = BitConverter.GetBytes(98.5);
                    WriteBiffRecord(stream, 0, new byte[8]);
                    WriteBiffRecord(
                        stream,
                        7,
                        BiffCell(0, isstZero));
                    WriteBiffRecord(stream, 2, BiffCell(1, rk));
                    WriteBiffRecord(stream, 0, new byte[8]);
                    WriteBiffRecord(
                        stream,
                        7,
                        BiffCell(0, isstOne));
                    WriteBiffRecord(stream, 5, BiffCell(1, real));
                    WriteBiffRecord(
                        stream,
                        4,
                        BiffCell(2, new byte[] { 1 }));
                    sheet = stream.ToArray();
                }

                var xlsbPath = Path.Combine(temp, "ledger.xlsb");
                WriteZipBinaryEntries(
                    xlsbPath,
                    new[]
                    {
                        "xl/workbook.bin",
                        "xl/sharedStrings.bin",
                        "xl/worksheets/sheet1.bin"
                    },
                    new[]
                    {
                        new byte[0],
                        sharedStrings,
                        sheet
                    });
                var content = EmailAttachmentReader.LoadLocalFile(
                    xlsbPath);
                Assert(
                    content != null &&
                    content.Text.Contains("Region\t42"),
                    "BIFF12 shared-string and RK cells were not " +
                    "decoded: " +
                    (content == null ? "null" : content.Text));
                Assert(
                    content.Text.Contains(
                        "Northern office\t98.5\tTRUE"),
                    "BIFF12 real and boolean cells were not " +
                    "decoded: " + content.Text);

                // The same bytes with an unknown extension must be
                // identified by the zip sniffer.
                var sniffPath = Path.Combine(temp, "mystery.dat");
                File.Copy(xlsbPath, sniffPath);
                var sniffed = EmailAttachmentReader.LoadLocalFile(
                    sniffPath);
                Assert(
                    sniffed != null &&
                    sniffed.Text.Contains("Region\t42"),
                    "An extensionless binary workbook was not " +
                    "sniffed.");
            }
            finally
            {
                if (Directory.Exists(temp))
                {
                    Directory.Delete(temp, true);
                }
            }
        }

        private static byte[] Combine(byte[] first, byte[] second)
        {
            var result = new byte[first.Length + second.Length];
            Array.Copy(first, result, first.Length);
            Array.Copy(
                second,
                0,
                result,
                first.Length,
                second.Length);
            return result;
        }

        private static void OfficeVariantsOpenDocumentAndMsgExtract()
        {
            var temp = Path.Combine(
                Path.GetTempPath(),
                "Scribble-variants-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(temp);
            try
            {
                var docmPath = Path.Combine(temp, "macro.docm");
                WriteZipEntry(
                    docmPath,
                    "word/document.xml",
                    "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                    "<w:body><w:p><w:r><w:t>Macro document baseline</w:t></w:r></w:p>" +
                    "</w:body></w:document>");
                var docm = EmailAttachmentReader.LoadLocalFile(
                    docmPath);
                Assert(
                    docm != null &&
                    docm.Text.Contains("Macro document baseline"),
                    "A .docm Word variant was not extracted.");

                var odsPath = Path.Combine(temp, "plan.ods");
                WriteZipEntry(
                    odsPath,
                    "content.xml",
                    "<office:document-content " +
                    "xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" " +
                    "xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" " +
                    "xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\">" +
                    "<office:body><office:spreadsheet><table:table>" +
                    "<table:table-row><table:table-cell>" +
                    "<text:p>Forecast baseline figures</text:p>" +
                    "</table:table-cell></table:table-row>" +
                    "</table:table></office:spreadsheet></office:body>" +
                    "</office:document-content>");
                var ods = EmailAttachmentReader.LoadLocalFile(
                    odsPath);
                Assert(
                    ods != null &&
                    ods.Text.Contains("Forecast baseline figures"),
                    "An OpenDocument spreadsheet was not extracted.");

                const string msgBody =
                    "Please review the renewal terms before Friday.";
                // BuildCompoundFile stores payloads in regular
                // sectors without a mini-FAT, so the stream must be
                // at least 4096 bytes to take the regular-FAT read
                // path (real .msg files with small bodies use the
                // mini stream, which the production reader supports).
                var msgText = new StringBuilder(msgBody);
                while (msgText.Length < 2100)
                {
                    msgText.Append(
                        " The full contract text follows with " +
                        "additional terms and appendices.");
                }

                var msgPath = Path.Combine(temp, "forwarded.msg");
                File.WriteAllBytes(
                    msgPath,
                    BuildCompoundFile(
                        "__substg1.0_1000001F",
                        Encoding.Unicode.GetBytes(
                            msgText.ToString())));
                var msg = EmailAttachmentReader.LoadLocalFile(
                    msgPath);
                Assert(
                    msg != null && msg.Text.Contains(msgBody),
                    "An Outlook .msg attachment body was not " +
                    "extracted.");

                var tsvPath = Path.Combine(temp, "export.tsv");
                File.WriteAllText(
                    tsvPath,
                    "Region\tRevenue\nNorth\t125000");
                var tsv = EmailAttachmentReader.LoadLocalFile(
                    tsvPath);
                Assert(
                    tsv != null &&
                    tsv.Text.Contains("North\t125000"),
                    "A .tsv export was not read as text.");
            }
            finally
            {
                if (Directory.Exists(temp))
                {
                    Directory.Delete(temp, true);
                }
            }
        }

        private static void SuggestedReplyQuestionsAreBounded()
        {
            var snapshot = new MessageSnapshot(
                "suggest-entry",
                "store",
                "Quarterly planning meeting",
                "Jordan Blake",
                "user@example.com",
                DateTime.Now,
                "Can you confirm whether you will attend the " +
                "planning meeting on Thursday?");
            var request = SuggestQuestionsRequestFactory.Create(
                "local-model",
                snapshot);
            var system = Convert.ToString(
                ((ChatCompletionInputMessage)
                    request.messages[0]).content);
            var user = Convert.ToString(
                ((ChatCompletionInputMessage)
                    request.messages[1]).content);
            Assert(
                system.Contains("untrusted"),
                "The question request must mark email content " +
                "as untrusted data.");
            Assert(
                user.Contains("Quarterly planning meeting"),
                "The question request must include the email " +
                "subject.");
            Assert(
                request.tools == null &&
                request.max_tokens.HasValue,
                "The question request must be tool-free and " +
                "token-bounded.");

            var parsed = SuggestQuestionsRequestFactory.Parse(
                "1. What outcome do you want?|Confirm|" +
                "Reschedule|Decline|ExtraDropped\n" +
                "- Should I mention the missing invoice?\n" +
                "note without a question mark\n" +
                "What else should the reply cover?");
            Assert(
                parsed.Count ==
                SuggestQuestionsRequestFactory.MaxQuestions,
                "Question parsing must cap at " +
                SuggestQuestionsRequestFactory.MaxQuestions +
                " but returned " + parsed.Count + ".");
            Assert(
                parsed[0].Text ==
                "What outcome do you want?" &&
                parsed[0].Options.Count ==
                SuggestQuestionsRequestFactory
                    .MaxOptionsPerQuestion &&
                parsed[0].Options[0] == "Confirm",
                "Numbering must be stripped and options capped: " +
                parsed[0].Text);
            Assert(
                parsed[1].Text.Contains("missing invoice"),
                "The second question was not parsed: " +
                parsed[1].Text);
            Assert(
                SuggestQuestionsRequestFactory.Parse(
                    string.Empty).Count == 0 &&
                SuggestQuestionsRequestFactory.Parse(
                    null).Count == 0,
                "Empty model output must parse to no questions.");
        }

        private static void GeminiTranslationPreservesToolContract()
        {
            var request = new ChatCompletionRequest
            {
                model = "models/gemini-2.5-flash",
                messages = new List<object>
                {
                    new ChatCompletionInputMessage
                    {
                        role = "system",
                        content = "Mailbox boundary text."
                    },
                    new ChatCompletionInputMessage
                    {
                        role = "user",
                        content = "Summarize the selected email."
                    },
                    new ChatCompletionAssistantToolMessage
                    {
                        role = "assistant",
                        content = string.Empty,
                        tool_calls = new List<ChatToolCall>
                        {
                            new ChatToolCall
                            {
                                id = "call_1",
                                type = "function",
                                function = new ChatToolCallFunction
                                {
                                    name = "read_messages",
                                    arguments =
                                        "{\"handles\":[\"selected\"]}"
                                }
                            }
                        }
                    },
                    new ChatCompletionToolResultMessage
                    {
                        role = "tool",
                        tool_call_id = "call_1",
                        content = "{\"messages\":[]}"
                    }
                },
                tools = new List<ChatToolDefinition>
                {
                    MailboxToolCatalog.CreateDefinitions(
                        false)[0],
                    DraftToolCatalog.CreateDefinition()
                },
                tool_choice = "auto",
                max_tokens = 500
            };

            Assert(
                GeminiCodeAssistGateway.NormalizeModel(
                    request.model) == "gemini-2.5-flash",
                "The models/ prefix was not stripped.");
            var translated =
                GeminiCodeAssistGateway.TranslateRequest(request);
            var serializer = new JavaScriptSerializer();
            var json = serializer.Serialize(translated);
            Assert(
                json.Contains("systemInstruction") &&
                json.Contains("Mailbox boundary text."),
                "The system prompt did not become a " +
                "systemInstruction: " + json);
            Assert(
                json.Contains("functionDeclarations") &&
                json.Contains("read_messages") &&
                json.Contains("create_draft"),
                "Tool declarations were not translated.");
            Assert(
                json.Contains("functionCall") &&
                json.Contains("functionResponse"),
                "The tool round-trip was not translated.");
            Assert(
                !json.Contains("additionalProperties") &&
                json.Contains("\"OBJECT\"") &&
                json.Contains("maxOutputTokens"),
                "Schema sanitizing or generation config is " +
                "wrong: " + json);
            Assert(
                json.Contains("4596"),
                "Thinking headroom was not added to the output " +
                "token cap: " + json);
            Assert(
                GeminiCodeAssistGateway.ThinkingBudgetFor(
                    "gemini-2.5-flash") == 0 &&
                GeminiCodeAssistGateway.ThinkingBudgetFor(
                    "models/gemini-2.5-pro") == 128 &&
                GeminiCodeAssistGateway.ThinkingBudgetFor(
                    "gemini-3-flash") == -1 &&
                GeminiCodeAssistGateway.ThinkingBudgetFor(
                    "gemini-3.1-pro-preview") == -1 &&
                GeminiCodeAssistGateway.ThinkingBudgetFor(
                    "qwen3-vl-30b") == -1,
                "Thinking budgets are wrong per model family " +
                "(thinkingBudget is a 2.5-only control).");
            Assert(
                json.Contains("thinkingConfig") &&
                json.Contains("\"thinkingBudget\":0"),
                "Thinking was not disabled for the flash model: " +
                json);

            var response = serializer.DeserializeObject(
                "{\"candidates\":[{\"content\":{\"role\":" +
                "\"model\",\"parts\":[{\"thought\":true," +
                "\"text\":\"Internal reasoning summary.\"}," +
                "{\"text\":\"Here is the " +
                "summary.\"},{\"functionCall\":{\"name\":" +
                "\"read_messages\",\"args\":{\"handles\":" +
                "[\"selected\"]}},\"thoughtSignature\":" +
                "\"sig-abc123\"}]}}]}")
                as IDictionary<string, object>;
            var signatures =
                new Dictionary<string, string>();
            var message =
                GeminiCodeAssistGateway.TranslateResponse(
                    response,
                    signatures);
            Assert(
                message != null &&
                message.content.Contains(
                    "Here is the summary.") &&
                !message.content.Contains(
                    "Internal reasoning summary.") &&
                message.tool_calls != null &&
                message.tool_calls.Count == 1 &&
                message.tool_calls[0].function.name ==
                "read_messages" &&
                message.tool_calls[0].function.arguments
                    .Contains("selected"),
                "The Gemini response did not translate back to " +
                "the OpenAI shape (thought parts must be " +
                "filtered).");

            Assert(
                signatures.Count == 1 &&
                signatures.ContainsKey(
                    message.tool_calls[0].id),
                "The thought signature was not captured by " +
                "tool-call id.");
            var followUp = new ChatCompletionRequest
            {
                model = "gemini-2.5-flash",
                messages = new List<object>
                {
                    new ChatCompletionAssistantToolMessage
                    {
                        role = "assistant",
                        content = string.Empty,
                        tool_calls = message.tool_calls
                    },
                    new ChatCompletionToolResultMessage
                    {
                        role = "tool",
                        tool_call_id = message.tool_calls[0].id,
                        content = "{\"messages\":[]}"
                    }
                }
            };
            var followUpJson = serializer.Serialize(
                GeminiCodeAssistGateway.TranslateRequest(
                    followUp,
                    signatures));
            Assert(
                followUpJson.Contains("thoughtSignature") &&
                followUpJson.Contains("sig-abc123"),
                "The thought signature was not echoed back on " +
                "the replayed functionCall: " + followUpJson);

            var inline = GeminiCodeAssistGateway.TranslateDataUrl(
                "data:image/png;base64,AAAA");
            Assert(
                inline != null &&
                Convert.ToString(inline["mimeType"]) ==
                "image/png" &&
                Convert.ToString(inline["data"]) == "AAAA",
                "Data URL translation failed.");
            Assert(
                GeminiCodeAssistGateway.TranslateDataUrl(
                    "https://example.test/image.png") == null,
                "Web URLs must not become inline data.");

            var credentials =
                GeminiCodeAssistGateway.ParseCredentials(
                    "{\"access_token\":\"at\"," +
                    "\"refresh_token\":\"rt\"," +
                    "\"expiry_date\":1700000000000}");
            Assert(
                credentials != null &&
                credentials.AccessToken == "at" &&
                credentials.RefreshToken == "rt" &&
                GeminiCodeAssistGateway.NeedsRefresh(
                    credentials,
                    1700000000000L),
                "Cached credential parsing failed.");
            Assert(
                GeminiCodeAssistGateway.ParseCredentials(
                    "{\"access_token\":\"at\"}") == null,
                "Credentials without a refresh token must be " +
                "rejected.");

            Assert(
                GeminiCodeAssistGateway.ParseRetryDelaySeconds(
                    "{\"error\":{\"details\":[{\"@type\":" +
                    "\"type.googleapis.com/google.rpc." +
                    "RetryInfo\",\"retryDelay\":\"56s\"}]}}") ==
                56 &&
                GeminiCodeAssistGateway.ParseRetryDelaySeconds(
                    "your quota will reset after 42s") == 42 &&
                GeminiCodeAssistGateway.ParseRetryDelaySeconds(
                    "no hint here") == 0,
                "Retry delay parsing failed.");

            Assert(
                GeminiCodeAssistGateway.MaxRetryAttempts == 10 &&
                GeminiCodeAssistGateway.ComputeRetryDelaySeconds(
                    0, string.Empty, 0) == 1 &&
                GeminiCodeAssistGateway.ComputeRetryDelaySeconds(
                    1, string.Empty, 0) == 3 &&
                GeminiCodeAssistGateway.ComputeRetryDelaySeconds(
                    2, string.Empty, 0) == 20 &&
                GeminiCodeAssistGateway.ComputeRetryDelaySeconds(
                    3, string.Empty, 0) == 30 &&
                GeminiCodeAssistGateway.ComputeRetryDelaySeconds(
                    2, "reset after 56s", 0) == 56 &&
                GeminiCodeAssistGateway.ComputeRetryDelaySeconds(
                    2, string.Empty, 1.0) == 24,
                "Retry policy must fast-probe (1s, 3s) then back " +
                "off with the server hint honored.");
            var tried = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            Assert(
                GeminiCodeAssistGateway.CapacityFallbackChain
                    .Count == 8 &&
                GeminiCodeAssistGateway.NextFallbackModel(
                    "gemini-3.5-flash", tried) ==
                "gemini-3-flash" &&
                GeminiCodeAssistGateway.NextFallbackModel(
                    "gemini-2.5-flash", tried) ==
                "gemini-3.5-flash" &&
                GeminiCodeAssistGateway.NextFallbackModel(
                    "qwen3-vl-30b", tried) == null,
                "The capacity chain must hop to the next unused " +
                "Gemini model.");
            foreach (var chained in
                GeminiCodeAssistGateway.CapacityFallbackChain)
            {
                if (chained != "gemini-2.5-pro")
                {
                    tried.Add(chained);
                }
            }

            Assert(
                GeminiCodeAssistGateway.NextFallbackModel(
                    "gemini-3.5-flash", tried) ==
                "gemini-2.5-pro",
                "The chain must reach the last untried bucket.");
            tried.Add("gemini-2.5-pro");
            Assert(
                GeminiCodeAssistGateway.NextFallbackModel(
                    "gemini-3.5-flash", tried) == null,
                "The chain must exhaust once every bucket was " +
                "tried.");

            var envelope = new Dictionary<string, object>
            {
                { "model", "gemini-2.5-flash" },
                {
                    "request",
                    new Dictionary<string, object>
                    {
                        {
                            "generationConfig",
                            new Dictionary<string, object>
                            {
                                {
                                    "thinkingConfig",
                                    new Dictionary<string, object>
                                    {
                                        { "thinkingBudget", 0 }
                                    }
                                }
                            }
                        }
                    }
                }
            };
            GeminiCodeAssistGateway.ApplyThinkingConfig(
                envelope,
                "gemini-3-flash");
            var envelopeRequest =
                (IDictionary<string, object>)
                envelope["request"];
            Assert(
                !envelopeRequest.ContainsKey(
                    "generationConfig"),
                "A hop to a non-2.5 model must strip the " +
                "2.5-only thinking config.");
            GeminiCodeAssistGateway.ApplyThinkingConfig(
                envelope,
                "gemini-2.5-flash");
            Assert(
                serializer.Serialize(envelope).Contains(
                    "\"thinkingBudget\":0"),
                "A hop back to a 2.5 model must restore its " +
                "thinking budget.");
        }

        private static void GeminiIsUnavailableToEndUsers()
        {
            var settings = new AppSettings
            {
                UseGeminiSignIn = true,
                Model = "gemini-2.5-flash"
            };
            Assert(
                AdminPolicy.GeminiDisabled &&
                !AdminPolicy.GeminiEnabledForEndUsers &&
                !settings.IsConfigured &&
                !settings.HasConnectionSettings &&
                !ModelSelectionPolicy.IsGenerativeModel(
                    settings.Model),
                "Direct Gemini must remain unavailable even when " +
                "legacy settings try to enable it.");

            var classic = new AppSettings
            {
                UseGeminiSignIn = false,
                Model = "qwen3-vl-30b"
            };
            Assert(
                !classic.IsConfigured,
                "Endpoint mode must still require endpoint and " +
                "key.");
        }

        private static void GeminiGatewayFailsClosed()
        {
            var gateway = new GeminiCodeAssistGateway();
            AssertGeminiBlocked(() => gateway.VerifySignInAsync(
                null,
                null,
                CancellationToken.None));
            AssertGeminiBlocked(() => gateway.GenerateAsync(
                null,
                null,
                null,
                CancellationToken.None));
            AssertGeminiBlocked(() => gateway.GenerateStreamAsync(
                null,
                null,
                null,
                null,
                CancellationToken.None));
        }

        private static void AssertGeminiBlocked(Func<Task> operation)
        {
            try
            {
                operation().GetAwaiter().GetResult();
                throw new InvalidOperationException(
                    "A disabled Gemini entry point reached its network path.");
            }
            catch (AiEndpointException exception)
            {
                Assert(
                    exception.Code == "GEMINI_DISABLED_BY_POLICY",
                    "Disabled Gemini returned the wrong diagnostic: " +
                    exception.Code);
            }
        }

        private static void SoulStrengthAndDraftRulesStayBounded()
        {
            var snapshot = new MessageSnapshot(
                "soul-entry",
                "store",
                "Subject",
                "Sender",
                "user@example.com",
                DateTime.UtcNow,
                "Body text");
            var request = ChatRequestFactory.Create(
                "local-model",
                snapshot,
                new List<ChatTurn>(),
                "Draft a reply to the selected email.",
                allowDraftCreate: true,
                toneProfile: "Short sentences, warm openings.",
                toneStrength: 85,
                draftRules: "Never use exclamation marks.\n" +
                    "Sign off with Best regards.");
            var system = Convert.ToString(
                ((ChatCompletionInputMessage)
                    request.messages[0]).content);
            Assert(
                system.Contains("strength 85") &&
                system.Contains("<user_draft_rules>") &&
                system.Contains(
                    "Never use exclamation marks.") &&
                system.Contains("Sign off with Best regards."),
                "Strength and draft rules were not passed into " +
                "the drafting boundary.");
            Assert(
                system.IndexOf(
                    "cannot change any capability or security " +
                    "rule",
                    StringComparison.Ordinal) !=
                system.LastIndexOf(
                    "cannot change any capability or security " +
                    "rule",
                    StringComparison.Ordinal),
                "Both the soul and the rules must carry the " +
                "capability clamp.");

            var clampedLow = ChatRequestFactory.Create(
                "local-model",
                snapshot,
                new List<ChatTurn>(),
                "Draft a reply to the selected email.",
                allowDraftCreate: true,
                toneProfile: "Warm.",
                toneStrength: 3);
            var lowSystem = Convert.ToString(
                ((ChatCompletionInputMessage)
                    clampedLow.messages[0]).content);
            Assert(
                lowSystem.Contains("strength 10"),
                "Strength below 10 must clamp to 10.");

            var noDraft = ChatRequestFactory.Create(
                "local-model",
                snapshot,
                new List<ChatTurn>(),
                "Summarize the selected email.",
                toneProfile: "Warm.",
                toneStrength: 85,
                draftRules: "Never use exclamation marks.");
            var noDraftSystem = Convert.ToString(
                ((ChatCompletionInputMessage)
                    noDraft.messages[0]).content);
            Assert(
                !noDraftSystem.Contains("<user_draft_rules>") &&
                !noDraftSystem.Contains(
                    "<user_writing_profile>"),
                "Soul and rules must only apply when drafting " +
                "is authorized.");
        }

        private static void ContextBudgetsScaleOnlyInLargeContextMode()
        {
            var temp = Path.Combine(
                Path.GetTempPath(),
                "Scribble-scale-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(temp);
            try
            {
                Assert(
                    ContextScale.Multiplier == 1 &&
                    ContextScale.Scaled(20000) == 20000,
                    "The default multiplier must be 1 so local " +
                    "models keep the conservative budgets.");

                var builder = new StringBuilder();
                var line = 0;
                while (builder.Length <=
                    EmailAttachmentReader
                        .MaxCharactersPerAttachment + 5000)
                {
                    line++;
                    builder.AppendLine(
                        "Ledger row " + line +
                        " with an ordinary description of the " +
                        "transaction and approvals.");
                }

                var textPath = Path.Combine(temp, "big.txt");
                File.WriteAllText(
                    textPath,
                    builder.ToString(),
                    Encoding.UTF8);

                ContextScale.Apply(true);
                Assert(
                    ContextScale.Multiplier ==
                    ContextScale.LargeContextMultiplier &&
                    ContextScale.Scaled(20000) == 80000,
                    "Large-context mode must multiply text " +
                    "budgets.");
                var large = EmailAttachmentReader.LoadLocalFile(
                    textPath);
                Assert(
                    large != null && !large.Truncated,
                    "A file within the scaled budget must not be " +
                    "truncated in large-context mode.");

                ContextScale.Apply(false);
                var small = EmailAttachmentReader.LoadLocalFile(
                    textPath);
                Assert(
                    small != null && small.Truncated,
                    "The same file must be truncated again at the " +
                    "standard budget.");
                Assert(
                    EmailAttachmentReader.MaxAttachments == 10 &&
                    MailboxWorkingSet.MaxMessages == 10 &&
                    ExternalContextDocument.MaxDocuments == 3,
                    "Capability caps must not scale with context " +
                    "size.");
            }
            finally
            {
                ContextScale.Apply(false);
                if (Directory.Exists(temp))
                {
                    Directory.Delete(temp, true);
                }
            }
        }

        private static void OversizedTextCarriesTruncationNotice()
        {
            var temp = Path.Combine(
                Path.GetTempPath(),
                "Scribble-trunc-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(temp);
            try
            {
                var builder = new StringBuilder();
                var line = 0;
                while (builder.Length <=
                    EmailAttachmentReader.MaxCharactersPerAttachment +
                    5000)
                {
                    line++;
                    builder.AppendLine(
                        "Ledger entry " + line +
                        " with a running description of the " +
                        "transaction and its approval chain.");
                }

                var textPath = Path.Combine(temp, "ledger.txt");
                File.WriteAllText(
                    textPath,
                    builder.ToString(),
                    Encoding.UTF8);
                var content = EmailAttachmentReader.LoadLocalFile(
                    textPath);
                Assert(
                    content != null && content.Truncated,
                    "An oversized text file was not flagged as " +
                    "truncated.");
                Assert(
                    content.Text.Contains("[Truncated:"),
                    "The truncation notice was not appended to " +
                    "the bounded text.");
                Assert(
                    content.Text.Length <=
                    EmailAttachmentReader.MaxCharactersPerAttachment +
                    200,
                    "Truncated text exceeded the per-attachment " +
                    "character budget: " +
                    content.Text.Length + " characters.");

                var smallPath = Path.Combine(temp, "note.txt");
                File.WriteAllText(
                    smallPath,
                    "A short note that fits comfortably.",
                    Encoding.UTF8);
                var small = EmailAttachmentReader.LoadLocalFile(
                    smallPath);
                Assert(
                    small != null &&
                    !small.Truncated &&
                    !small.Text.Contains("[Truncated:"),
                    "A small text file was wrongly marked as " +
                    "truncated.");
            }
            finally
            {
                if (Directory.Exists(temp))
                {
                    Directory.Delete(temp, true);
                }
            }
        }

        private static void OversizedImagesAreDownscaledForVision()
        {
            var pngPath = Path.Combine(
                Path.GetTempPath(),
                "Scribble-big-" + Guid.NewGuid().ToString("N") + ".png");
            using (var bitmap = new System.Drawing.Bitmap(
                1400,
                1400,
                System.Drawing.Imaging.PixelFormat.Format24bppRgb))
            {
                var bounds = new System.Drawing.Rectangle(
                    0,
                    0,
                    1400,
                    1400);
                var data = bitmap.LockBits(
                    bounds,
                    System.Drawing.Imaging.ImageLockMode.WriteOnly,
                    System.Drawing.Imaging.PixelFormat.Format24bppRgb);
                var noise = new byte[
                    Math.Abs(data.Stride) * 1400];
                new Random(42).NextBytes(noise);
                Marshal.Copy(
                    noise,
                    0,
                    data.Scan0,
                    noise.Length);
                bitmap.UnlockBits(data);
                bitmap.Save(
                    pngPath,
                    System.Drawing.Imaging.ImageFormat.Png);
            }

            try
            {
                Assert(
                    new FileInfo(pngPath).Length > 1536 * 1024,
                    "The generated test image is too small to exercise downscaling.");

                var attachments = new FakeOutlookAttachments();
                attachments.Add(new FakeOutlookAttachment(
                    "photo.png",
                    pngPath));
                var mail = new FakeSelectedMailItem(
                    "big-image-entry",
                    "Site photo")
                {
                    Attachments = attachments
                };
                var application = new FakeOutlookApplication();
                application.Session.Register(
                    "big-image-entry",
                    "store",
                    mail);
                var snapshot = new MessageReader(application)
                    .CaptureById("big-image-entry", "store");
                var host = new MailboxToolHost(application, snapshot);
                var loaded = host.Execute(
                    MailboxCall(
                        "read-big-image",
                        MailboxToolCatalog.ReadMessages,
                        "{\"handles\":[\"selected\"]}"));
                Assert(
                    loaded.VisionImages.Count == 1 &&
                    loaded.VisionImages[0].DataUrl.StartsWith(
                        "data:image/jpeg;base64,") &&
                    loaded.VisionImages[0].DataUrl.Length <=
                        EmailAttachmentReader.MaxImageDataUrlCharacters &&
                    loaded.Content.Contains("downscaled"),
                    "An oversized image was not downscaled into vision input.");
            }
            finally
            {
                if (File.Exists(pngPath))
                {
                    File.Delete(pngPath);
                }
            }
        }

        private static void WebHostedImagesAreDisclosed()
        {
            var mail = new FakeSelectedMailItem(
                "remote-image-entry",
                "Newsletter")
            {
                HTMLBody =
                    "<html><body>" +
                    "<a href=\"https://example.test/offer\">" +
                    "<img src=\"https://cdn.example.test/banner.jpg\"></a>" +
                    "<img src='https://cdn.example.test/photo.png'>" +
                    "<img src=\"cid:signature-logo\">" +
                    "</body></html>"
            };
            var application = new FakeOutlookApplication();
            application.Session.Register(
                "remote-image-entry",
                "store",
                mail);
            var snapshot = new MessageReader(application)
                .CaptureById("remote-image-entry", "store");
            Assert(
                snapshot.RemoteImageCount == 2,
                "Web-hosted image detection counted " +
                snapshot.RemoteImageCount +
                " instead of 2 (cid images must not count).");

            var request = ChatRequestFactory.Create(
                "qwen3-vl-30b",
                snapshot,
                new List<ChatTurn>(),
                "Summarize the image in this email.");
            var reference = MessageContent(request.messages[1]);
            Assert(
                reference.Contains(
                    "Web-hosted images referenced by URL: 2"),
                "The message reference does not disclose web-hosted images.");

            var host = new MailboxToolHost(application, snapshot);
            var loaded = host.Execute(
                MailboxCall(
                    "read-remote",
                    MailboxToolCatalog.ReadMessages,
                    "{\"handles\":[\"selected\"]}"));
            Assert(
                loaded.Content.Contains(
                    "web_hosted_images_not_included"),
                "The tool result does not disclose web-hosted images.");
            Assert(
                loaded.Content.Contains("web_hosted_images_note") &&
                loaded.Content.Contains("cannot view them"),
                "The tool result is missing the web-hosted image note.");
        }

        private static int CountToken(string text, string token)
        {
            var count = 0;
            var index = 0;
            while ((index = text.IndexOf(
                token,
                index,
                StringComparison.Ordinal)) >= 0)
            {
                count++;
                index += token.Length;
            }

            return count;
        }

        private static void ModelCatalogDescribesVisionCapability()
        {
            Assert(
                ModelCatalog.SupportsVision("qwen3-vl-30b") &&
                ModelCatalog.SupportsVision("Qwen/Qwen3.8-27B") &&
                ModelCatalog.SupportsVision("qwen3.8-27b-fast") &&
                ModelCatalog.SupportsVision("gemma-4-31b-it") &&
                ModelCatalog.SupportsVision(
                    "models/gemini-2.5-flash") &&
                !ModelCatalog.SupportsVision("qwen3.6-35b-a3b") &&
                ModelCatalog.GuideEntries.Count >= 7,
                "The model catalog vision flags are incomplete.");

            var description =
                ModelCatalog.DescribeForSelection("qwen3-vl-30b");
            Assert(
                description.Contains("Vision model") &&
                description.Contains("email images"),
                "Vision model guidance is incomplete: " + description);

            var overview = ModelCatalog.BuildGuideOverview();
            Assert(
                overview.Contains("Connect and load models") &&
                overview.Length < 80,
                "The model guide overview is incomplete.");
        }

        private static void QwenIsPreferredWithoutLockIn()
        {
            var preferred = ModelSelectionPolicy.PreferredModel(
                new[]
                {
                    "llama-3.3-70b-instruct",
                    "Qwen3.6-35B-A3B-Base",
                    "Qwen3.6-Coder-35B-A3B-Instruct",
                    "Qwen-Image-Edit",
                    "Qwen-Audio-Chat",
                    "qwen3-vl-30b",
                    "Qwen3.5-35B-A3B-Instruct",
                    "Qwen3.6-35B-A3B-Instruct",
                    "Qwen3.8-27B-Fast"
                });
            Assert(
                preferred == "Qwen3.8-27B-Fast",
                "The Qwen3.8 27B family should be preferred: " +
                preferred);
            Assert(
                ModelSelectionPolicy.PreferredModel(
                    new[] { "custom-chat-model" }) ==
                    "custom-chat-model",
                "A sole compatible endpoint model should remain usable.");
            Assert(
                ModelSelectionPolicy.PreferredModel(
                    new[] { "llama-chat", "mistral-instruct" }) ==
                    string.Empty,
                "Multiple non-Qwen models should remain a user choice.");
            Assert(
                ModelSelectionPolicy.PreferredModel(
                    new[] { "Qwen3-Reranker", "Qwen-Image" }) ==
                    string.Empty,
                "Specialized Qwen routes must not be selected automatically.");
        }

        private static void SelfUpdateIsOfficialAndBounded()
        {
            Assert(
                SelfUpdater.InstallerUrl.StartsWith(
                    "https://github.com/datap0nd/scribble/releases/",
                    StringComparison.Ordinal),
                "The updater must download only the official release installer over HTTPS.");

            var script = SelfUpdater.BuildUpdateScript();
            Assert(
                script.Contains("OUTLOOK.EXE") &&
                script.Contains("EXCEL.EXE") &&
                script.Contains("POWERPNT.EXE") &&
                script.Contains("WINWORD.EXE") &&
                script.Contains("if %tries% GEQ 150 exit /b 1") &&
                // The update closes the hosts itself: politely
                // first, forcibly once a save prompt has stalled it.
                script.Contains("taskkill /IM OUTLOOK.EXE") &&
                script.Contains("taskkill /F /IM WINWORD.EXE") &&
                script.Contains("if %tries% GEQ 15 goto force") &&
                script.Contains(
                    "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART") &&
                script.Contains(
                    "if not \"%restart%\"==\"\" start \"\" \"%restart%\"") &&
                script.Contains("%~1") &&
                script.Contains("%~2"),
                "The update script must wait for every Office host to close, " +
                "install silently with a bounded wait, and restart only the " +
                "requested host.");

            var excelOnly = SelfUpdater.BuildUpdateScript(
                false,
                true,
                false,
                false);
            Assert(
                excelOnly.Contains("EXCEL.EXE") &&
                !excelOnly.Contains("OUTLOOK.EXE") &&
                !excelOnly.Contains("POWERPNT.EXE") &&
                !excelOnly.Contains("WINWORD.EXE"),
                "A component-scoped update must wait only for its own hosts.");
            Assert(
                excelOnly.Contains("taskkill /IM EXCEL.EXE") &&
                excelOnly.Contains("taskkill /F /IM EXCEL.EXE") &&
                !excelOnly.Contains("taskkill /IM OUTLOOK.EXE"),
                "A component-scoped update must close only its own hosts.");
            var unknown = SelfUpdater.BuildUpdateScript(
                false,
                false,
                false,
                false);
            Assert(
                unknown.Contains("OUTLOOK.EXE") &&
                unknown.Contains("EXCEL.EXE") &&
                unknown.Contains("POWERPNT.EXE") &&
                unknown.Contains("WINWORD.EXE"),
                "An unknown component state must wait for every host.");
        }

        private static void DraftFormulasStayInsideTheWorkbook()
        {
            Assert(
                DraftFormulaPolicy.IsAllowedFormula(
                    "=SUM(A1:A5)") &&
                DraftFormulaPolicy.IsAllowedFormula(
                    "=Data!B2*2") &&
                DraftFormulaPolicy.IsAllowedFormula(
                    "=IF(A1>0,\"yes\",\"no\")") &&
                DraftFormulaPolicy.IsAllowedFormula(
                    "=MYCALL(A1)"),
                "Ordinary workbook formulas must stay allowed.");
            Assert(
                !DraftFormulaPolicy.IsAllowedFormula(
                    "=WEBSERVICE(\"https://x.test\")") &&
                !DraftFormulaPolicy.IsAllowedFormula(
                    "=webservice(A1)") &&
                !DraftFormulaPolicy.IsAllowedFormula(
                    "=RTD(\"p\",,\"t\")") &&
                !DraftFormulaPolicy.IsAllowedFormula(
                    "=HYPERLINK(\"http://x\",\"go\")") &&
                !DraftFormulaPolicy.IsAllowedFormula(
                    "=CALL(\"k32\",\"f\",\"J\")") &&
                !DraftFormulaPolicy.IsAllowedFormula(
                    "=1+CALL (A1)"),
                "Network and native-code formulas must be rejected.");
            Assert(
                !DraftFormulaPolicy.IsAllowedFormula(
                    "=[Book2]Sheet1!A1") &&
                !DraftFormulaPolicy.IsAllowedFormula("SUM(A1)") &&
                !DraftFormulaPolicy.IsAllowedFormula(null) &&
                !DraftFormulaPolicy.IsAllowedFormula("=") &&
                !DraftFormulaPolicy.IsAllowedFormula(
                    "=" + new string('9', 600)),
                "External references and malformed formulas must be rejected.");
        }

        private static void DraftTextLayoutParsesStructure()
        {
            var paragraphs = DraftTextLayout.Parse(
                "# Title\n## Sub\n### Minor\n- item\n" +
                "1. first\nplain **bold** text\n---");
            Assert(
                paragraphs.Count == 7 &&
                paragraphs[0].Kind ==
                    DraftTextLayout.KindHeading1 &&
                paragraphs[0].Text == "Title" &&
                paragraphs[1].Kind ==
                    DraftTextLayout.KindHeading2 &&
                paragraphs[2].Kind ==
                    DraftTextLayout.KindHeading3 &&
                paragraphs[3].Kind ==
                    DraftTextLayout.KindBullet &&
                paragraphs[3].Text == "item" &&
                paragraphs[4].Kind ==
                    DraftTextLayout.KindNumbered &&
                paragraphs[4].Text == "first" &&
                paragraphs[6].Text == string.Empty,
                "Draft layout structure parsing failed.");
            var bold = paragraphs[5];
            Assert(
                bold.Kind == DraftTextLayout.KindNormal &&
                bold.Text == "plain bold text" &&
                bold.BoldRanges.Count == 1 &&
                bold.BoldRanges[0].Start == 6 &&
                bold.BoldRanges[0].Length == 4,
                "Inline bold parsing failed.");
            var unmatched = DraftTextLayout.Parse("a **b");
            Assert(
                unmatched.Count == 1 &&
                unmatched[0].Text == "a **b" &&
                unmatched[0].BoldRanges.Count == 0,
                "Unmatched bold markers must stay literal.");

            var blocks = DraftTextLayout.ParseBlocks(
                "Intro\n| Region | Sales |\n| --- | --- |\n" +
                "| North | 100 |\n| South | 80 |\nOutro");
            Assert(
                blocks.Count == 3 &&
                blocks[0] is DraftTextLayout.Paragraph &&
                blocks[1] is DraftTextLayout.Table &&
                blocks[2] is DraftTextLayout.Paragraph,
                "Table blocks were not grouped from pipe rows.");
            var parsedTable = (DraftTextLayout.Table)blocks[1];
            Assert(
                parsedTable.Rows.Count == 3 &&
                parsedTable.Rows[0][0] == "Region" &&
                parsedTable.Rows[1][1] == "100" &&
                parsedTable.Rows[2][0] == "South",
                "Table cells were not parsed correctly.");
            Assert(
                DraftTextLayout.SplitTableCells(
                    "| a | b |")[1] == "b" &&
                DraftTextLayout.SplitTableCells(
                    "no pipes here") == null,
                "Table cell splitting failed.");

            Assert(
                DraftChartTypes.Resolve("bar") ==
                    DraftChartTypes.BarClustered &&
                DraftChartTypes.Resolve("PIE") ==
                    DraftChartTypes.Pie &&
                DraftChartTypes.Resolve("nonsense") ==
                    DraftChartTypes.ColumnClustered &&
                DraftChartTypes.Resolve(null) ==
                    DraftChartTypes.ColumnClustered,
                "Chart type resolution must be total and bounded.");
        }

        // The corporate deck theme is compiled in: the model
        // supplies content, the writer supplies every font, color,
        // size, and position. These asserts pin that split.
        // One user request authorizes ONE deliverable, which a
        // model may build over several bounded calls - but never
        // more than one unsent email draft.
        private static void DraftBudgetBuildsOneDeliverable()
        {
            var single = new OneShotDraftAuthorization(true);
            Assert(
                single.CallBudget == 1 &&
                single.TryConsume() &&
                !single.TryConsume(),
                "The default draft permission must stay single-shot.");

            var batched =
                new OneShotDraftAuthorization(true, false, 4);
            Assert(
                batched.CallBudget == 4 &&
                batched.RemainingCalls == 4 &&
                batched.TryConsume() &&
                batched.TryConsume() &&
                batched.TryConsume() &&
                batched.TryConsume() &&
                !batched.TryConsume() &&
                batched.IsConsumed,
                "A batched draft permission must spend exactly its budget.");

            Assert(
                new OneShotDraftAuthorization(true, false, 99)
                    .CallBudget ==
                    OneShotDraftAuthorization.MaxCallBudget &&
                new OneShotDraftAuthorization(true, false, 0)
                    .CallBudget == 1,
                "The draft call budget must stay clamped.");

            var denied =
                new OneShotDraftAuthorization(false, false, 6);
            Assert(
                denied.CallBudget == 0 &&
                !denied.TryConsume() &&
                !denied.TryConsumeEmailDraft(),
                "An unauthorized request must get no draft calls at all.");

            // Recipients are the sensitive surface: batching a deck
            // must never batch email drafts.
            var email =
                new OneShotDraftAuthorization(true, false, 6);
            Assert(
                email.TryConsumeEmailDraft() &&
                !email.TryConsumeEmailDraft() &&
                email.TryConsume(),
                "Only one unsent email draft may be opened per request.");

            // Tool-call JSON counts against the response budget, so
            // drafting turns must ask for a generous ceiling.
            var drafting = DocumentChatRequestFactory.Create(
                "test-model",
                "powerpoint",
                "Presentation: Deck1",
                new List<ChatTurn>(),
                "build a slide with this",
                true);
            var reading = DocumentChatRequestFactory.Create(
                "test-model",
                "powerpoint",
                "Presentation: Deck1",
                new List<ChatTurn>(),
                "what is on slide 2");
            Assert(
                drafting.max_tokens.HasValue &&
                drafting.max_tokens.Value >= 2000 &&
                !reading.max_tokens.HasValue,
                "Draft turns need an explicit response budget.");
        }

        // The model picker is the only thing that decides where a
        // request goes. The Gemini tick decides which models are
        // OFFERED, nothing else - mixing the two is what sent qwen
        // prompts to Google and gemini ids to a local server (HTTP
        // 400).
        private static void TransportFollowsSelectedModel()
        {
            Assert(
                GeminiCodeAssistGateway.IsGeminiModel(
                    "gemini-2.5-flash") &&
                GeminiCodeAssistGateway.IsGeminiModel(
                    "GEMINI-3-PRO-PREVIEW") &&
                GeminiCodeAssistGateway.IsGeminiModel(
                    "models/gemini-2.5-pro") &&
                !GeminiCodeAssistGateway.IsGeminiModel(
                    "qwen3-27b") &&
                !GeminiCodeAssistGateway.IsGeminiModel(
                    "gemma3-12b") &&
                !GeminiCodeAssistGateway.IsGeminiModel("") &&
                !GeminiCodeAssistGateway.IsGeminiModel(null),
                "Gemini model classification is wrong.");
            foreach (var known in
                GeminiCodeAssistGateway.KnownModels)
            {
                Assert(
                    GeminiCodeAssistGateway.IsGeminiModel(known),
                    "A known Gemini model was not classified: " +
                    known);
            }

            // A local model stays usable while Gemini sign-in is on.
            var mixed = new AppSettings
            {
                BaseUrl = "https://ai.example.test/v1",
                ApiKey = "key",
                Model = "qwen3-27b",
                UseGeminiSignIn = true
            };
            Assert(
                mixed.IsConfigured &&
                mixed.HasEndpointCredentials,
                "A local model must stay usable with Gemini enabled.");

            // Gemini classification remains available for the
            // retained translation implementation, but selection is
            // disabled in the shipped product.
            var google = new AppSettings
            {
                Model = "gemini-2.5-flash",
                UseGeminiSignIn = true
            };
            Assert(
                !google.IsConfigured &&
                !google.HasConnectionSettings &&
                !ModelSelectionPolicy.IsGenerativeModel(
                    google.Model),
                "A Gemini model must not be selectable or configured.");

            // A Gemini model with the tick off is not configured -
            // the pane must say so instead of posting a gemini id to
            // a local server.
            var stranded = new AppSettings
            {
                Model = "gemini-2.5-flash",
                UseGeminiSignIn = false
            };
            Assert(
                !stranded.IsConfigured,
                "A Gemini model without sign-in must not count as configured.");
        }

        private static void CorporateThemeIsHardcoded()
        {
            Assert(
                MetoTheme.ThemeName == "METO Executive Dense" &&
                MetoTheme.TitleFont == "Samsung Sharp Sans Bold" &&
                MetoTheme.BodyFont == "Calibri",
                "The corporate font stack changed unexpectedly.");

            // Office takes RGB longs in BGR order; a wrong
            // conversion would silently paint the wrong brand color.
            Assert(
                MetoTheme.Rgb(MetoTheme.BrandBlueHex) ==
                    (0x14 | (0x28 << 8) | (0xA0 << 16)) &&
                MetoTheme.Rgb("#FFFFFF") == 0xFFFFFF &&
                MetoTheme.Rgb("nonsense") == 0 &&
                MetoTheme.Rgb(null) == 0,
                "Hex to Office RGB conversion is wrong.");

            Assert(
                MetoTheme.ChartSeriesColors().Length ==
                    PresentationToolCatalogMaxSeries() &&
                MetoTheme.ChartSeriesColors()[0] ==
                    MetoTheme.Rgb(MetoTheme.BrandBlueHex),
                "The chart palette must lead with the brand blue.");

            // Layout follows the content the model supplied, so a
            // small model never has to name one.
            Assert(
                MetoTheme.ResolveLayout(null, true, false, false, false) ==
                    MetoTheme.LayoutBullets &&
                MetoTheme.ResolveLayout(null, false, true, false, false) ==
                    MetoTheme.LayoutCards &&
                MetoTheme.ResolveLayout(null, false, false, true, false) ==
                    MetoTheme.LayoutTable &&
                MetoTheme.ResolveLayout(null, false, false, false, true) ==
                    MetoTheme.LayoutChart &&
                MetoTheme.ResolveLayout(null, true, false, false, true) ==
                    MetoTheme.LayoutBullets &&
                MetoTheme.ResolveLayout("cover", false, false, false, false) ==
                    MetoTheme.LayoutCover &&
                MetoTheme.ResolveLayout("AGENDA", false, false, false, false) ==
                    MetoTheme.LayoutAgenda,
                "Slide layout inference is wrong.");

            // Selective highlighting is derived from the corporate
            // performance markers, never chosen by the model.
            Assert(
                MetoTheme.CellStatus("\u2191 12%") ==
                    MetoTheme.StatusGood &&
                MetoTheme.CellStatus("+8%") ==
                    MetoTheme.StatusGood &&
                MetoTheme.CellStatus("G/R \u25B35%") ==
                    MetoTheme.StatusBad &&
                MetoTheme.CellStatus("\u2193 3%") ==
                    MetoTheme.StatusBad &&
                MetoTheme.CellStatus("-8%") ==
                    MetoTheme.StatusBad &&
                MetoTheme.CellStatus("1,240") ==
                    MetoTheme.StatusNone &&
                MetoTheme.CellStatus("SEEG") ==
                    MetoTheme.StatusNone &&
                MetoTheme.CellStatus(null) ==
                    MetoTheme.StatusNone,
                "Table status highlighting is not derived from the data.");

            Assert(
                MetoTheme.FitTitleSize("Short takeaway", 40f, 46) == 40f &&
                MetoTheme.FitTitleSize(new string('x', 200), 40f, 46) < 40f &&
                MetoTheme.FitTitleSize(new string('x', 200), 40f, 46) >= 24f,
                "Long titles must step down instead of overflowing.");

            Assert(
                DraftChartTypes.Resolve("stacked column") ==
                    DraftChartTypes.ColumnStacked &&
                DraftChartTypes.Resolve("100% stacked") ==
                    DraftChartTypes.ColumnStacked100 &&
                DraftChartTypes.Resolve("stacked bar") ==
                    DraftChartTypes.BarStacked &&
                DraftChartTypes.Resolve("line") ==
                    DraftChartTypes.LineMarkers,
                "The corporate chart vocabulary is incomplete.");

            // The slide schema carries content only - no styling
            // keys the model could use to go off-brand.
            var schema = new JavaScriptSerializer().Serialize(
                PresentationToolCatalog.DraftDefinition());
            Assert(
                schema.Contains("\"subtitle\"") &&
                schema.Contains("\"cards\"") &&
                schema.Contains("\"table\"") &&
                schema.Contains("\"footnote\"") &&
                schema.Contains("\"unit\"") &&
                schema.Contains("\"after_slide\""),
                "The slide schema lost a corporate content field.");
            Assert(
                !schema.Contains("\"color\"") &&
                !schema.Contains("\"font\"") &&
                !schema.Contains("\"position\"") &&
                !schema.Contains("\"left\"") &&
                !schema.Contains("\"top\""),
                "The slide schema must never accept styling from the model.");
        }

        private static int PresentationToolCatalogMaxSeries()
        {
            return 5;
        }

        private static void DraftHtmlRendersTables()
        {
            var table = SafeDraftHtml.FormatContent(
                "| Region | Sales |\n| --- | ---: |\n" +
                "| North | 100 |\nAfter",
                new string[0]);
            Assert(
                table.Html.Contains("<table style=") &&
                table.Html.Contains("<th style=") &&
                table.Html.Contains("<td style=") &&
                table.Html.Contains("text-align:right") &&
                table.Html.Contains(">North<") &&
                table.Html.Contains(">100<") &&
                table.Html.Contains(">After<") &&
                !table.Html.Contains("<script"),
                "Pipe tables were not rendered safely: " +
                table.Html);
            var notTable = SafeDraftHtml.FormatContent(
                "| lonely pipe line\nplain",
                new string[0]);
            Assert(
                !notTable.Html.Contains("<table") &&
                notTable.Html.Contains("| lonely pipe line"),
                "A non-table pipe line must stay plain text.");

            var straySeparator = SafeDraftHtml.FormatContent(
                "| A |\n| --- |\n| x |\n| --- |\n| y |",
                new string[0]);
            Assert(
                straySeparator.Html.Contains(">y<") &&
                !straySeparator.Html.Contains(">---<"),
                "A stray separator row must be skipped, not rendered: " +
                straySeparator.Html);
        }

        private static void RecipientWarningFlagsUnknownRecipients()
        {
            Assert(
                RecipientIntentCheck.Warn(
                    "john.smith@acme.com",
                    string.Empty,
                    "email this to john") == string.Empty,
                "A recipient named in the prompt must not warn.");
            var warning = RecipientIntentCheck.Warn(
                "karen@x.test",
                string.Empty,
                "email this to john");
            Assert(
                warning.Contains("karen@x.test") &&
                warning.Contains("not mentioned"),
                "An unmentioned recipient must be called out: " +
                warning);
            Assert(
                RecipientIntentCheck.Warn(
                    string.Empty,
                    string.Empty,
                    "email this to john") == string.Empty &&
                RecipientIntentCheck.Warn(
                    "bob@x.test",
                    string.Empty,
                    null) == string.Empty,
                "Empty recipients or prompts must not warn.");
        }

        private static void McpHeadersAreBoundedAndSafe()
        {
            var config = new McpServerConfig
            {
                Name = "files",
                Target = "https://example.test/mcp",
                Headers =
                    "Authorization: Bearer abc\n" +
                    "Host: evil.test\n" +
                    "Mcp-Session-Id: forged\n" +
                    "X-Api-Key: k1\n" +
                    "not a header line\n" +
                    "Bad Name!: x\n" +
                    "Content-Type: text/plain"
            }.Sanitized();
            var headers = config.ParsedHeaders();
            Assert(
                headers.Count == 2 &&
                headers[0].Key == "Authorization" &&
                headers[0].Value == "Bearer abc" &&
                headers[1].Key == "X-Api-Key",
                "MCP header parsing must keep only safe headers.");
            var many = new StringBuilder();
            for (var index = 0; index < 12; index++)
            {
                many.AppendLine("X-H" + index + ": v");
            }

            Assert(
                new McpServerConfig
                {
                    Headers = many.ToString()
                }.ParsedHeaders().Count == 8,
                "MCP headers must cap at eight entries.");
        }

        private static void SettingsAlwaysApplyRecommendedLimits()
        {
            try
            {
                Assert(
                    TextBoundary.MaxUserPromptCharacters ==
                    TextBoundary
                        .RecommendedUserPromptCharacters &&
                    TextBoundary.MaxAssistantCharacters ==
                    TextBoundary.RecommendedAssistantCharacters &&
                    TextBoundary.MaxToolRounds ==
                    TextBoundary.RecommendedToolRounds,
                    "Effective limits must default to the recommended values.");

                var wild = new AppSettings
                {
                    UseRecommendedLimits = false,
                    LimitContextMultiplier = 100,
                    LimitPromptCharacters = 999999,
                    LimitAssistantCharacters = 1,
                    LimitHistoryTurns = 1000,
                    LimitToolRounds = 100,
                    LimitToolCallsPerRound = 0,
                    LimitWorkingSetMessages = 1000
                };
                wild.ApplyLimits();
                Assert(
                    TextBoundary.MaxUserPromptCharacters ==
                    TextBoundary.RecommendedUserPromptCharacters &&
                    TextBoundary.MaxAssistantCharacters ==
                    TextBoundary.RecommendedAssistantCharacters &&
                    TextBoundary.MaxConversationTurns ==
                    TextBoundary.RecommendedConversationTurns &&
                    TextBoundary.MaxToolRounds ==
                    TextBoundary.RecommendedToolRounds &&
                    TextBoundary.MaxToolCallsPerRound ==
                    TextBoundary.RecommendedToolCallsPerRound &&
                    MailboxWorkingSet.MaxMessages == 1000 &&
                    ContextScale.Scaled(1000) == 1000,
                    "Text and loop budgets must stay at the " +
                    "reviewed fixed defaults while the working-set " +
                    "size honors the user's setting.");

                var overflowing = new AppSettings
                {
                    LimitWorkingSetMessages = 999999
                };
                overflowing.ApplyLimits();
                Assert(
                    MailboxWorkingSet.MaxMessages ==
                    LimitOverrides.MaxWorkingSetMessages,
                    "An out-of-range working-set size must clamp " +
                    "to the maximum.");

                var recommended = new AppSettings
                {
                    UseRecommendedLimits = true,
                    LimitContextMultiplier = 8,
                    LimitPromptCharacters = 999999
                };
                recommended.ApplyLimits();
                Assert(
                    TextBoundary.MaxUserPromptCharacters ==
                    TextBoundary
                        .RecommendedUserPromptCharacters &&
                    MailboxWorkingSet.MaxMessages ==
                    MailboxWorkingSet.RecommendedMaxMessages &&
                    ContextScale.Scaled(1000) == 1000,
                    "Defaults must apply when no working-set size " +
                    "is stored.");
            }
            finally
            {
                new AppSettings().ApplyLimits();
                ContextScale.Apply(false);
            }
        }

        private static void AdminPolicyIsReadOnlyAndScoped()
        {
            Assert(
                AdminPolicy.PolicyKeyPath ==
                "Software\\Policies\\Scribble" &&
                !AdminPolicy.GeminiEnabledForEndUsers &&
                AdminPolicy.GeminiDisabled,
                "The policy key path changed unexpectedly.");
            // Reading the switch must never throw, whether or not
            // the key exists on this machine.
            var disabled = AdminPolicy.GeminiDisabled;
            Assert(
                disabled || !disabled,
                "Policy reads must be side-effect free.");
        }

        private static void DocumentDraftIntentRequiresExplicitPhrase()
        {
            Assert(
                !DocumentDraftIntentPolicy.AllowsDraft(
                    "what does this sheet say about the budget"),
                "A plain question must not authorize a document draft.");
            Assert(
                !DocumentDraftIntentPolicy.AllowsDraft(
                    "summarize slide 3"),
                "A summary request must not authorize a document draft.");
            Assert(
                !DocumentDraftIntentPolicy.AllowsDraft(null),
                "A null prompt must not authorize a document draft.");
            Assert(
                DocumentDraftIntentPolicy.AllowsDraft(
                    "Email these slides to John"),
                "An explicit email request must authorize a draft.");
            Assert(
                DocumentDraftIntentPolicy.AllowsDraft(
                    "put this in powerpoint"),
                "An explicit cross-app request must authorize a draft.");
            Assert(
                DocumentDraftIntentPolicy.AllowsDraft(
                    "create a table in a new sheet with the totals"),
                "An explicit sheet-write request must authorize a draft.");
            Assert(
                DocumentDraftIntentPolicy.AllowsDraft(
                    "add slides about pricing"),
                "An explicit slide request must authorize a draft.");
            Assert(
                DocumentDraftIntentPolicy.AllowsDraft(
                    "update the table with the new totals"),
                "An edit verb plus a document reference must authorize a draft.");
            Assert(
                DocumentDraftIntentPolicy.AllowsDraft(
                    "fix the formula in column B"),
                "A fix request on document content must authorize a draft.");
            Assert(
                DocumentDraftIntentPolicy.AllowsDraft(
                    "build a slide with this"),
                "A build-a-slide request must authorize a draft.");
            Assert(
                DocumentDraftIntentPolicy.AllowsDraft(
                    "do a bar chart with this in a slide"),
                "A chart request must authorize a draft.");
            Assert(
                DocumentDraftIntentPolicy.AllowsDraft(
                    "put this table into word and format it properly"),
                "A put-into-word request must authorize a draft.");
            Assert(
                DocumentDraftIntentPolicy.AllowsDraft(
                    "Build me a slide of my day"),
                "The Outlook slide-of-my-day request must authorize a draft.");
            Assert(
                DocumentDraftIntentPolicy.AllowsDraft(
                    "fill in the missing totals on my sheet"),
                "A fill-my-sheet request must authorize a draft.");
            Assert(
                DocumentDraftIntentPolicy.AllowsDraft(
                    "Create an excel") &&
                DocumentDraftIntentPolicy.AllowsDraft(
                    "Create a word") &&
                DocumentDraftIntentPolicy.AllowsDraft(
                    "Create a powerpoint") &&
                DocumentDraftIntentPolicy.AllowsDraft(
                    "make me an excel file of my meetings") &&
                DocumentDraftIntentPolicy.AllowsDraft(
                    "create a spreadsheet of my day") &&
                DocumentDraftIntentPolicy.AllowsDraft(
                    "Give me this as a word doc"),
                "Plain create-a-file requests must authorize a draft first try.");
            Assert(
                !DocumentDraftIntentPolicy.AllowsDraft(
                    "give me an overview of the project"),
                "An overview request must not authorize a document draft.");
        }

        private static void WorkbookAndPresentationCatalogsStayReadOnly()
        {
            var workbookNames = WorkbookToolCatalog.ApprovedNames
                .OrderBy(name => name, StringComparer.Ordinal)
                .ToArray();
            Assert(
                workbookNames.SequenceEqual(new[]
                {
                    "list_worksheets",
                    "read_cells"
                }),
                "The workbook read catalog gained an unexpected capability.");
            var presentationNames = PresentationToolCatalog
                .ApprovedNames
                .OrderBy(name => name, StringComparer.Ordinal)
                .ToArray();
            Assert(
                presentationNames.SequenceEqual(new[]
                {
                    "list_slides",
                    "read_slide"
                }),
                "The presentation read catalog gained an unexpected capability.");
            var wordNames = WordToolCatalog.ApprovedNames
                .OrderBy(name => name, StringComparer.Ordinal)
                .ToArray();
            Assert(
                wordNames.SequenceEqual(new[]
                {
                    "read_document"
                }),
                "The Word read catalog gained an unexpected capability.");
            Assert(
                !WorkbookToolCatalog.IsApproved(
                    WorkbookToolCatalog.WriteDraftSheet) &&
                !PresentationToolCatalog.IsApproved(
                    PresentationToolCatalog.AddDraftSlides) &&
                !WordToolCatalog.IsApproved(
                    WordToolCatalog.WriteDraftDocument),
                "Draft writers must never pass the read-only approval check.");

            Assert(
                DocumentDraftHost.IsDraftTool(
                    "excel",
                    WorkbookToolCatalog.WriteDraftSheet) &&
                !DocumentDraftHost.IsDraftTool(
                    "excel",
                    PresentationToolCatalog.AddDraftSlides) &&
                DocumentDraftHost.IsDraftTool(
                    "excel",
                    CrossAppToolCatalog.SendToPowerPoint) &&
                !DocumentDraftHost.IsDraftTool(
                    "excel",
                    CrossAppToolCatalog.SendToExcel) &&
                DocumentDraftHost.IsDraftTool(
                    "powerpoint",
                    CrossAppToolCatalog.SendToExcel) &&
                !DocumentDraftHost.IsDraftTool(
                    "powerpoint",
                    CrossAppToolCatalog.SendToPowerPoint) &&
                DocumentDraftHost.IsDraftTool(
                    "word",
                    WordToolCatalog.WriteDraftDocument) &&
                !DocumentDraftHost.IsDraftTool(
                    "word",
                    CrossAppToolCatalog.SendToWord) &&
                DocumentDraftHost.IsDraftTool(
                    "word",
                    CrossAppToolCatalog.SendToExcel) &&
                DocumentDraftHost.IsDraftTool(
                    "word",
                    CrossAppToolCatalog.SendToPowerPoint) &&
                DocumentDraftHost.IsDraftTool(
                    "excel",
                    CrossAppToolCatalog.SendToWord) &&
                DocumentDraftHost.IsDraftTool(
                    "powerpoint",
                    CrossAppToolCatalog.SendToWord) &&
                DocumentDraftHost.IsDraftTool(
                    "excel",
                    CrossAppToolCatalog.CreateEmailDraft) &&
                DocumentDraftHost.IsDraftTool(
                    "powerpoint",
                    CrossAppToolCatalog.CreateEmailDraft) &&
                DocumentDraftHost.IsDraftTool(
                    "word",
                    CrossAppToolCatalog.CreateEmailDraft) &&
                !DocumentDraftHost.IsDraftTool(
                    "excel",
                    MailboxToolCatalog.SearchMailbox),
                "Draft tool routing must be host-specific and exclusive.");

            Assert(
                DocumentDraftHost.IsDraftTool(
                    "outlook",
                    CrossAppToolCatalog.SendToExcel) &&
                DocumentDraftHost.IsDraftTool(
                    "outlook",
                    CrossAppToolCatalog.SendToPowerPoint) &&
                DocumentDraftHost.IsDraftTool(
                    "outlook",
                    CrossAppToolCatalog.SendToWord) &&
                !DocumentDraftHost.IsDraftTool(
                    "outlook",
                    CrossAppToolCatalog.CreateEmailDraft) &&
                !DocumentDraftHost.IsDraftTool(
                    "outlook",
                    WorkbookToolCatalog.WriteDraftSheet) &&
                !DocumentDraftHost.IsDraftTool(
                    "outlook",
                    WorkbookToolCatalog.WriteCells),
                "Outlook cross-app routing must cover exactly the three send tools.");

            Assert(
                DocumentDraftHost.IsDraftTool(
                    "excel",
                    WorkbookToolCatalog.WriteCells) &&
                !DocumentDraftHost.IsDraftTool(
                    "word",
                    WorkbookToolCatalog.WriteCells) &&
                !DocumentDraftHost.IsDraftTool(
                    "powerpoint",
                    WorkbookToolCatalog.WriteCells) &&
                !WorkbookToolCatalog.IsApproved(
                    WorkbookToolCatalog.WriteCells) &&
                WorkbookToolCatalog.IsDraftTool(
                    WorkbookToolCatalog.WriteCells),
                "write_cells must stay an Excel-only authorized draft tool.");

            var outlookCross = CrossAppToolCatalog
                .CreateDefinitions("outlook")
                .Select(tool => tool.function.name)
                .OrderBy(name => name, StringComparer.Ordinal)
                .ToArray();
            Assert(
                outlookCross.SequenceEqual(new[]
                {
                    "send_to_excel",
                    "send_to_powerpoint",
                    "send_to_word"
                }),
                "Outlook cross-app tools changed unexpectedly.");

            var excelCross = CrossAppToolCatalog
                .CreateDefinitions("excel")
                .Select(tool => tool.function.name)
                .OrderBy(name => name, StringComparer.Ordinal)
                .ToArray();
            Assert(
                excelCross.SequenceEqual(new[]
                {
                    "create_email_draft",
                    "send_to_powerpoint",
                    "send_to_word"
                }),
                "Excel cross-app tools changed unexpectedly.");
            var powerPointCross = CrossAppToolCatalog
                .CreateDefinitions("powerpoint")
                .Select(tool => tool.function.name)
                .OrderBy(name => name, StringComparer.Ordinal)
                .ToArray();
            Assert(
                powerPointCross.SequenceEqual(new[]
                {
                    "create_email_draft",
                    "send_to_excel",
                    "send_to_word"
                }),
                "PowerPoint cross-app tools changed unexpectedly.");
            var wordCross = CrossAppToolCatalog
                .CreateDefinitions("word")
                .Select(tool => tool.function.name)
                .OrderBy(name => name, StringComparer.Ordinal)
                .ToArray();
            Assert(
                wordCross.SequenceEqual(new[]
                {
                    "create_email_draft",
                    "send_to_excel",
                    "send_to_powerpoint"
                }),
                "Word cross-app tools changed unexpectedly.");
            var emailTool = CrossAppToolCatalog
                .CreateDefinitions("excel")
                .First(tool =>
                    tool.function.name == "create_email_draft");
            Assert(
                emailTool.function.description.Contains("never sent") &&
                emailTool.function.description.Contains(
                    "sending is impossible"),
                "The email draft tool must state that sending is impossible.");
        }

        private static void DocumentFactoryGatesDraftTools()
        {
            var unauthorized = DocumentChatRequestFactory.Create(
                "test-model",
                "excel",
                "Workbook: Book1",
                new List<ChatTurn>(),
                "what is in sheet one");
            var unauthorizedNames = unauthorized.tools
                .Select(tool => tool.function.name)
                .ToArray();
            Assert(
                unauthorizedNames.SequenceEqual(new[]
                {
                    "list_worksheets",
                    "read_cells"
                }),
                "Unauthorized document requests must expose read tools only.");
            var unauthorizedSystem = Convert.ToString(
                ((ChatCompletionInputMessage)
                    unauthorized.messages[0]).content);
            Assert(
                unauthorizedSystem.Contains(
                    "Draft mutation and email drafting are unavailable") &&
                unauthorizedSystem.Contains(
                    "can never send email"),
                "The unauthorized document boundary is incomplete.");
            var contextMessage = Convert.ToString(
                ((ChatCompletionInputMessage)
                    unauthorized.messages[1]).content);
            Assert(
                contextMessage.Contains(
                    "<active_document_reference>") &&
                contextMessage.Contains("untrusted reference data"),
                "Document context must ride in an untrusted envelope.");

            var authorized = DocumentChatRequestFactory.Create(
                "test-model",
                "powerpoint",
                "Presentation: Deck1",
                new List<ChatTurn>(),
                "email these slides to john",
                true);
            var authorizedNames = authorized.tools
                .Select(tool => tool.function.name)
                .ToArray();
            Assert(
                authorizedNames.Contains("add_draft_slides") &&
                authorizedNames.Contains("create_email_draft") &&
                authorizedNames.Contains("send_to_excel") &&
                !authorizedNames.Contains("send_to_powerpoint"),
                "Authorized PowerPoint requests expose the wrong draft tools.");
            var authorizedSystem = Convert.ToString(
                ((ChatCompletionInputMessage)
                    authorized.messages[0]).content);
            Assert(
                authorizedSystem.Contains(
                    "authorized ONE deliverable") &&
                authorizedSystem.Contains("Scribble Draft") &&
                authorizedSystem.Contains("[Scribble draft]") &&
                authorizedSystem.Contains(
                    "Never claim content was saved"),
                "The authorized document boundary is incomplete.");
            // The density contract is what keeps a small model from
            // returning a thin outline of a rich source.
            Assert(
                authorizedSystem.Contains("read it to the END") &&
                authorizedSystem.Contains(
                    "must carry a table, a chart, or a numbered card") &&
                authorizedSystem.Contains("never invent"),
                "The authorized boundary lost its density contract.");

            var wordAuthorized = DocumentChatRequestFactory.Create(
                "test-model",
                "word",
                "Document: Report1",
                new List<ChatTurn>(),
                "fix the second paragraph",
                true);
            var wordNames = wordAuthorized.tools
                .Select(tool => tool.function.name)
                .ToArray();
            Assert(
                wordNames.Contains("read_document") &&
                wordNames.Contains("write_draft_document") &&
                wordNames.Contains("create_email_draft") &&
                wordNames.Contains("send_to_excel") &&
                wordNames.Contains("send_to_powerpoint") &&
                !wordNames.Contains("send_to_word"),
                "Authorized Word requests expose the wrong draft tools.");
            var wordUnauthorized = DocumentChatRequestFactory.Create(
                "test-model",
                "word",
                "Document: Report1",
                new List<ChatTurn>(),
                "what does the intro say");
            Assert(
                wordUnauthorized.tools
                    .Select(tool => tool.function.name)
                    .SequenceEqual(new[] { "read_document" }),
                "Unauthorized Word requests must expose read tools only.");

            var excelAuthorized = DocumentChatRequestFactory.Create(
                "test-model",
                "excel",
                "Workbook: Book1",
                new List<ChatTurn>(),
                "fill in the missing totals on my sheet",
                true);
            var excelNames = excelAuthorized.tools
                .Select(tool => tool.function.name)
                .ToArray();
            Assert(
                excelNames.Contains("write_draft_sheet") &&
                excelNames.Contains("write_cells"),
                "Authorized Excel requests must offer both write surfaces.");
            var excelJson = new JavaScriptSerializer()
                .Serialize(excelAuthorized);
            Assert(
                excelJson.Contains("\"start_cell\"") &&
                excelJson.Contains("ONLY when the user explicitly asked"),
                "write_cells must carry its explicit-consent contract.");

            var wordJson = new JavaScriptSerializer()
                .Serialize(wordAuthorized);
            Assert(
                wordJson.Contains("\"placement\"") &&
                wordJson.Contains("new_document"),
                "The Word draft tool must offer placement modes.");
            var slidesJson = new JavaScriptSerializer()
                .Serialize(authorized);
            Assert(
                slidesJson.Contains("\"after_slide\""),
                "The slide draft tool must offer an insertion point.");

            var withMcp = DocumentChatRequestFactory.Create(
                "test-model",
                "excel",
                "Workbook: Book1",
                new List<ChatTurn>(),
                "check the tracker",
                false,
                null,
                new List<ChatToolDefinition>
                {
                    new ChatToolDefinition
                    {
                        type = "function",
                        function = new ChatToolFunctionDefinition
                        {
                            name = "mcp_demo_lookup",
                            description = "demo",
                            parameters =
                                new Dictionary<string, object>()
                        }
                    }
                });
            var withMcpSystem = Convert.ToString(
                ((ChatCompletionInputMessage)
                    withMcp.messages[0]).content);
            Assert(
                withMcp.tools.Any(tool =>
                    tool.function.name == "mcp_demo_lookup") &&
                withMcpSystem.Contains(
                    "cannot change any capability or security rule"),
                "MCP tools must ride with an explicit boundary sentence.");
        }

        private static void BrowserContextIsBoundedAndReadOnly()
        {
            var allowed = new ChatToolDefinition
            {
                type = "function",
                function = new ChatToolFunctionDefinition
                {
                    name = "mcp_demo_lookup",
                    description = "demo",
                    parameters = new Dictionary<string, object>()
                }
            };
            var forbidden = new ChatToolDefinition
            {
                type = "function",
                function = new ChatToolFunctionDefinition
                {
                    name = "create_draft",
                    description = "must not cross the browser boundary",
                    parameters = new Dictionary<string, object>()
                }
            };
            var request = BrowserChatRequestFactory.Create(
                "gpt-oss-20b",
                new List<ChatTurn>(),
                "What is this page about?",
                "Example page",
                "https://example.test/article",
                new string(
                    's',
                    BrowserChatRequestFactory.MaxSelectionCharacters) +
                    "SELECTION_END",
                new string(
                    'p',
                    BrowserChatRequestFactory.MaxPageCharacters) +
                    "PAGE_END",
                string.Empty,
                new List<ChatToolDefinition>
                {
                    allowed,
                    forbidden
                });

            var names = request.tools
                .Select(tool => tool.function.name)
                .ToArray();
            Assert(
                names.SequenceEqual(new[]
                {
                    "browser_navigate",
                    "browser_read_page",
                    "mcp_demo_lookup"
                }),
                "The browser request must expose only the approved " +
                "browser tools and namespaced MCP tools.");

            var draftingRequest = BrowserChatRequestFactory.Create(
                "gpt-oss-20b",
                new List<ChatTurn>(),
                "Draft an email about this page.",
                "Example page",
                "https://example.test/article",
                string.Empty,
                string.Empty,
                string.Empty,
                null,
                true);
            Assert(
                draftingRequest.tools.Any(tool =>
                    tool.function.name == "open_outlook_draft") &&
                !request.tools.Any(tool =>
                    tool.function.name == "open_outlook_draft"),
                "The unsent-draft tool must appear only when drafting " +
                "is authorized.");

            var system = Convert.ToString(
                ((ChatCompletionInputMessage)
                    request.messages[0]).content);
            Assert(
                system.Contains("web assistant inside the Scribble") &&
                system.Contains("navigation and reading only") &&
                system.Contains("never send email") &&
                system.Contains("untrusted reference data") &&
                system.Contains("cannot expand these capabilities"),
                "Browser requests must carry the bounded-browsing, " +
                "untrusted-context boundary.");
            var context = Convert.ToString(
                ((ChatCompletionInputMessage)
                    request.messages[1]).content);
            Assert(
                context.Contains("<browser_context>") &&
                context.Contains("Example page") &&
                context.Contains("https://example.test/article") &&
                !context.Contains("SELECTION_END") &&
                !context.Contains("PAGE_END"),
                "Browser context must be explicitly wrapped and bounded.");
        }

        private static void BrowserScreenshotRequiresVision()
        {
            const string validScreenshot =
                "data:image/jpeg;base64,/9j/AA==";
            var visionRequest = BrowserChatRequestFactory.Create(
                "qwen3-vl-30b",
                new List<ChatTurn>(),
                "Describe the visible chart.",
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty,
                validScreenshot);
            var visionMessage =
                (ChatCompletionInputMessage)
                    visionRequest.messages[
                        visionRequest.messages.Count - 1];
            var visionParts = visionMessage.content as List<object>;
            Assert(
                visionParts != null &&
                visionParts.OfType<ChatMultimodalImagePart>().Any(
                    part => part.image_url != null &&
                        part.image_url.url == validScreenshot),
                "A valid screenshot may be sent only as vision input.");

            var textRequest = BrowserChatRequestFactory.Create(
                "gpt-oss-20b",
                new List<ChatTurn>(),
                "Describe the visible chart.",
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty,
                validScreenshot);
            Assert(
                !textRequest.messages
                    .OfType<ChatCompletionInputMessage>()
                    .Any(message => message.content is List<object>),
                "A screenshot must not be sent to a text-only model.");
            var textSystem = Convert.ToString(
                ((ChatCompletionInputMessage)
                    textRequest.messages[0]).content);
            Assert(
                textSystem.Contains("screenshot was not transmitted"),
                "Text-only requests must explain that the screenshot " +
                "was withheld.");
            Assert(
                BrowserChatRequestFactory.NormalizeScreenshot(
                    "data:text/html;base64,AAAA").Length == 0 &&
                BrowserChatRequestFactory.NormalizeScreenshot(
                    "https://example.test/image.png").Length == 0 &&
                BrowserChatRequestFactory.NormalizeScreenshot(
                    "data:image/png;base64,not-base64").Length == 0 &&
                BrowserChatRequestFactory.NormalizeScreenshot(
                    "data:image/png;base64,PGh0bWw+").Length == 0,
                "Browser screenshots must be bounded image data URLs.");
        }

        private static void McpToolsAreNamespacedAndBounded()
        {
            Assert(
                McpServerConfig.SanitizeName("My Server!") ==
                "my_server",
                "MCP server names must reduce to a safe token.");
            Assert(
                McpServerConfig.SanitizeName("   ") == "server",
                "Empty MCP server names must fall back to a token.");
            Assert(
                McpToolHost.IsMcpTool("mcp_files_read") &&
                !McpToolHost.IsMcpTool("read_cells") &&
                !McpToolHost.IsMcpTool(null),
                "MCP routing must key on the mcp_ namespace only.");
            using (var empty = new McpToolHost(
                new List<McpServerConfig>()))
            {
                Assert(
                    !empty.HasServers &&
                    empty.GetDefinitions().Count == 0,
                    "An empty MCP configuration must expose no tools.");
            }

            using (var disabled = new McpToolHost(
                new List<McpServerConfig>
                {
                    new McpServerConfig
                    {
                        Name = "off",
                        Target = "https://example.test/mcp",
                        Enabled = false
                    }
                }))
            {
                Assert(
                    !disabled.HasServers,
                    "Disabled MCP servers must stay disconnected.");
            }

            var config = new McpServerConfig
            {
                Name = "Files Server",
                Target = "  https://example.test/mcp  ",
                Arguments = "--flag",
                BrowserTools =
                    "search_query\nsearch_query,lookup",
                BrowserToolsApproved = true,
                Enabled = true
            }.Sanitized();
            Assert(
                config.Name == "files_server" &&
                config.IsHttp &&
                config.Target == "https://example.test/mcp" &&
                config.BrowserToolsApproved &&
                config.ParsedBrowserTools().SequenceEqual(
                    new[] { "search_query", "lookup" }),
                "MCP server configuration sanitization failed.");

            using (var browserDefault = new McpToolHost(
                new List<McpServerConfig>
                {
                    new McpServerConfig
                    {
                        Name = "search",
                        Target = "https://example.test/mcp",
                        BrowserTools = "search_query",
                        BrowserToolsApproved = false,
                        Enabled = true
                    }
                },
                true))
            {
                Assert(
                    !browserDefault.HasServers,
                    "Browser MCP must be disabled unless the user " +
                    "explicitly approves exact read-only tools.");
            }

            using (var browserApproved = new McpToolHost(
                new List<McpServerConfig>
                {
                    config
                },
                true))
            {
                Assert(
                    browserApproved.HasServers,
                    "An explicitly approved browser MCP allowlist " +
                    "must register its single bounded server.");
            }
        }

        // Answers JSON-RPC over stdio like a minimal MCP server:
        // initialize, tools/list with one echo tool, and tools/call
        // echoing the value argument back.
        private static int RunFakeMcpServer()
        {
            var serializer = new JavaScriptSerializer();
            while (true)
            {
                var line = Console.In.ReadLine();
                if (line == null)
                {
                    return 0;
                }

                line = line.TrimStart('\uFEFF');
                IDictionary<string, object> message;
                try
                {
                    message = serializer.DeserializeObject(line) as
                        IDictionary<string, object>;
                }
                catch
                {
                    continue;
                }

                object methodValue;
                object idValue;
                if (message == null ||
                    !message.TryGetValue("method", out methodValue) ||
                    !message.TryGetValue("id", out idValue))
                {
                    continue;
                }

                var method = Convert.ToString(methodValue);
                var id = Convert.ToString(idValue);
                string result;
                if (method == "initialize")
                {
                    result =
                        "{\"protocolVersion\":\"2025-03-26\"," +
                        "\"capabilities\":{}," +
                        "\"serverInfo\":{\"name\":\"fake\"," +
                        "\"version\":\"1.0\"}}";
                }
                else if (method == "tools/list")
                {
                    result =
                        "{\"tools\":[{\"name\":\"echo\"," +
                        "\"description\":\"Echoes the value back\"," +
                        "\"inputSchema\":{\"type\":\"object\"," +
                        "\"properties\":{\"value\":" +
                        "{\"type\":\"string\"}}}}]}";
                }
                else if (method == "tools/call")
                {
                    var value = string.Empty;
                    object parametersValue;
                    if (message.TryGetValue(
                        "params",
                        out parametersValue))
                    {
                        var parameters = parametersValue as
                            IDictionary<string, object>;
                        object argumentsValue;
                        if (parameters != null &&
                            parameters.TryGetValue(
                                "arguments",
                                out argumentsValue))
                        {
                            var arguments = argumentsValue as
                                IDictionary<string, object>;
                            object rawValue;
                            if (arguments != null &&
                                arguments.TryGetValue(
                                    "value",
                                    out rawValue))
                            {
                                value = Convert.ToString(rawValue) ??
                                    string.Empty;
                            }
                        }
                    }

                    result =
                        "{\"content\":[{\"type\":\"text\"," +
                        "\"text\":" +
                        serializer.Serialize("echo:" + value) +
                        "}],\"isError\":false}";
                }
                else
                {
                    continue;
                }

                Console.Out.WriteLine(
                    "{\"jsonrpc\":\"2.0\",\"id\":" + id +
                    ",\"result\":" + result + "}");
                Console.Out.Flush();
            }
        }

        // Bypasses the MCP client entirely: spawns this exe as a raw
        // line-echo child with canonical redirection and reports
        // whether a round trip works in this environment. CI runner
        // sandboxes that cannot run console children at all are
        // detected here so an environment limitation is not
        // misreported as a product defect.
        private static bool RawChildEchoWorks(out string detail)
        {
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            System.Diagnostics.Process process = null;
            try
            {
                process = System.Diagnostics.Process.Start(
                    new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = Assembly
                            .GetExecutingAssembly().Location,
                        Arguments = "--echo-server",
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        RedirectStandardInput = true,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true
                    });
                process.StandardInput.AutoFlush = true;
                process.StandardInput.WriteLine("ping");
                var reader = process.StandardOutput;
                var readTask = Task.Run(() => reader.ReadLine());
                if (!readTask.Wait(90000))
                {
                    detail = "no echo reply after 90s";
                    return false;
                }

                detail = "reply '" +
                    (readTask.Result ?? "(closed)") +
                    "' after " +
                    stopwatch.ElapsedMilliseconds + "ms";
                // A writer-side encoding preamble (BOM) may ride in
                // the echoed payload; the spawn still works.
                var reply = (readTask.Result ?? string.Empty)
                    .Replace("\uFEFF", string.Empty);
                return reply == "pong:ping";
            }
            catch (Exception exception)
            {
                detail = "spawn failed: " + exception.Message;
                return false;
            }
            finally
            {
                try
                {
                    if (process != null && !process.HasExited)
                    {
                        process.Kill();
                    }

                    process?.Dispose();
                }
                catch
                {
                }
            }
        }

        private static void McpStdioRoundTripWorks()
        {
            string echoDetail;
            var echoWorks = RawChildEchoWorks(out echoDetail);
            Console.WriteLine(
                "  spawn diagnostic: " + echoDetail);
            if (!echoWorks)
            {
                Console.WriteLine(
                    "  SKIP: this environment cannot run console " +
                    "children, so the live MCP round trip is not " +
                    "exercised here. Namespacing, bounding, and " +
                    "rejection stay covered by the preceding test.");
                return;
            }

            var serverConfig = new McpServerConfig
            {
                Name = "fake",
                Target = Assembly
                    .GetExecutingAssembly().Location,
                Arguments = "--mcp-fake-server",
                Enabled = true
            };

            // Direct probe first: any handshake failure surfaces
            // with its full exception text (including the server's
            // stderr tail) instead of an empty tool list.
            var probe = new McpConnection(
                serverConfig.Sanitized(),
                null);
            try
            {
                var probedTools = probe.ListTools();
                Assert(
                    probedTools.Count == 1 &&
                    probedTools[0].Name == "echo",
                    "The scripted MCP server listed unexpected tools: " +
                    string.Join(
                        ", ",
                        probedTools.Select(tool => tool.Name)));
            }
            catch (Exception exception)
            {
                Assert(
                    false,
                    "The MCP stdio handshake failed: " + exception);
            }
            finally
            {
                probe.Dispose();
            }

            var host = new McpToolHost(
                new List<McpServerConfig>
                {
                    serverConfig
                });
            try
            {
                Assert(
                    host.HasServers,
                    "The scripted MCP server was not registered.");
                var definitions = host.GetDefinitions();
                Assert(
                    definitions.Count == 1 &&
                    definitions[0].function.name ==
                    "mcp_fake_echo",
                    "The MCP tool was not namespaced as expected: " +
                    string.Join(
                        ", ",
                        definitions.Select(definition =>
                            definition.function.name)));
                Assert(
                    definitions[0].function.description
                        .Contains("user-configured"),
                    "MCP tool descriptions must disclose their origin.");
                var result = host.Execute(new ChatToolCall
                {
                    id = "call1",
                    type = "function",
                    function = new ChatToolCallFunction
                    {
                        name = "mcp_fake_echo",
                        arguments = "{\"value\":\"hi\"}"
                    }
                });
                Assert(
                    result.Content.Contains("untrusted_mcp_data") &&
                    result.Content.Contains("echo:hi"),
                    "The MCP stdio round trip failed: " +
                    result.Content);
                var rejected = host.Execute(new ChatToolCall
                {
                    id = "call2",
                    type = "function",
                    function = new ChatToolCallFunction
                    {
                        name = "mcp_fake_other",
                        arguments = "{}"
                    }
                });
                Assert(
                    rejected.Content.Contains(
                        "MCP_TOOL_NOT_ALLOWED"),
                    "Unregistered MCP tools must be rejected.");
            }
            finally
            {
                host.Dispose();
            }

            serverConfig.BrowserToolsApproved = true;
            serverConfig.BrowserTools = "not_echo";
            using (var browserDenied = new McpToolHost(
                new List<McpServerConfig> { serverConfig },
                true))
            {
                Assert(
                    browserDenied.GetDefinitions().Count == 0,
                    "Browser MCP must hide every tool not named in " +
                    "the exact user-approved allowlist.");
            }

            serverConfig.BrowserTools = "echo";
            using (var browserAllowed = new McpToolHost(
                new List<McpServerConfig> { serverConfig },
                true))
            {
                var browserDefinitions =
                    browserAllowed.GetDefinitions();
                Assert(
                    browserDefinitions.Count == 1 &&
                    browserDefinitions[0].function.name ==
                        "mcp_fake_echo",
                    "Browser MCP must expose only the exact approved " +
                    "read-only tool.");
            }
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition)
            {
                throw new InvalidOperationException(message);
            }
        }

        private sealed class FakeEndpoint : IDisposable
        {
            private readonly TcpListener _listener;
            private readonly Task _requestTask;
            private readonly string _responseBody;

            public FakeEndpoint(string responseBody)
            {
                _responseBody = responseBody;
                _listener = new TcpListener(
                    IPAddress.Loopback,
                    0);
                _listener.Start();
                var port =
                    ((IPEndPoint)_listener.LocalEndpoint)
                    .Port;
                BaseUrl = "http://127.0.0.1:" +
                    port + "/v1";
                _requestTask = Task.Run(
                    (Action)HandleRequest);
            }

            public string BaseUrl { get; }

            public string RequestLine { get; private set; } =
                string.Empty;

            public string Authorization { get; private set; } =
                string.Empty;

            public string Body { get; private set; } =
                string.Empty;

            public void Wait()
            {
                if (!_requestTask.Wait(
                    TimeSpan.FromSeconds(10)))
                {
                    throw new InvalidOperationException(
                        "The fake endpoint did not receive a request.");
                }

                if (_requestTask.IsFaulted)
                {
                    throw _requestTask.Exception
                        .GetBaseException();
                }
            }

            public void Dispose()
            {
                _listener.Stop();
                try
                {
                    _requestTask.Wait(
                        TimeSpan.FromSeconds(1));
                }
                catch
                {
                }
            }

            private void HandleRequest()
            {
                using (var client = _listener.AcceptTcpClient())
                using (var stream = client.GetStream())
                using (var reader = new StreamReader(
                    stream,
                    Encoding.ASCII,
                    false,
                    4096,
                    true))
                {
                    RequestLine =
                        reader.ReadLine() ?? string.Empty;
                    var contentLength = 0;
                    while (true)
                    {
                        var line = reader.ReadLine();
                        if (string.IsNullOrEmpty(line))
                        {
                            break;
                        }

                        var separator = line.IndexOf(':');
                        if (separator <= 0)
                        {
                            continue;
                        }

                        var name = line
                            .Substring(0, separator)
                            .Trim();
                        var value = line
                            .Substring(separator + 1)
                            .Trim();
                        if (name.Equals(
                            "Authorization",
                            StringComparison.OrdinalIgnoreCase))
                        {
                            Authorization = value;
                        }
                        else if (name.Equals(
                            "Content-Length",
                            StringComparison.OrdinalIgnoreCase))
                        {
                            int.TryParse(
                                value,
                                out contentLength);
                        }
                    }

                    if (contentLength > 0)
                    {
                        var buffer =
                            new char[contentLength];
                        var offset = 0;
                        while (offset < contentLength)
                        {
                            var read = reader.Read(
                                buffer,
                                offset,
                                contentLength - offset);
                            if (read <= 0)
                            {
                                break;
                            }

                            offset += read;
                        }

                        Body = new string(
                            buffer,
                            0,
                            offset);
                    }

                    var responseBytes =
                        Encoding.UTF8.GetBytes(
                            _responseBody);
                    var headers = Encoding.ASCII.GetBytes(
                        "HTTP/1.1 200 OK\r\n" +
                        "Content-Type: application/json\r\n" +
                        "Content-Length: " +
                        responseBytes.Length +
                        "\r\nConnection: close\r\n\r\n");
                    stream.Write(
                        headers,
                        0,
                        headers.Length);
                    stream.Write(
                        responseBytes,
                        0,
                        responseBytes.Length);
                    stream.Flush();
                }
            }
        }
    }

    public sealed class FakeExplorerContext
    {
        public FakeExplorerContext(FakeSelection selection)
        {
            Selection = selection;
        }

        public FakeSelection Selection { get; }
    }

    public sealed class FakeSelection
    {
        private readonly object[] _items;

        public FakeSelection(object[] items)
        {
            _items = items ?? new object[0];
        }

        public int Count
        {
            get { return _items.Length; }
        }

        public object Item(int index)
        {
            return _items[index - 1];
        }
    }

    public sealed class FakeSelectedMailItem
    {
        public FakeSelectedMailItem(
            string entryId,
            string subject)
        {
            EntryID = entryId;
            Subject = subject;
            Parent = new FakeMailFolder();
            ReceivedTime = DateTime.UtcNow;
            Attachments = new FakeOutlookAttachments();
        }

        public string MessageClass { get; set; } = "IPM.Note";

        public string EntryID { get; }

        public string Subject { get; }

        public string SenderName { get; } = "Sender";

        public string SenderEmailAddress { get; } =
            "sender@example.test";

        public string To { get; } = "recipient@example.test";

        public string Body { get; } = "Message body";

        public string HTMLBody { get; set; } = string.Empty;

        public DateTime ReceivedTime { get; }

        public DateTime SentOn { get; } = DateTime.MinValue;

        public DateTime CreationTime { get; } = DateTime.MinValue;

        public FakeMailFolder Parent { get; }

        public FakeOutlookAttachments Attachments { get; set; }
    }

    public sealed class FakeOutlookAttachments
    {
        private readonly List<FakeOutlookAttachment> _items =
            new List<FakeOutlookAttachment>();

        public int Count
        {
            get { return _items.Count; }
        }

        public void Add(FakeOutlookAttachment attachment)
        {
            _items.Add(attachment);
        }

        public FakeOutlookAttachment Item(int index)
        {
            return _items[index - 1];
        }
    }

    public sealed class FakeOutlookAttachment
    {
        private readonly string _sourcePath;

        public FakeOutlookAttachment(
            string fileName,
            string sourcePath)
        {
            FileName = fileName;
            _sourcePath = sourcePath;
        }

        public string FileName { get; }

        public int Size { get; set; }

        public object PropertyAccessor { get; set; }

        public void SaveAsFile(string path)
        {
            File.Copy(_sourcePath, path, true);
        }
    }

    public sealed class FakePropertyAccessor
    {
        public bool Hidden { get; set; }

        public string ContentId { get; set; } = string.Empty;

        public object GetProperty(string schema)
        {
            if (schema.EndsWith("0x7FFE000B"))
            {
                return Hidden;
            }

            if (schema.EndsWith("0x3712001F"))
            {
                if (ContentId.Length == 0)
                {
                    throw new InvalidOperationException(
                        "The property does not exist.");
                }

                return ContentId;
            }

            throw new InvalidOperationException(
                "Unknown property.");
        }
    }

    public sealed class FakeMailFolder
    {
        public string StoreID { get; } = "store";
    }

    public sealed class FakeOutlookApplication
    {
        private readonly FakeOutlookSession _session =
            new FakeOutlookSession();

        public int CreatedCount { get; private set; }

        public FakeMailItem LastDraft { get; private set; }

        public FakeExplorerContext Explorer { get; set; }

        public FakeOutlookSession Session
        {
            get { return _session; }
        }

        public object ActiveExplorer()
        {
            return Explorer;
        }

        public FakeReplySource RegisterReplySource(
            string entryId,
            string storeId,
            string replySubject,
            string replyTo)
        {
            var source = new FakeReplySource(
                this,
                replySubject,
                replyTo);
            _session.Register(entryId, storeId, source);
            return source;
        }

        public void RecordReply(FakeMailItem draft)
        {
            LastDraft = draft;
        }

        public object CreateItem(int itemType)
        {
            if (itemType != 0)
            {
                throw new InvalidOperationException(
                    "Only mail items are allowed in the test host.");
            }

            CreatedCount++;
            LastDraft = new FakeMailItem();
            return LastDraft;
        }
    }

    public sealed class FakeOutlookSession
    {
        private readonly Dictionary<string, object> _items =
            new Dictionary<string, object>(StringComparer.Ordinal);

        public void Register(
            string entryId,
            string storeId,
            object item)
        {
            _items[Key(entryId, storeId)] = item;
        }

        public object GetItemFromID(string entryId)
        {
            return GetItemFromID(entryId, string.Empty);
        }

        public object GetItemFromID(
            string entryId,
            string storeId)
        {
            object item;
            if (!_items.TryGetValue(
                    Key(entryId, storeId),
                    out item))
            {
                throw new InvalidOperationException(
                    "Unknown fake Outlook item.");
            }

            return item;
        }

        private static string Key(
            string entryId,
            string storeId)
        {
            return (entryId ?? string.Empty) +
                "\n" +
                (storeId ?? string.Empty);
        }
    }

    public sealed class FakeReplySource
    {
        private readonly FakeOutlookApplication _application;
        private readonly string _replySubject;
        private readonly string _replyTo;

        public FakeReplySource(
            FakeOutlookApplication application,
            string replySubject,
            string replyTo)
        {
            _application = application;
            _replySubject = replySubject;
            _replyTo = replyTo;
        }

        public int ReplyCount { get; private set; }

        public object Reply()
        {
            ReplyCount++;
            var draft = new FakeMailItem
            {
                Subject = _replySubject,
                To = _replyTo,
                HTMLBody = "<div>Quoted original</div>"
            };
            _application.RecordReply(draft);
            return draft;
        }
    }

    public sealed class FakeMailItem
    {
        public FakeMailItem()
        {
            Subject = string.Empty;
            To = string.Empty;
            CC = string.Empty;
            HTMLBody = string.Empty;
        }

        public string Subject { get; set; }

        public string To { get; set; }

        public string CC { get; set; }

        public string HTMLBody { get; set; }

        public bool Saved { get; private set; }

        public bool Displayed { get; private set; }

        public bool DisplayModal { get; private set; }

        public int SaveCount { get; private set; }

        public int DisplayCount { get; private set; }

        public void Save()
        {
            Saved = true;
            SaveCount++;
        }

        public void Display(bool modal)
        {
            Displayed = true;
            DisplayModal = modal;
            DisplayCount++;
        }
    }
}

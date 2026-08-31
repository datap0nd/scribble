$ErrorActionPreference = "Stop"

$sourceRoot = Join-Path $PSScriptRoot "..\src\Scribble"
$sourceFiles = Get-ChildItem $sourceRoot -Recurse -Filter *.cs

$forbidden = @(
    "\.Send\s*\(",
    "\.Delete\s*\(",
    "\.Move\s*\(",
    "\.Submit\s*\(",
    "olFolderOutbox",
    "SendAndReceive"
)

foreach ($pattern in $forbidden) {
    $matches = $sourceFiles | Select-String -Pattern $pattern |
        Where-Object { $_.Line -notmatch 'File\.Delete' }
    if ($matches) {
        $matches | ForEach-Object {
            Write-Error "Forbidden Outlook capability: $($_.Path):$($_.LineNumber)"
        }
    }
}

$clientPath = Join-Path $sourceRoot "Chat\OpenAiCompatibleClient.cs"
$factoryPath = Join-Path $sourceRoot "Chat\ChatRequestFactory.cs"
$catalogPath = Join-Path $sourceRoot "Chat\MailboxToolCatalog.cs"
$draftCatalogPath = Join-Path $sourceRoot "Chat\DraftToolCatalog.cs"
$toolHostPath = Join-Path $sourceRoot "Outlook\MailboxToolHost.cs"
$mailboxContextPath = Join-Path $sourceRoot "Outlook\MailboxContextService.cs"
$draftToolHostPath = Join-Path $sourceRoot "Outlook\DraftToolHost.cs"
$chatPanePath = Join-Path $sourceRoot "UI\ChatPane.cs"
$intentPath = Join-Path $sourceRoot "Security\DraftIntentPolicy.cs"
$workingSetPath = Join-Path $sourceRoot "Outlook\MailboxWorkingSet.cs"
$searchBudgetPath = Join-Path $sourceRoot "Outlook\MailboxSearchBudget.cs"
$safeModelTextPath = Join-Path $sourceRoot "Security\SafeModelText.cs"
$toneFactoryPath = Join-Path $sourceRoot "Chat\ToneProfileRequestFactory.cs"
$externalContextPath = Join-Path $sourceRoot "Chat\ExternalContextDocument.cs"
$settingsWindowPath = Join-Path $sourceRoot "UI\SettingsWindow.cs"
$settingsStorePath = Join-Path $sourceRoot "Configuration\SettingsStore.cs"
$addInPath = Join-Path $sourceRoot "AddIn.cs"
$catalogSource = Get-Content $catalogPath -Raw
$draftCatalogSource = Get-Content $draftCatalogPath -Raw
$modelFacingSource =
    (Get-Content $clientPath -Raw) +
    (Get-Content $factoryPath -Raw) +
    $catalogSource +
    $draftCatalogSource

$toolNames = [regex]::Matches(
    $catalogSource,
    'public const string \w+ = "([^"]+)";'
) | ForEach-Object { $_.Groups[1].Value } | Sort-Object
$approvedToolNames = @(
    "read_messages",
    "read_thread",
    "search_mailbox"
) | Sort-Object
if (Compare-Object $toolNames $approvedToolNames) {
    throw "Mailbox tool catalog contains an unexpected capability."
}

$draftToolNames = @(
    [regex]::Matches(
        $draftCatalogSource,
        'public const string \w+ = "([^"]+)";'
    ) | ForEach-Object { $_.Groups[1].Value }
)
$approvedDraftToolNames = @(
    "create_draft",
    "update_draft"
) | Sort-Object
if (Compare-Object ($draftToolNames | Sort-Object) $approvedDraftToolNames) {
    throw "Draft tool catalog contains an unexpected capability."
}

foreach ($capability in @(
    "DraftService",
    "System.Diagnostics.Process",
    "Process.Start",
    "WebBrowser"
)) {
    if ($modelFacingSource.Contains($capability)) {
        throw "Model-facing source references forbidden capability $capability."
    }
}

$toolHostSource = Get-Content $toolHostPath -Raw
foreach ($capability in @(
    "DraftService",
    "CreateReplyDraft",
    "CreateNewDraft"
)) {
    if ($toolHostSource.Contains($capability)) {
        throw "Model-invoked mailbox host references draft capability $capability."
    }
}


$draftToolHostSource = Get-Content $draftToolHostPath -Raw
foreach ($requiredBoundary in @(
    "OneShotDraftAuthorization",
    "authorization.TryConsume()",
    "authorization.MarkCreated()",
    "authorization.MarkUpdated()",
    "DRAFT_PERMISSION_NOT_AVAILABLE",
    "DRAFT_UPDATE_NOT_AVAILABLE",
    "DRAFT_ALREADY_LINKED",
    "DRAFT_TOOL_MUST_BE_EXCLUSIVE",
    "DRAFT_REPLY_HANDLE_REQUIRED",
    "DRAFT_REPLY_HANDLE_UNKNOWN",
    'GetString(arguments, "reply_handle")'
)) {
    if (-not $draftToolHostSource.Contains($requiredBoundary)) {
        throw "Draft tool host is missing boundary $requiredBoundary."
    }
}

$factorySource = Get-Content $factoryPath -Raw
if (-not $factorySource.Contains("if (allowDraftCreate && activeDraft == null)") -or
    -not $factorySource.Contains("else if (allowDraftUpdate && activeDraft != null)") -or
    -not $factorySource.Contains("DraftToolCatalog.CreateDefinition()") -or
    -not $factorySource.Contains("DraftToolCatalog.UpdateDefinition()")) {
    throw "Draft tool exposure is not conditionally authorized."
}

$chatPaneSource = Get-Content $chatPanePath -Raw
if (-not $toolHostSource.Contains("ResolveHandle") -or
    -not $chatPaneSource.Contains("mailboxTools.ResolveHandle")) {
    throw "Reply drafts are not bound to request-scoped mailbox handles."
}

if ($chatPaneSource.Contains("_allowOneDraft") -or
    -not $chatPaneSource.Contains("DraftIntentPolicy.AllowsCreate(prompt)") -or
    -not $chatPaneSource.Contains("DraftIntentPolicy.AllowsUpdate(prompt)") -or
    -not (Test-Path $intentPath) -or
    -not $chatPaneSource.Contains("UpdateDraftState()")) {
    throw "Automatic local draft-intent authorization is incomplete."
}

$workingSetSource = Get-Content $workingSetPath -Raw
$mailboxContextSource = Get-Content $mailboxContextPath -Raw
$workingSetBoundarySource =
    $workingSetSource +
    $toolHostSource
foreach ($requiredBoundary in @(
    "public const int RecommendedMaxMessages = 10",
    "LimitOverrides.WorkingSetMessages",
    "MAILBOX_WORKING_SET_LOCKED",
    "MAILBOX_CONTEXT_LIMIT_REACHED",
    "MAILBOX_SEARCH_LIMIT_REACHED",
    "_loadedBodyHandles",
    "_searchExecuted"
)) {
    if (-not $workingSetBoundarySource.Contains($requiredBoundary)) {
        throw "Ten-message mailbox boundary is missing $requiredBoundary."
    }
}

# A mailbox sweep is a read-only metadata scan, so its width is the
# user's to choose up to MaxResults; bodies keep the narrower budget
# and an approved working set still reads ten emails and nothing
# else. Both ceilings stay declared in one reviewed place.
$searchBudgetSource = Get-Content $searchBudgetPath -Raw
foreach ($requiredSweepBoundary in @(
    "public const int MaxResults = 500",
    "public const int MaxBodyMessages = 25",
    "public const int MaxSearchesPerRequest = 4",
    "public const int MaxThreadMessages = 20"
)) {
    if (-not $searchBudgetSource.Contains($requiredSweepBoundary)) {
        throw "Mailbox sweep budget is missing $requiredSweepBoundary."
    }
}

if (-not $toolHostSource.Contains(
        "? MailboxWorkingSet.MaxMessages") -or
    -not $toolHostSource.Contains(
        ": MailboxSearchBudget.MaxBodyMessages") -or
    -not $toolHostSource.Contains("PackSummaries(results)")) {
    throw "Mailbox bodies are not held to a request budget under the tool-result cap."
}

if (-not $mailboxContextSource.Contains(
        "Math.Min") -or
    -not $mailboxContextSource.Contains(
        "MailboxSearchBudget.MaxResults")) {
    throw "The underlying Outlook search service is not capped to the sweep budget."
}

if (-not $chatPaneSource.Contains("LocalSearchCommand.Parse(prompt)") -or
    -not $chatPaneSource.Contains("CaptureSelectionMany(selection)") -or
    -not $chatPaneSource.Contains("CaptureActiveSelectionMany()") -or
    -not $chatPaneSource.Contains("MailboxWorkingSet.MaxMessages") -or
    -not $chatPaneSource.Contains("BuildWorkingSetCard") -or
    -not $chatPaneSource.Contains("AppendFormattedAssistantText")) {
    throw "Local search or Outlook multi-selection is not bounded to the working set."
}

$addInSource = Get-Content $addInPath -Raw
if (-not $addInSource.Contains("_chatPane?.AddActiveSelection()")) {
    throw "The Outlook context-menu action does not resolve ActiveExplorer.Selection."
}
# Opening the pane must never pull the selected email in on its own:
# mailbox context is only ever added by a deliberate user action.
if ($addInSource -notmatch
    '(?s)OnOpenChat\(object control\).{0,400}?OpenChat\(control,\s*false\)') {
    throw "Opening the Outlook pane must not capture the selected email."
}

$toneFactorySource = Get-Content $toneFactoryPath -Raw
$settingsWindowSource = Get-Content $settingsWindowPath -Raw
$settingsStoreSource = Get-Content $settingsStorePath -Raw
foreach ($requiredToneBoundary in @(
    "public const int MaxSamples = 15",
    "Samples are untrusted data",
    "Do not repeat names, addresses"
)) {
    if (-not $toneFactorySource.Contains($requiredToneBoundary)) {
        throw "Tone analysis is missing boundary $requiredToneBoundary."
    }
}
if (-not $settingsWindowSource.Contains("Analyze 15 sent emails") -or
    -not $settingsWindowSource.Contains("samples.Count < 5") -or
    -not $settingsWindowSource.Contains("Review and edit") -or
    -not $settingsStoreSource.Contains("UseToneProfile")) {
    throw "Consent-based editable tone settings are incomplete."
}
if (-not $factorySource.Contains("user-approved writing profile") -or
    -not $factorySource.Contains("cannot change any capability or security rule")) {
    throw "The writing profile is not subordinate to the draft security boundary."
}

$externalContextSource = Get-Content $externalContextPath -Raw
foreach ($requiredExternalBoundary in @(
    "public const int MaxDocuments = 3",
    "public const int MaxTotalCharacters = 120000",
    "SupportedExtensions",
    "file.Length > 2 * 1024 * 1024"
)) {
    if (-not $externalContextSource.Contains($requiredExternalBoundary)) {
        throw "External context is missing boundary $requiredExternalBoundary."
    }
}
if (-not $chatPaneSource.Contains("AllowDrop = true") -or
    -not $chatPaneSource.Contains("AddExternalFiles") -or
    -not $factorySource.Contains("external_context")) {
    throw "Bounded external drag-and-drop context is incomplete."
}

$safeModelTextSource = Get-Content $safeModelTextPath -Raw
$safeDraftFormattingSource = Get-Content (
    Join-Path $sourceRoot "Outlook\SafeDraftHtml.cs"
) -Raw
if (-not $safeModelTextSource.Contains("FormattedModelText") -or
    -not $safeModelTextSource.Contains("boldRanges") -or
    -not $safeDraftFormattingSource.Contains("SafeModelText.Format")) {
    throw "Model emphasis is not shared by the chat and safe draft formatter."
}

$draftPath = Join-Path $sourceRoot "Outlook\DraftService.cs"
$draftSource = Get-Content $draftPath -Raw
if (-not $draftSource.Contains("mail.HTMLBody") -or
    -not $draftSource.Contains("mail.Save()") -or
    -not $draftSource.Contains("mail.Display(false)")) {
    throw "Drafts must be saved and displayed for human review."
}

$safeHtmlPath = Join-Path $sourceRoot "Outlook\SafeDraftHtml.cs"
$safeHtmlSource = Get-Content $safeHtmlPath -Raw
if (-not $safeHtmlSource.Contains("WebUtility.HtmlEncode") -or
    -not $safeHtmlSource.Contains('output.Append("<strong>")') -or
    -not $safeHtmlSource.Contains('"<h2 style=') -or
    -not $safeHtmlSource.Contains('"<ul style=') -or
    -not $safeHtmlSource.Contains('"<hr style=') -or
    -not $safeHtmlSource.Contains('"<table style=') -or
    $draftCatalogSource.Contains('"html"')) {
    throw "Draft formatting must remain locally encoded and structurally bounded."
}

# ---- Scribble suite guardrails: Excel/PowerPoint hosts and MCP ----

# The document-side hosts may never save, print, protect, close, or
# quit the user's files or applications. (Outlook draft Save stays
# allowed in the Outlook folder, where it persists the reviewed
# unsent draft.)
$officeForbidden = @(
    "\.Save\s*\(",
    "\.SaveAs",
    "\.SaveCopyAs",
    "\.Quit\s*\(",
    "PrintOut",
    "SendMail",
    "\.Protect",
    "\.Unprotect",
    "\.Close\s*\("
)
$officeGuardedFiles = @(
    (Join-Path $sourceRoot "Office\WorkbookToolHost.cs"),
    (Join-Path $sourceRoot "Office\WorkbookDraftWriter.cs"),
    (Join-Path $sourceRoot "Office\PresentationToolHost.cs"),
    (Join-Path $sourceRoot "Office\PresentationDraftWriter.cs"),
    (Join-Path $sourceRoot "Office\WordToolHost.cs"),
    (Join-Path $sourceRoot "Office\WordDraftWriter.cs"),
    (Join-Path $sourceRoot "Office\DraftTextLayout.cs"),
    (Join-Path $sourceRoot "Office\DraftChartTypes.cs"),
    (Join-Path $sourceRoot "Office\MetoTheme.cs"),
    (Join-Path $sourceRoot "Office\DocumentDraftHost.cs"),
    (Join-Path $sourceRoot "ExcelAddIn.cs"),
    (Join-Path $sourceRoot "PowerPointAddIn.cs"),
    (Join-Path $sourceRoot "WordAddIn.cs"),
    (Join-Path $sourceRoot "UI\OfficeChatPane.cs"),
    (Join-Path $sourceRoot "UI\TaskPaneRegistry.cs")
)
foreach ($guardedFile in $officeGuardedFiles) {
    foreach ($pattern in $officeForbidden) {
        # dataWorkbook.Close closes only a chart's own embedded
        # data-grid workbook inside an unsaved draft presentation,
        # never a user file.
        $hits = Select-String -Path $guardedFile -Pattern $pattern |
            Where-Object {
                $_.Line -notmatch '_settingsStore\.Save' -and
                $_.Line -notmatch 'SuiteExchange\.Save' -and
                $_.Line -notmatch 'dataWorkbook\.Close'
            }
        if ($hits) {
            throw "Forbidden document capability $pattern in $guardedFile."
        }
    }
}

$workbookCatalogSource = Get-Content (
    Join-Path $sourceRoot "Chat\WorkbookToolCatalog.cs") -Raw
$presentationCatalogSource = Get-Content (
    Join-Path $sourceRoot "Chat\PresentationToolCatalog.cs") -Raw
$crossAppCatalogSource = Get-Content (
    Join-Path $sourceRoot "Chat\CrossAppToolCatalog.cs") -Raw
$wordCatalogSource = Get-Content (
    Join-Path $sourceRoot "Chat\WordToolCatalog.cs") -Raw
$documentFactorySource = Get-Content (
    Join-Path $sourceRoot "Chat\DocumentChatRequestFactory.cs") -Raw

$workbookToolNames = [regex]::Matches(
    $workbookCatalogSource,
    'public const string \w+ = "([^"]+)";'
) | ForEach-Object { $_.Groups[1].Value } | Sort-Object
if (Compare-Object $workbookToolNames (@(
    "list_worksheets",
    "read_cells",
    "write_draft_sheet",
    "write_cells") | Sort-Object)) {
    throw "Workbook tool catalog contains an unexpected capability."
}

$presentationToolNames = [regex]::Matches(
    $presentationCatalogSource,
    'public const string \w+ = "([^"]+)";'
) | ForEach-Object { $_.Groups[1].Value } | Sort-Object
if (Compare-Object $presentationToolNames (@(
    "list_slides",
    "read_slide",
    "add_draft_slides") | Sort-Object)) {
    throw "Presentation tool catalog contains an unexpected capability."
}

$crossAppToolNames = [regex]::Matches(
    $crossAppCatalogSource,
    'public const string \w+ = "([^"]+)";'
) | ForEach-Object { $_.Groups[1].Value } | Sort-Object
if (Compare-Object $crossAppToolNames (@(
    "create_email_draft",
    "send_to_powerpoint",
    "send_to_excel",
    "send_to_word") | Sort-Object)) {
    throw "Cross-app tool catalog contains an unexpected capability."
}

$wordToolNames = [regex]::Matches(
    $wordCatalogSource,
    'public const string \w+ = "([^"]+)";'
) | ForEach-Object { $_.Groups[1].Value } | Sort-Object
if (Compare-Object $wordToolNames (@(
    "read_document",
    "write_draft_document") | Sort-Object)) {
    throw "Word tool catalog contains an unexpected capability."
}

$documentDraftHostSource = Get-Content (
    Join-Path $sourceRoot "Office\DocumentDraftHost.cs") -Raw
foreach ($requiredBoundary in @(
    "OneShotDraftAuthorization",
    "authorization.TryConsume()",
    "authorization.MarkCreated()",
    "DRAFT_PERMISSION_NOT_AVAILABLE",
    "DRAFT_TOOL_MUST_BE_EXCLUSIVE"
)) {
    if (-not $documentDraftHostSource.Contains($requiredBoundary)) {
        throw "Document draft host is missing boundary $requiredBoundary."
    }
}

$workbookWriterSource = Get-Content (
    Join-Path $sourceRoot "Office\WorkbookDraftWriter.cs") -Raw
if (-not $workbookWriterSource.Contains(
        'DraftSheetName = "Scribble Draft"')) {
    throw "The Excel write surface must stay pinned to the Scribble Draft sheet."
}
if (-not $workbookWriterSource.Contains(
        "DraftFormulaPolicy.IsAllowedFormula")) {
    throw "Draft formulas must pass the formula safety policy."
}

$formulaPolicySource = Get-Content (
    Join-Path $sourceRoot "Security\DraftFormulaPolicy.cs") -Raw
foreach ($blockedFunction in @(
    '"WEBSERVICE"',
    '"RTD"',
    '"CALL"',
    '"HYPERLINK"'
)) {
    if (-not $formulaPolicySource.Contains($blockedFunction)) {
        throw "The formula policy no longer blocks $blockedFunction."
    }
}

$presentationWriterSource = Get-Content (
    Join-Path $sourceRoot "Office\PresentationDraftWriter.cs") -Raw
if (-not $presentationWriterSource.Contains(
        'DraftMarker = "[Scribble draft]"')) {
    throw "Appended slides must stay marked as Scribble drafts."
}
# The marker is written by the writer onto every drafted slide, so
# the model can never suppress it.
if (-not $presentationWriterSource.Contains("AddDraftTag(")) {
    throw "Every drafted slide must carry the Scribble draft marker."
}

# The corporate theme is hardcoded and owned by the writer: the
# model supplies content, never fonts, colors, sizes, or positions.
$themeSource = Get-Content (
    Join-Path $sourceRoot "Office\MetoTheme.cs") -Raw
foreach ($requiredToken in @(
    'ThemeName = "METO Executive Dense"',
    'TitleFont = "Samsung Sharp Sans Bold"',
    'BodyFont = "Calibri"',
    'BrandBlueHex = "#1428A0"',
    'CardSlateHex = "#E7ECF0"',
    'GoodGreenHex = "#E2EFDA"',
    'BadYellowHex = "#FFF2CC"',
    "MaxHighlightsPerStatus = 4"
)) {
    if (-not $themeSource.Contains($requiredToken)) {
        throw "The corporate theme is missing token $requiredToken."
    }
}
if (-not $presentationWriterSource.Contains("MetoTheme.")) {
    throw "Draft slides must be painted from the corporate theme."
}
foreach ($styleKey in @(
    '"color"',
    '"font"',
    '"font_size"',
    '"position"',
    '"left"',
    '"top"'
)) {
    if ($presentationCatalogSource.Contains($styleKey)) {
        throw "The slide schema must not accept styling from the model."
    }
}

$wordWriterSource = Get-Content (
    Join-Path $sourceRoot "Office\WordDraftWriter.cs") -Raw
if (-not $wordWriterSource.Contains(
        'DraftMarker = "[Scribble draft]"') -or
    -not $wordWriterSource.Contains("Documents.Add()")) {
    throw "Word drafts must stay marked, new, and unsaved."
}

foreach ($requiredDocumentBoundary in @(
    "can never send email",
    "The local host recognized an explicit draft request",
    "did not recognize an explicit draft, insert, or",
    "untrusted reference data"
)) {
    if (-not $documentFactorySource.Contains($requiredDocumentBoundary)) {
        throw "Document factory is missing boundary $requiredDocumentBoundary."
    }
}

$officePaneSource = Get-Content (
    Join-Path $sourceRoot "UI\OfficeChatPane.cs") -Raw
if (-not $officePaneSource.Contains(
        "DocumentDraftIntentPolicy.AllowsDraft(prompt)")) {
    throw "Document drafts are not gated by the local intent policy."
}

# End-user settings always reset request budgets to the reviewed
# defaults. The legacy clamp implementation remains as an internal
# defense, but no Settings tab can activate it.
$textBoundarySource = Get-Content (
    Join-Path $sourceRoot "Security\TextBoundary.cs") -Raw
foreach ($requiredLimitBoundary in @(
    "MaxPromptCharacters = 16000",
    "MaxAssistantCharactersLimit = 48000",
    "MaxToolRoundsLimit = 8",
    "MaxUserMultiplier = 8",
    "MaxWorkingSetMessages = 50",
    "if (useRecommended)"
)) {
    if (-not $textBoundarySource.Contains($requiredLimitBoundary)) {
        throw "Limit overrides are missing clamp $requiredLimitBoundary."
    }
}
if (-not $chatPaneSource.Contains("ApplyLimits()") -or
    -not $officePaneSource.Contains("ApplyLimits()")) {
    throw "Panes do not apply the configured limits."
}
$appSettingsSource = Get-Content (
    Join-Path $sourceRoot "Configuration\AppSettings.cs") -Raw
if (-not $appSettingsSource.Contains(
        "LimitOverrides.Apply(`r`n                true,") -and
    -not $appSettingsSource.Contains(
        "LimitOverrides.Apply(`n                true,")) {
    throw "App settings must always select the reviewed limits."
}
if ($settingsWindowSource.Contains(
        "tabs.TabPages.Add(BuildLimitsPage())")) {
    throw "The end-user Limits tab must remain absent."
}

# MCP tools stay namespaced, bounded, and separated from the draft
# and mailbox capability surfaces.
$mcpHostSource = Get-Content (
    Join-Path $sourceRoot "Chat\McpToolHost.cs") -Raw
foreach ($requiredMcpBoundary in @(
    'ToolPrefix = "mcp_"',
    "untrusted_mcp_data",
    "MaxExposedTools = 40",
    "MaxBrowserServers = 1",
    "BrowserOperationTimeoutMs = 30000",
    "BrowserToolsApproved",
    "browserTools.Contains(tool.Name)"
)) {
    if (-not $mcpHostSource.Contains($requiredMcpBoundary)) {
        throw "MCP host is missing boundary $requiredMcpBoundary."
    }
}
foreach ($mcpForbidden in @(
    "DraftService",
    "DraftToolHost",
    "DocumentDraftHost",
    "MailboxContextService"
)) {
    if ($mcpHostSource.Contains($mcpForbidden)) {
        throw "MCP host references forbidden capability $mcpForbidden."
    }
}

if (-not $settingsWindowSource.Contains(
        "outside this add-in's guardrails")) {
    throw "The MCP settings page is missing its trust notice."
}
$mcpConfigSource = Get-Content (
    Join-Path $sourceRoot "Configuration\McpServerConfig.cs") -Raw
foreach ($browserMcpBoundary in @(
    "ParsedBrowserTools",
    "BrowserToolsApproved",
    "Exact, case-sensitive MCP tool names"
)) {
    if (-not $mcpConfigSource.Contains($browserMcpBoundary)) {
        throw "Browser MCP configuration is missing boundary $browserMcpBoundary."
    }
}
if (-not $settingsWindowSource.Contains("I verified ") -or
    -not $settingsWindowSource.Contains(
        "that they are read-only")) {
    throw "MCP settings must require explicit read-only browser-tool approval."
}
if (([regex]::Matches(
        $settingsStoreSource,
        "\bBrowserToolsApproved\s*=")).Count -lt 2 -or
    ([regex]::Matches(
        $settingsStoreSource,
        "\bBrowserTools\s*=")).Count -lt 2) {
    throw "Settings storage must round-trip the browser MCP allowlist and approval."
}

# Direct Gemini is default-deny for end users. Settings load avoids
# decrypting dormant Google credentials, the UI omits the tab, and
# every gateway entry point refuses to run before credential access.
$settingsStoreSource = Get-Content (
    Join-Path $sourceRoot "Configuration\SettingsStore.cs") -Raw
if (-not $settingsStoreSource.Contains(
        "AdminPolicy.GeminiDisabled")) {
    throw "Settings load must honor the Gemini disable policy."
}
foreach ($requiredRenameMigration in @(
    "LegacyProductDirectoryName",
    "ResolveLoadPath()",
    "File.Copy(_legacySettingsPath, _settingsPath, false)",
    'LegacyProductDirectoryName + ".Settings.v1"'
)) {
    if (-not $settingsStoreSource.Contains($requiredRenameMigration)) {
        throw "Settings rename migration is missing $requiredRenameMigration."
    }
}
$adminPolicySource = Get-Content (
    Join-Path $sourceRoot "Configuration\AdminPolicy.cs") -Raw
if (-not $adminPolicySource.Contains(
        "GeminiEnabledForEndUsers") -or
    -not $adminPolicySource.Contains("#if SCRIBBLE_DIRECT_GEMINI") -or
    -not $adminPolicySource.Contains("return false;") -or
    -not $adminPolicySource.Contains("LegacyPolicyKeyPath") -or
    ([regex]::Matches(
        $adminPolicySource,
        "LegacyPolicyKeyPath")).Count -lt 3) {
    throw "The capability-reducing legacy admin policy is not honored."
}
$geminiGatewaySource = Get-Content (
    Join-Path $sourceRoot "Chat\GeminiCodeAssistGateway.cs") -Raw
if (-not $geminiGatewaySource.Contains(
        "GEMINI_DISABLED_BY_POLICY") -or
    ([regex]::Matches(
        $geminiGatewaySource,
        "EnsureGeminiAllowed\(\);")).Count -lt 4) {
    throw "The Gemini gateway must refuse to run under policy."
}
if (-not $settingsWindowSource.Contains(
        "if (!AdminPolicy.GeminiDisabled)") -or
    -not $settingsWindowSource.Contains(
        "tabs.TabPages.Add(BuildGeminiPage())")) {
    throw "The Gemini settings page must remain behind the product gate."
}
if (-not $settingsStoreSource.Contains(
        "GeminiRefreshToken = geminiDisabled") -or
    -not $settingsStoreSource.Contains(
        "ProtectedGeminiRefreshToken =")) {
    throw "Disabled Gemini credentials must stay out of memory and new saves."
}

# HTTP is accepted without an opt-in control. A non-loopback HTTP
# endpoint must still produce a clear plaintext-transport warning.
if ($settingsWindowSource.Contains("_allowInsecureHttp") -or
    -not $settingsWindowSource.Contains(
        "this remote HTTP endpoint receives the API key")) {
    throw "The streamlined HTTP behavior or warning is missing."
}
if (-not $appSettingsSource.Contains(
        "AllowInsecureHttp { get; set; } = true") -or
    -not $settingsStoreSource.Contains(
        "AllowInsecureHttp = true")) {
    throw "HTTP endpoints must work without a separate opt-in."
}

# The document-side model-facing sources carry the same capability
# hygiene as the mailbox ones.
$documentModelFacingSource =
    $workbookCatalogSource +
    $presentationCatalogSource +
    $crossAppCatalogSource +
    $wordCatalogSource +
    $documentFactorySource
foreach ($capability in @(
    "DraftService",
    "System.Diagnostics.Process",
    "Process.Start",
    "WebBrowser"
)) {
    if ($documentModelFacingSource.Contains($capability)) {
        throw "Document model-facing source references forbidden capability $capability."
    }
}

# ---- Scribble browser companion guardrails ----

$repositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$browserExtensionRoot = Join-Path $repositoryRoot "src\Scribble.BrowserExtension"
$browserHostRoot = Join-Path $repositoryRoot "src\Scribble.BrowserHost"
$browserInstallerPath = Join-Path $repositoryRoot "installer\Scribble.iss"
$browserManifestPath = Join-Path $browserExtensionRoot "manifest.json"
$nativeManifestPath = Join-Path $browserHostRoot "com.scribble.browser.json"
$browserHostProgramPath = Join-Path $browserHostRoot "Program.cs"
$browserSetupPath = Join-Path $browserHostRoot "BrowserSetup.cs"
$nativeProtocolPath = Join-Path $browserHostRoot "NativeMessageProtocol.cs"
$browserFactoryPath = Join-Path $sourceRoot "Chat\BrowserChatRequestFactory.cs"
$browserServicePath = Join-Path $sourceRoot "Chat\BrowserChatService.cs"
$expectedExtensionId = "olkepladbgkfkhlglooilnmalckpdada"
$expectedOrigin = "chrome-extension://$expectedExtensionId/"
$nativeHostName = "com.scribble.browser"

foreach ($requiredBrowserFile in @(
    $browserManifestPath,
    (Join-Path $browserExtensionRoot "background.js"),
    (Join-Path $browserExtensionRoot "sidepanel.html"),
    (Join-Path $browserExtensionRoot "sidepanel.css"),
    (Join-Path $browserExtensionRoot "sidepanel.js"),
    $nativeManifestPath,
    $browserHostProgramPath,
    $browserSetupPath,
    $nativeProtocolPath,
    $browserFactoryPath,
    $browserServicePath,
    $browserInstallerPath
)) {
    if (-not (Test-Path -LiteralPath $requiredBrowserFile -PathType Leaf)) {
        throw "Browser companion file is missing: $requiredBrowserFile"
    }
}

try {
    $browserManifest = Get-Content -LiteralPath $browserManifestPath -Raw |
        ConvertFrom-Json
}
catch {
    throw "The browser extension manifest is not valid JSON: $($_.Exception.Message)"
}
try {
    $nativeManifest = Get-Content -LiteralPath $nativeManifestPath -Raw |
        ConvertFrom-Json
}
catch {
    throw "The native messaging manifest is not valid JSON: $($_.Exception.Message)"
}

if ($browserManifest.manifest_version -ne 3) {
    throw "The browser extension must remain on Manifest V3."
}
$approvedBrowserPermissions = @(
    "activeTab",
    "contextMenus",
    "nativeMessaging",
    "scripting",
    "sidePanel"
) | Sort-Object
$actualBrowserPermissions = @($browserManifest.permissions) | Sort-Object
if (Compare-Object $approvedBrowserPermissions $actualBrowserPermissions) {
    throw "The browser extension permission set changed from the approved temporary-access surface."
}
foreach ($forbiddenManifestProperty in @(
    "host_permissions",
    "optional_host_permissions",
    "optional_permissions",
    "externally_connectable",
    "web_accessible_resources"
)) {
    if ($browserManifest.PSObject.Properties.Name -contains
        $forbiddenManifestProperty) {
        throw "The browser extension must not declare $forbiddenManifestProperty."
    }
}
if ($browserManifest.background.service_worker -ne "background.js" -or
    $browserManifest.side_panel.default_path -ne "sidepanel.html") {
    throw "The browser extension entry points changed unexpectedly."
}
$extensionCsp = [string]$browserManifest.content_security_policy.extension_pages
if (-not $extensionCsp.Contains("connect-src 'none'") -or
    -not $extensionCsp.Contains("object-src 'none'") -or
    $extensionCsp -match "unsafe-(?:eval|inline)|https?:|\*://") {
    throw "The extension content security policy permits remote or unsafe content."
}

function Get-ChromiumExtensionId([string]$ManifestKey) {
    try {
        $publicKey = [Convert]::FromBase64String($ManifestKey)
    }
    catch {
        throw "The extension manifest key is not valid Base64."
    }

    $sha256 = [Security.Cryptography.SHA256]::Create()
    try {
        $hash = $sha256.ComputeHash($publicKey)
    }
    finally {
        $sha256.Dispose()
    }

    $id = New-Object Text.StringBuilder
    for ($index = 0; $index -lt 16; $index++) {
        [void]$id.Append([char]([int][char]'a' + ($hash[$index] -shr 4)))
        [void]$id.Append([char]([int][char]'a' + ($hash[$index] -band 15)))
    }
    return $id.ToString()
}

$derivedExtensionId = Get-ChromiumExtensionId ([string]$browserManifest.key)
if ($derivedExtensionId -ne $expectedExtensionId) {
    throw "The manifest key derives extension ID $derivedExtensionId instead of $expectedExtensionId."
}
if ($nativeManifest.name -ne $nativeHostName -or
    $nativeManifest.type -ne "stdio" -or
    $nativeManifest.path -ne "ScribbleBrowserHost.exe") {
    throw "The native messaging manifest host contract changed unexpectedly."
}
$allowedOrigins = @($nativeManifest.allowed_origins)
if ($allowedOrigins.Count -ne 1 -or $allowedOrigins[0] -ne $expectedOrigin) {
    throw "The native messaging manifest is not pinned to the approved extension origin."
}
$browserHostProgramSource = Get-Content -LiteralPath $browserHostProgramPath -Raw
$programOrigins = @(
    [regex]::Matches(
        $browserHostProgramSource,
        'chrome-extension://[a-p]{32}/'
    ) | ForEach-Object { $_.Value } | Sort-Object -Unique
)
if ($programOrigins.Count -ne 1 -or $programOrigins[0] -ne $expectedOrigin) {
    throw "The native host executable and manifest do not enforce the same extension origin."
}

# No extension file may fetch executable content, inject HTML, navigate
# tabs, or grow a second browser capability surface. Page reads and the
# visible-tab screenshot remain explicit active-tab operations.
$browserExecutableFiles = Get-ChildItem $browserExtensionRoot -Recurse -File |
    Where-Object { $_.Extension -in @(".js", ".html", ".css") }
$dangerousBrowserPatterns = @(
    @{ Pattern = 'https?:\/\/|(?:src|href)\s*=\s*["'']\/\/'; Name = "remote URL" },
    @{ Pattern = '\b(?:eval|Function)\s*\(|\bimportScripts\s*\('; Name = "dynamic code execution" },
    @{ Pattern = '\b(?:fetch|XMLHttpRequest|WebSocket|EventSource)\b'; Name = "remote transport" },
    @{ Pattern = '\b(?:innerHTML|outerHTML|insertAdjacentHTML|document\.write)\b'; Name = "HTML injection" },
    @{ Pattern = 'createElement\s*\(\s*["'']script["'']'; Name = "dynamic script element" },
    @{ Pattern = 'chrome\.(?:bookmarks|cookies|debugger|downloads|history|management|webRequest)\b'; Name = "unapproved browser API" },
    @{ Pattern = 'chrome\.tabs\.(?:create|discard|duplicate|executeScript|goBack|goForward|group|highlight|move|reload|remove|ungroup|update)\b'; Name = "tab mutation" },
    @{ Pattern = 'chrome\.scripting\.(?:insertCSS|registerContentScripts|removeCSS|unregisterContentScripts|updateContentScripts)\b'; Name = "page mutation" },
    @{ Pattern = 'chrome\.windows\.(?:create|remove|update)\b|\bwindow\.open\s*\('; Name = "window mutation" },
    @{ Pattern = '@import\s+url|url\s*\(\s*["'']?https?:'; Name = "remote stylesheet" }
)
foreach ($dangerousBrowserPattern in $dangerousBrowserPatterns) {
    $hits = $browserExecutableFiles |
        Select-String -Pattern $dangerousBrowserPattern.Pattern
    if ($hits) {
        $firstHit = $hits | Select-Object -First 1
        throw "Browser extension contains $($dangerousBrowserPattern.Name): $($firstHit.Path):$($firstHit.LineNumber)."
    }
}

$browserInstallerSource = Get-Content -LiteralPath $browserInstallerPath -Raw
$browserInstallerLines = Get-Content -LiteralPath $browserInstallerPath
if ($browserInstallerSource -notmatch
    '(?m)^PrivilegesRequired=lowest\s*$') {
    throw "Browser support must preserve the per-user, non-elevated installer."
}
if ($browserInstallerSource -match
    'Source:\s*"[^\r\n"]*Scribble\.BrowserExtension\\\*') {
    throw "Installer must enumerate browser extension assets explicitly."
}
$browserHostInstallSource = (
    Get-ChildItem $browserHostRoot -Recurse -Filter *.cs |
        ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw }
) -join [Environment]::NewLine
$browserInstallSurface = $browserInstallerSource + $browserHostInstallSource
foreach ($browserInstallTrick in @(
    "--load-extension",
    "ExtensionInstallForcelist",
    "ExtensionInstallSources",
    "ExtensionSettings",
    "ExtensionInstallAllowlist",
    "ExtensionInstallBlocklist",
    "NativeMessagingAllowlist",
    "NativeMessagingBlocklist",
    "Software\Policies\Google\Chrome",
    "Software\Policies\Microsoft\Edge",
    "Software\Policies\Chromium",
    "Software\Google\Chrome\Extensions",
    "Software\Microsoft\Edge\Extensions",
    "Software\Chromium\Extensions",
    "Software\Wow6432Node\Google\Chrome\Extensions",
    "External Extensions",
    "master_preferences",
    "Secure Preferences",
    "Local Extension Settings",
    "User Data\",
    "--disable-extensions-except",
    "--user-data-dir"
)) {
    if ($browserInstallSurface.IndexOf(
            $browserInstallTrick,
            [StringComparison]::OrdinalIgnoreCase) -ge 0) {
        throw "Browser setup must not use policy, profile, or forced-install trick: $browserInstallTrick."
    }
}

$browserRunLines = @(
    Get-Content -LiteralPath $browserInstallerPath |
        Where-Object {
            $_ -match 'ScribbleBrowserHost\.exe' -and
            $_ -match '--setup\s+auto'
        }
)
if (@($browserRunLines).Count -ne 1 -or
    $browserRunLines[0] -notmatch 'Components:\s*browser' -or
    $browserRunLines[0] -notmatch 'Flags:[^;]*\bpostinstall\b' -or
    $browserRunLines[0] -notmatch 'Flags:[^;]*\bskipifsilent\b') {
    throw "The guided browser setup must be one optional post-install action and must skip silent installs."
}

if ($browserInstallerSource -notmatch
    '#define\s+BrowserNativeHostName\s+"com\.scribble\.browser"') {
    throw "Installer native host macro does not match the signed host identity."
}
$browserRegistryContracts = @(
    @{
        Display = "Software\Microsoft\Edge\NativeMessagingHosts\$nativeHostName"
        Literal = "Software\Microsoft\Edge\NativeMessagingHosts\{#BrowserNativeHostName}"
    },
    @{
        Display = "Software\Google\Chrome\NativeMessagingHosts\$nativeHostName"
        Literal = "Software\Google\Chrome\NativeMessagingHosts\{#BrowserNativeHostName}"
    }
)
foreach ($browserRegistryContract in $browserRegistryContracts) {
    $browserRegistryPath = $browserRegistryContract.Display
    $browserRegistryLiteral = $browserRegistryContract.Literal
    $registrationLines = @(
        Get-Content -LiteralPath $browserInstallerPath |
            Where-Object {
                $_.Contains($browserRegistryLiteral) -and
                $_ -match '^Root:\s*HKCU32;' -and
                $_ -match 'ValueData:\s*"\{app\}\\com\.scribble\.browser\.json"' -and
                $_ -match 'Components:\s*browser'
            }
    )
    if (@($registrationLines).Count -ne 1) {
        throw "Installer is missing the per-user native host registration for $browserRegistryPath."
    }
    if ($registrationLines[0] -notmatch 'Flags:[^;]*\buninsdeletekey\b') {
        throw "Native host registration must be removed on uninstall: $browserRegistryPath."
    }

    $deselectionLines = Get-Content -LiteralPath $browserInstallerPath |
        Where-Object {
            $_.Contains($browserRegistryLiteral) -and
            $_ -match '^Root:\s*HKCU32;' -and
            $_ -match 'Flags:[^;]*\bdeletekey\b' -and
            $_ -match 'Components:\s*not browser'
        }
    if (@($deselectionLines).Count -ne 1) {
        throw "Installer does not unregister a deselected browser component: $browserRegistryPath."
    }
}
$allNativeRegistrationLines = @(
    $browserInstallerLines |
        Where-Object { $_ -match 'NativeMessagingHosts' }
)
if ($allNativeRegistrationLines.Count -ne 6 -or
    @($allNativeRegistrationLines |
        Where-Object { $_ -notmatch '^Root:\s*HKCU32;' }).Count -ne 0) {
    throw "Native messaging registration must consist only of six HKCU32 entries."
}
$legacyNativeCleanupLines = @(
    $allNativeRegistrationLines |
        Where-Object {
            $_ -match '\{#LegacyBrowserNativeHostName\}'
        }
)
if ($legacyNativeCleanupLines.Count -ne 2 -or
    @($legacyNativeCleanupLines |
        Where-Object {
            $_ -notmatch 'Flags:[^;]*\bdeletekey\b' -or
            $_ -match 'ValueData:'
        }).Count -ne 0) {
    throw "The retired native-host registrations are not deleted safely."
}

$nativeProtocolSource = Get-Content -LiteralPath $nativeProtocolPath -Raw
foreach ($requiredProtocolBoundary in @(
    "MaxRequestBytes = 16 * 1024 * 1024",
    "MaxResponseBytes = 900 * 1024",
    "MaxHistoryTurns = 12",
    "new UTF8Encoding(false, true)",
    "length <= 0 || length > MaxRequestBytes",
    "TimeSpan.FromSeconds(230)",
    "REQUEST_TYPE_NOT_ALLOWED"
)) {
    if (-not $nativeProtocolSource.Contains($requiredProtocolBoundary)) {
        throw "Native messaging protocol is missing boundary $requiredProtocolBoundary."
    }
}

$browserFactorySource = Get-Content -LiteralPath $browserFactoryPath -Raw
foreach ($requiredBrowserBoundary in @(
    "read-only web-page assistant",
    "McpToolHost.IsMcpTool",
    "MaxSelectionCharacters = 16000",
    "MaxPageCharacters = 48000",
    "MaxHistoryTurns = 12",
    "MaxScreenshotDataUrlCharacters",
    "MaxScreenshotBytes",
    "5 * 1024 * 1024",
    "Convert.FromBase64String",
    "HasImageSignature"
)) {
    if (-not $browserFactorySource.Contains($requiredBrowserBoundary)) {
        throw "Browser request factory is missing boundary $requiredBrowserBoundary."
    }
}
foreach ($forbiddenBrowserCapability in @(
    "DraftService",
    "DraftToolHost",
    "DocumentDraftHost",
    "MailboxContextService",
    "Process.Start",
    "WebBrowser"
)) {
    if ($browserFactorySource.Contains($forbiddenBrowserCapability)) {
        throw "Browser request factory references forbidden capability $forbiddenBrowserCapability."
    }
}

$browserServiceSource = Get-Content -LiteralPath $browserServicePath -Raw
foreach ($requiredBrowserServiceBoundary in @(
    "_settings.ApplyLimits()",
    "BrowserChatRequestFactory.NormalizeScreenshot",
    "SCREENSHOT_INVALID",
    "ModelRouting.ResolveForRequest",
    "McpToolHost.IsMcpTool",
    "BROWSER_TOOL_NOT_ALLOWED",
    "MaxBrowserToolRounds = 1",
    "MaxBrowserToolCallsPerRound = 1"
)) {
    if (-not $browserServiceSource.Contains($requiredBrowserServiceBoundary)) {
        throw "Browser chat service is missing boundary $requiredBrowserServiceBoundary."
    }
}
if ($browserServiceSource -notmatch
    'new McpToolHost\(\s*_settings\.McpServers,\s*true\)') {
    throw "Browser chat must construct MCP in explicit allowlist-only mode."
}
foreach ($forbiddenBrowserServiceCapability in @(
    "DraftService",
    "DraftToolHost",
    "DocumentDraftHost",
    "MailboxContextService",
    "System.Diagnostics.Process",
    "Process.Start",
    "WebBrowser"
)) {
    if ($browserServiceSource.Contains($forbiddenBrowserServiceCapability)) {
        throw "Browser chat service references forbidden capability $forbiddenBrowserServiceCapability."
    }
}

Write-Host "PASS: static guardrail scan"

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;
using Scribble.Chat;
using Scribble.Outlook;
using Scribble.Security;
using Scribble.Utilities;

namespace Scribble.Office
{
    // Every mutating tool of the Excel and PowerPoint panes runs
    // through this host, behind the same one-shot authorization the
    // Outlook draft host uses: the user's own prompt must have
    // authorized a draft, and the tool must be the only call in its
    // response. Normal drafts consume permission on their first
    // attempt; selection output consumes it only after staging and
    // final preflight succeed.
    // The normal write surfaces are clearly marked drafts. Excel's
    // selection output is the narrow exception: it commits inert
    // text once to a snapshot-bound blank column, or to that exact
    // source only after explicit replacement intent. Nothing is
    // saved or sent.
    public sealed partial class DocumentDraftHost : IDisposable
    {
        private static readonly HashSet<string> EmailArguments =
            new HashSet<string>(StringComparer.Ordinal)
            {
                "to",
                "cc",
                "subject",
                "body",
                "attach_current_file"
            };

        private static readonly HashSet<string> SheetArguments =
            new HashSet<string>(StringComparer.Ordinal)
            {
                "title",
                "rows",
                "chart"
            };

        private static readonly HashSet<string> CellsArguments =
            new HashSet<string>(StringComparer.Ordinal)
            {
                "start_cell",
                "rows"
            };

        private static readonly HashSet<string> SelectionOutputArguments =
            new HashSet<string>(StringComparer.Ordinal)
            {
                "selection_handle",
                "destination_column",
                "start_offset",
                "values",
                "complete",
                "replace_source"
            };

        private static readonly HashSet<string> KoreanOutputArguments =
            new HashSet<string>(StringComparer.Ordinal)
            {
                "workbook_handle",
                "start_offset",
                "values",
                "complete"
            };

        private static readonly HashSet<string> SlideArguments =
            new HashSet<string>(StringComparer.Ordinal)
            {
                "slides",
                "after_slide"
            };

        private static readonly HashSet<string> DocumentArguments =
            new HashSet<string>(StringComparer.Ordinal)
            {
                "title",
                "placement",
                "body"
            };

        private readonly string _hostKind;
        private readonly object _hostApplication;
        private readonly JavaScriptSerializer _serializer =
            new JavaScriptSerializer();
        private DraftSession _emailDraft;
        private string _latestUserPrompt = string.Empty;
        private ExcelSelectionRequestContext _selectionRequest;
        private ExcelSelectionOutputSession _selectionOutput;
        private bool _allowSelectionSourceReplacement;
        private bool _selectionReplaceSource;
        private KoreanWorkbookRequestContext _koreanWorkbookRequest;
        private KoreanWorkbookOutputSession _koreanWorkbookOutput;

        public DocumentDraftHost(
            string hostKind,
            object hostApplication)
        {
            _hostKind = hostKind ?? string.Empty;
            _hostApplication = hostApplication ??
                throw new ArgumentNullException(
                    nameof(hostApplication));
        }

        public static bool IsDraftTool(
            string hostKind,
            string name)
        {
            if (CrossAppToolCatalog.IsCrossAppTool(name))
            {
                if (string.Equals(
                        name,
                        CrossAppToolCatalog.SendToPowerPoint,
                        StringComparison.Ordinal))
                {
                    return hostKind == "excel" ||
                           hostKind == "word" ||
                           hostKind == "outlook";
                }

                if (string.Equals(
                        name,
                        CrossAppToolCatalog.SendToExcel,
                        StringComparison.Ordinal))
                {
                    return hostKind == "powerpoint" ||
                           hostKind == "word" ||
                           hostKind == "outlook";
                }

                if (string.Equals(
                        name,
                        CrossAppToolCatalog.SendToWord,
                        StringComparison.Ordinal))
                {
                    return hostKind == "excel" ||
                           hostKind == "powerpoint" ||
                           hostKind == "outlook";
                }

                // create_email_draft: the Outlook pane has its own
                // dedicated draft host, so only the document hosts
                // route email drafts through here.
                return hostKind == "excel" ||
                       hostKind == "powerpoint" ||
                       hostKind == "word";
            }

            if (hostKind == "excel")
            {
                return WorkbookToolCatalog.IsDraftTool(name);
            }

            if (hostKind == "powerpoint")
            {
                return PresentationToolCatalog.IsDraftTool(name);
            }

            if (hostKind == "word")
            {
                return WordToolCatalog.IsDraftTool(name);
            }

            return false;
        }

        public MailboxToolResult Execute(
            ChatToolCall call,
            OneShotDraftAuthorization authorization,
            bool isOnlyToolCall,
            string userPrompt = null)
        {
            _latestUserPrompt = userPrompt ?? string.Empty;
            var name = call?.function?.name;
            if (call?.function == null ||
                string.IsNullOrWhiteSpace(call.id) ||
                !IsDraftTool(_hostKind, name))
            {
                return Error(
                    call?.id,
                    authorization,
                    "DRAFT_TOOL_CALL_INVALID",
                    "The model returned an invalid draft tool call.");
            }

            if (!isOnlyToolCall)
            {
                return Error(
                    call.id,
                    authorization,
                    "DRAFT_TOOL_MUST_BE_EXCLUSIVE",
                    name + " must be the only tool call in its response.");
            }

            IDictionary<string, object> arguments;
            try
            {
                arguments = ToolArguments.Parse(
                    _serializer,
                    call.function.arguments);
                RequireAllowedArguments(arguments, name);
            }
            catch (Exception exception)
            {
                return Error(
                    call.id,
                    authorization,
                    "DRAFT_ARGUMENTS_INVALID",
                    DiagnosticDetails.ForException(
                        exception,
                        "DRAFT_ARGUMENTS_INVALID"));
            }

            if (string.Equals(
                name,
                WorkbookToolCatalog.WriteSelectionOutput,
                StringComparison.Ordinal))
            {
                return ExecuteSelectionOutput(
                    call.id,
                    arguments,
                    authorization);
            }

            if (string.Equals(
                name,
                WorkbookToolCatalog.WriteKoreanTranslations,
                StringComparison.Ordinal))
            {
                return ExecuteKoreanWorkbookOutput(
                    call.id,
                    arguments,
                    authorization);
            }

            // A deck or workbook may be built over several bounded
            // calls, but one request may open at most ONE unsent
            // email draft - recipients are the sensitive surface,
            // so that path takes its own single-use permission.
            var isEmailDraft = string.Equals(
                name,
                CrossAppToolCatalog.CreateEmailDraft,
                StringComparison.Ordinal);
            var permitted = authorization != null &&
                authorization.CanCreate &&
                (isEmailDraft
                    ? authorization.TryConsumeEmailDraft()
                    : authorization.TryConsume());
            if (!permitted)
            {
                return Error(
                    call.id,
                    authorization,
                    "DRAFT_PERMISSION_NOT_AVAILABLE",
                    "No unused local draft permission is available for this request.");
            }

            try
            {
                string status;
                if (string.Equals(
                    name,
                    CrossAppToolCatalog.CreateEmailDraft,
                    StringComparison.Ordinal))
                {
                    status = CreateEmailDraft(arguments);
                }
                else if (string.Equals(
                             name,
                             WorkbookToolCatalog.WriteDraftSheet,
                             StringComparison.Ordinal))
                {
                    status = WorkbookDraftWriter.WriteDraftSheet(
                        _hostApplication,
                        ToolArguments.GetString(
                            arguments,
                            "title",
                            string.Empty),
                        ParsedRows(arguments),
                        ParsedChart(arguments));
                }
                else if (string.Equals(
                             name,
                             WorkbookToolCatalog.WriteCells,
                             StringComparison.Ordinal))
                {
                    status = WorkbookDraftWriter.WriteCells(
                        _hostApplication,
                        ToolArguments.GetString(
                            arguments,
                            "start_cell",
                            string.Empty),
                        ParsedRows(arguments));
                }
                else if (string.Equals(
                             name,
                             PresentationToolCatalog.AddDraftSlides,
                             StringComparison.Ordinal))
                {
                    status = PresentationDraftWriter.AddDraftSlides(
                        _hostApplication,
                        ParsedSlides(arguments),
                        ParsedAfterSlide(arguments));
                }
                else if (string.Equals(
                             name,
                             CrossAppToolCatalog.SendToPowerPoint,
                             StringComparison.Ordinal))
                {
                    // Cross-app sends always open a brand-new
                    // unsaved deck - the user was not looking at
                    // PowerPoint when they asked, so nothing lands
                    // inside whatever deck happens to be open.
                    status = PresentationDraftWriter.AddDraftSlides(
                        GetSiblingApplication(
                            "PowerPoint.Application"),
                        ParsedSlides(arguments),
                        ParsedAfterSlide(arguments),
                        true);
                }
                else if (string.Equals(
                             name,
                             CrossAppToolCatalog.SendToExcel,
                             StringComparison.Ordinal))
                {
                    // Cross-app sends always open a brand-new
                    // unsaved workbook - never a sheet slipped
                    // into whatever workbook happens to be open.
                    status = WorkbookDraftWriter.WriteDraftSheet(
                        GetSiblingApplication(
                            "Excel.Application"),
                        ToolArguments.GetString(
                            arguments,
                            "title",
                            string.Empty),
                        ParsedRows(arguments),
                        ParsedChart(arguments),
                        true);
                }
                else if (string.Equals(
                             name,
                             WordToolCatalog.WriteDraftDocument,
                             StringComparison.Ordinal))
                {
                    status = WordDraftWriter.WriteDraftDocument(
                        _hostApplication,
                        ToolArguments.GetString(
                            arguments,
                            "title",
                            string.Empty),
                        GetLongString(arguments, "body"),
                        ToolArguments.GetString(
                            arguments,
                            "placement",
                            "end"));
                }
                else if (string.Equals(
                             name,
                             CrossAppToolCatalog.SendToWord,
                             StringComparison.Ordinal))
                {
                    // Cross-app sends always land in a separate
                    // marked draft document - the user was not
                    // looking at Word when they asked.
                    status = WordDraftWriter.WriteDraftDocument(
                        GetSiblingApplication(
                            "Word.Application"),
                        ToolArguments.GetString(
                            arguments,
                            "title",
                            string.Empty),
                        GetLongString(arguments, "body"),
                        "new_document");
                }
                else
                {
                    throw new InvalidOperationException(
                        "The requested draft tool is not allowed.");
                }

                authorization.MarkCreated();
                return new MailboxToolResult(
                    call.id,
                    _serializer.Serialize(
                        new Dictionary<string, object>
                        {
                            { "ok", true },
                            { "action", name },
                            { "saved", false },
                            { "sent", false },
                            {
                                "status",
                                TextBoundary.PlainText(status, 600)
                            }
                        }),
                    TextBoundary.SingleLine(status, 300));
            }
            catch (Exception exception)
            {
                Log.Error("DocumentDraft." + name, exception);
                return Error(
                    call.id,
                    authorization,
                    "DRAFT_CREATION_FAILED",
                    DiagnosticDetails.ForException(
                        exception,
                        "DRAFT_CREATION_FAILED"));
            }
        }

        public void Dispose()
        {
            _emailDraft?.Dispose();
            _emailDraft = null;
            EndExcelSelectionRequest();
        }

        internal void BeginExcelSelectionRequest(
            ExcelSelectionRequestContext context)
        {
            _selectionRequest = context;
            _selectionOutput = null;
            _allowSelectionSourceReplacement =
                context != null && context.AllowSourceReplacement;
            _selectionReplaceSource = false;
        }

        internal void BeginKoreanWorkbookRequest(
            KoreanWorkbookRequestContext context)
        {
            _koreanWorkbookRequest = context;
            _koreanWorkbookOutput = null;
        }

        internal void AllowExcelSelectionSourceReplacement()
        {
            if (_selectionRequest != null)
            {
                _allowSelectionSourceReplacement = true;
                if (_taskContext != null)
                {
                    _taskContext.State.HostData["allow_source_replacement"] = "true";
                    _taskContext.Checkpoint();
                }
            }
        }

        internal void EndExcelSelectionRequest()
        {
            _selectionOutput = null;
            _selectionRequest = null;
            _allowSelectionSourceReplacement = false;
            _selectionReplaceSource = false;
            _koreanWorkbookOutput = null;
            _koreanWorkbookRequest = null;
        }

        private MailboxToolResult ExecuteKoreanWorkbookOutput(
            string callId,
            IDictionary<string, object> arguments,
            OneShotDraftAuthorization authorization)
        {
            if (authorization == null ||
                !authorization.CanCreate ||
                _koreanWorkbookRequest == null)
            {
                return Error(
                    callId,
                    authorization,
                    "KOREAN_WORKBOOK_PERMISSION_NOT_AVAILABLE",
                    "No workbook-wide Korean translation is attached " +
                    "to this request.");
            }

            var handle = ToolArguments.GetString(
                arguments,
                "workbook_handle",
                string.Empty);
            if (!string.Equals(
                handle,
                _koreanWorkbookRequest.Handle,
                StringComparison.Ordinal))
            {
                return Error(
                    callId,
                    authorization,
                    "KOREAN_WORKBOOK_HANDLE_UNKNOWN",
                    "The workbook translation handle is unknown or expired.");
            }

            try
            {
                if (_koreanWorkbookOutput == null)
                {
                    _koreanWorkbookOutput =
                        new KoreanWorkbookOutputSession(
                            handle,
                            _koreanWorkbookRequest.Snapshot.Cells.Count);
                }

                var values = ParseSelectionValues(arguments);
                var startOffset = ToolArguments.GetInteger(
                    arguments,
                    "start_offset",
                    -1,
                    -1,
                    _koreanWorkbookRequest.Snapshot.Cells.Count);
                var complete = ToolArguments.GetBoolean(
                    arguments,
                    "complete");
                var ready = _koreanWorkbookOutput.Stage(
                    handle,
                    startOffset,
                    values,
                    complete);
                if (!ready)
                {
                    return KoreanWorkbookSuccess(
                        callId,
                        false,
                        authorization);
                }

                if (_durableExcel != null) return KoreanWorkbookSuccess(callId, false, authorization);

                if (!authorization.TryConsume())
                {
                    return Error(
                        callId,
                        authorization,
                        "DRAFT_PERMISSION_NOT_AVAILABLE",
                        "No unused local write permission is available " +
                        "for the final workbook translation.");
                }

                var status = WorkbookSelectionOutputWriter
                    .CommitKoreanTranslations(
                        _hostApplication,
                        _koreanWorkbookRequest.Snapshot,
                        _koreanWorkbookOutput.Values);
                authorization.MarkCreated();
                return KoreanWorkbookSuccess(
                    callId,
                    true,
                    authorization,
                    status);
            }
            catch (Exception exception)
            {
                Log.Error("DocumentDraft." +
                    WorkbookToolCatalog.WriteKoreanTranslations,
                    exception);
                return Error(
                    callId,
                    authorization,
                    "KOREAN_WORKBOOK_OUTPUT_INVALID",
                    DiagnosticDetails.ForException(
                        exception,
                        "KOREAN_WORKBOOK_OUTPUT_INVALID"));
            }
        }

        private MailboxToolResult KoreanWorkbookSuccess(
            string callId,
            bool committed,
            OneShotDraftAuthorization authorization,
            string committedStatus = null)
        {
            var snapshot = _koreanWorkbookRequest.Snapshot;
            var staged = _koreanWorkbookOutput.StagedCount;
            var remaining = Math.Max(0, snapshot.Cells.Count - staged);
            var nextSourceCells = remaining == 0
                ? new Dictionary<string, string>[0]
                : WorkbookSelectionOutputWriter.ReadKoreanSourceWindow(
                    snapshot,
                    staged,
                    Math.Min(
                        ExcelSelectionOutputPolicy.PreferredBatchValues,
                        remaining));
            var nextBatchSize = nextSourceCells.Count;
            var status = committed
                ? committedStatus
                : "Prepared " + staged + " of " +
                  snapshot.Cells.Count +
                  " Korean cell translations. Excel is unchanged.";
            return new MailboxToolResult(
                callId,
                _serializer.Serialize(
                    new Dictionary<string, object>
                    {
                        { "ok", true },
                        { "action",
                            WorkbookToolCatalog.WriteKoreanTranslations },
                        { "committed", committed },
                        { "saved", false },
                        { "permission_consumed",
                            authorization.IsConsumed },
                        { "staged_count", staged },
                        { "expected_count", snapshot.Cells.Count },
                        { "next_start_offset", staged },
                        { "remaining_count", remaining },
                        { "next_batch_size", nextBatchSize },
                        { "next_source_cells", nextSourceCells },
                        { "complete_next",
                            remaining > 0 && remaining == nextBatchSize },
                        { "status", status }
                    }),
                TextBoundary.SingleLine(status, 300));
        }

        private MailboxToolResult ExecuteSelectionOutput(
            string callId,
            IDictionary<string, object> arguments,
            OneShotDraftAuthorization authorization)
        {
            if (authorization == null ||
                !authorization.CanCreate ||
                _selectionRequest == null)
            {
                return Error(
                    callId,
                    authorization,
                    "SELECTION_PERMISSION_NOT_AVAILABLE",
                    "No authorized Excel selection is attached to this request.");
            }

            var handle = ToolArguments.GetString(
                arguments,
                "selection_handle",
                string.Empty);
            if (!string.Equals(
                handle,
                _selectionRequest.Handle,
                StringComparison.Ordinal))
            {
                return Error(
                    callId,
                    authorization,
                    "SELECTION_HANDLE_UNKNOWN",
                    "The Excel selection handle is unknown or expired.");
            }

            var snapshot = _selectionRequest.Snapshot;
            if (snapshot.ColumnCount != 1 ||
                snapshot.RowCount < 1)
            {
                return Error(
                    callId,
                    authorization,
                    "SELECTION_OUTPUT_NOT_ELIGIBLE",
                    "Select one contiguous Excel column. The full " +
                    "selection is processed sequentially even when its " +
                    "inline preview is incomplete.");
            }

            var destination = ToolArguments.GetString(
                arguments,
                "destination_column",
                string.Empty);
            var replaceSourceSpecified =
                arguments.ContainsKey("replace_source");
            var replaceSource = replaceSourceSpecified &&
                ToolArguments.GetBoolean(arguments, "replace_source");
            if (_selectionOutput != null)
            {
                if (replaceSourceSpecified &&
                    replaceSource != _selectionReplaceSource)
                {
                    return Error(
                        callId,
                        authorization,
                        "SELECTION_OUTPUT_MODE_CHANGED",
                        "All batches must keep the replacement mode " +
                        "chosen by the first accepted batch.");
                }

                replaceSource = _selectionReplaceSource;
            }
            if (replaceSource && !_allowSelectionSourceReplacement)
            {
                return Error(
                    callId,
                    authorization,
                    "SELECTION_REPLACE_NOT_AUTHORIZED",
                    "Replacing the source requires the user's explicit " +
                    "replace, overwrite, or in-place instruction.");
            }

            var sourceColumn =
                ExcelSelectionOutputPolicy.ColumnNumberToName(
                    snapshot.StartColumn);
            if (replaceSource)
            {
                if (destination.Length > 0 &&
                    !string.Equals(
                        destination,
                        sourceColumn,
                        StringComparison.OrdinalIgnoreCase))
                {
                    return Error(
                        callId,
                        authorization,
                        "SELECTION_DESTINATION_CONFLICT",
                        "Do not provide a different destination column " +
                        "when replacing the selected source cells.");
                }

                destination = sourceColumn;
            }
            else if (destination.Length == 0)
            {
                destination =
                    ExcelSelectionOutputPolicy.ColumnNumberToName(
                        snapshot.StartColumn + 1);
            }

            if (destination.Length == 0)
            {
                return Error(
                    callId,
                    authorization,
                    "SELECTION_DESTINATION_REQUIRED",
                    "The source is in Excel's last column. Choose an " +
                    "empty destination column.");
            }

            try
            {
                // Read-only preflight happens before staging, and is
                // repeated before the final bulk commit.
                ValidateSelectionTarget(
                    snapshot,
                    destination,
                    replaceSource);

                if (_selectionOutput == null)
                {
                    _selectionOutput =
                        new ExcelSelectionOutputSession(
                            _selectionRequest.Handle,
                            snapshot.RowCount);
                    _selectionReplaceSource = replaceSource;
                }

                var values = ParseSelectionValues(arguments);
                var startOffset = ToolArguments.GetInteger(
                    arguments,
                    "start_offset",
                    -1,
                    -1,
                    snapshot.RowCount);
                var complete = ToolArguments.GetBoolean(
                    arguments,
                    "complete");
                if (complete && authorization.RemainingCalls <= 0)
                {
                    return Error(
                        callId,
                        authorization,
                        "DRAFT_PERMISSION_NOT_AVAILABLE",
                        "No unused local draft permission is available " +
                        "for the final selection write.");
                }

                var ready = _selectionOutput.Stage(
                    handle,
                    destination,
                    startOffset,
                    values,
                    complete);
                if (!ready)
                {
                    var staged = "Staged " +
                        _selectionOutput.StagedCount + " of " +
                        snapshot.RowCount +
                        " selection values. No Excel cells were changed.";
                    return SelectionSuccess(
                        callId,
                        staged,
                        false,
                        authorization,
                        _selectionOutput.StagedCount,
                        snapshot.RowCount);
                }

                if (_durableExcel != null) return SelectionSuccess(callId, "All rows staged for reviewed journaled commit.", false, authorization, _selectionOutput.StagedCount, snapshot.RowCount);

                // Validate once more immediately before consuming
                // permission; rejected preflights never spend it.
                ValidateSelectionTarget(
                    snapshot,
                    _selectionOutput.DestinationColumn,
                    replaceSource);
                if (!authorization.TryConsume())
                {
                    return Error(
                        callId,
                        authorization,
                        "DRAFT_PERMISSION_NOT_AVAILABLE",
                        "No unused local draft permission is available " +
                        "for the final selection write.");
                }

                var status = WorkbookSelectionOutputWriter.Commit(
                    _hostApplication,
                    snapshot,
                    _selectionOutput.DestinationColumn,
                    _selectionOutput.Values,
                    replaceSource);
                authorization.MarkCreated();
                return SelectionSuccess(
                    callId,
                    status,
                    true,
                    authorization,
                    _selectionOutput.StagedCount,
                    snapshot.RowCount);
            }
            catch (ExcelSelectionDestinationException exception)
            {
                return SelectionDestinationError(
                    callId,
                    authorization,
                    exception);
            }
            catch (Exception exception)
            {
                Log.Error("DocumentDraft." +
                    WorkbookToolCatalog.WriteSelectionOutput,
                    exception);
                return Error(
                    callId,
                    authorization,
                    "SELECTION_OUTPUT_INVALID",
                    DiagnosticDetails.ForException(
                        exception,
                        "SELECTION_OUTPUT_INVALID"));
            }
        }

        private void ValidateSelectionTarget(
            ExcelSelectionSnapshot snapshot,
            string destination,
            bool replaceSource)
        {
            if (replaceSource)
            {
                WorkbookSelectionOutputWriter.ValidateSource(
                    _hostApplication,
                    snapshot);
                return;
            }

            WorkbookSelectionOutputWriter.ValidateDestination(
                _hostApplication,
                snapshot,
                destination);
        }

        private static IReadOnlyList<string> ParseSelectionValues(
            IDictionary<string, object> arguments)
        {
            object raw;
            var outer = arguments.TryGetValue("values", out raw)
                ? raw as IEnumerable
                : null;
            if (outer == null || raw is string)
            {
                throw new InvalidOperationException(
                    "values must be an array of strings.");
            }

            var values = new List<string>();
            foreach (var value in outer)
            {
                var text = value as string;
                if (text == null)
                {
                    throw new InvalidOperationException(
                        "values must contain strings only.");
                }

                values.Add(text);
            }

            return values;
        }

        private MailboxToolResult SelectionSuccess(
            string callId,
            string status,
            bool committed,
            OneShotDraftAuthorization authorization,
            int stagedCount,
            int expectedCount)
        {
            var remaining = Math.Max(0, expectedCount - stagedCount);
            var nextSourceValues = remaining == 0
                ? new string[0]
                : WorkbookSelectionOutputWriter.ReadSourceValues(
                    _hostApplication,
                    _selectionRequest.Snapshot,
                    stagedCount,
                    Math.Min(
                        ExcelSelectionOutputPolicy.PreferredBatchValues,
                        remaining));
            var nextBatchSize = nextSourceValues.Count;
            return new MailboxToolResult(
                callId,
                _serializer.Serialize(
                    new Dictionary<string, object>
                    {
                        { "ok", true },
                        {
                            "action",
                            WorkbookToolCatalog.WriteSelectionOutput
                        },
                        { "committed", committed },
                        { "saved", false },
                        { "permission_consumed",
                            authorization.IsConsumed },
                        { "staged_count", stagedCount },
                        { "expected_count", expectedCount },
                        { "next_start_offset", stagedCount },
                        { "remaining_count", remaining },
                        { "next_batch_size", nextBatchSize },
                        { "next_source_values", nextSourceValues },
                        {
                            "complete_next",
                            remaining > 0 &&
                            remaining == nextBatchSize
                        },
                        { "status", status }
                    }),
                TextBoundary.SingleLine(status, 300));
        }

        private MailboxToolResult SelectionDestinationError(
            string callId,
            OneShotDraftAuthorization authorization,
            ExcelSelectionDestinationException exception)
        {
            return new MailboxToolResult(
                callId,
                _serializer.Serialize(
                    new Dictionary<string, object>
                    {
                        { "ok", false },
                        { "error_code", exception.Code },
                        { "message", exception.Message },
                        { "empty_destination_candidates",
                            exception.Candidates },
                        { "permission_consumed",
                            authorization != null &&
                            authorization.IsConsumed }
                    }),
                "[" + exception.Code + "] " +
                TextBoundary.PlainText(exception.Message, 500));
        }

        // Long text arguments (draft bodies) must not ride through
        // the 1000-character generic string reader.
        private static string GetLongString(
            IDictionary<string, object> arguments,
            string key)
        {
            object value;
            return arguments.TryGetValue(key, out value)
                ? TextBoundary.PlainText(
                    Convert.ToString(value),
                    WordDraftWriter.MaxDraftCharacters)
                : string.Empty;
        }

        private static IReadOnlyList<IReadOnlyList<string>>
            ParsedRows(IDictionary<string, object> arguments)
        {
            object rowsValue;
            arguments.TryGetValue("rows", out rowsValue);
            return WorkbookDraftWriter.ParseRows(rowsValue);
        }

        private static WorkbookDraftWriter.DraftSheetChart
            ParsedChart(IDictionary<string, object> arguments)
        {
            object chartValue;
            arguments.TryGetValue("chart", out chartValue);
            return WorkbookDraftWriter.ParseChart(chartValue);
        }

        private static
            IReadOnlyList<PresentationDraftWriter.DraftSlide>
            ParsedSlides(IDictionary<string, object> arguments)
        {
            object slidesValue;
            arguments.TryGetValue("slides", out slidesValue);
            return PresentationDraftWriter.ParseSlides(slidesValue);
        }

        // Null when the model omitted after_slide, so the writer
        // keeps its default append-at-the-end behavior.
        private static int? ParsedAfterSlide(
            IDictionary<string, object> arguments)
        {
            return arguments.ContainsKey("after_slide")
                ? (int?)ToolArguments.GetInteger(
                    arguments,
                    "after_slide",
                    0,
                    0,
                    1000)
                : null;
        }

        // Opens one unsent Outlook draft for review via the same
        // DraftService the Outlook pane uses: the draft is saved to
        // Drafts and displayed, never sent - sending stays a human
        // action in Outlook.
        private string CreateEmailDraft(
            IDictionary<string, object> arguments)
        {
            var body = TextBoundary.PlainText(
                GetLongString(arguments, "body"),
                TextBoundary.MaxAssistantCharacters);
            if (body.Length == 0)
            {
                throw new InvalidOperationException(
                    "A non-empty draft body is required.");
            }

            var outlook = GetSiblingApplication(
                "Outlook.Application");
            var drafts = new DraftService(outlook);
            var to = ToolArguments.GetString(
                arguments,
                "to",
                string.Empty);
            var cc = ToolArguments.GetString(
                arguments,
                "cc",
                string.Empty);
            _emailDraft?.Dispose();
            _emailDraft = drafts.CreateNewDraft(
                body,
                new string[0],
                ToolArguments.GetString(
                    arguments,
                    "subject",
                    string.Empty),
                to,
                cc);
            var attached = false;
            if (ToolArguments.GetBoolean(
                arguments,
                "attach_current_file"))
            {
                var path = CurrentDocumentPath();
                if (path.Length > 0)
                {
                    _emailDraft.AttachFile(path);
                    attached = true;
                }
            }

            return "One unsent Outlook draft was opened for " +
                "review" +
                (attached
                    ? " with the current file attached"
                    : string.Empty) +
                ". Scribble cannot send it." +
                RecipientIntentCheck.Warn(
                    to,
                    cc,
                    _latestUserPrompt);
        }

        // Full path of the host document when it exists on disk;
        // empty for unsaved documents.
        private string CurrentDocumentPath()
        {
            try
            {
                if (_hostKind == "outlook")
                {
                    // The mailbox host has no single current
                    // document file to attach.
                    return string.Empty;
                }

                dynamic application = _hostApplication;
                dynamic document;
                if (_hostKind == "excel")
                {
                    document = application.ActiveWorkbook;
                }
                else if (_hostKind == "word")
                {
                    document = application.ActiveDocument;
                }
                else
                {
                    document = application.ActivePresentation;
                }
                if (document == null)
                {
                    return string.Empty;
                }

                var folder = Convert.ToString(document.Path) ??
                    string.Empty;
                if (folder.Length == 0)
                {
                    return string.Empty;
                }

                return Convert.ToString(document.FullName) ??
                    string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        // Reuses the running sibling Office application when there
        // is one; otherwise starts it visibly so the draft opens in
        // front of the user.
        private static object GetSiblingApplication(string progId)
        {
            try
            {
                return Marshal.GetActiveObject(progId);
            }
            catch
            {
            }

            var type = Type.GetTypeFromProgID(progId);
            if (type == null)
            {
                throw new InvalidOperationException(
                    progId.Split('.')[0] +
                    " is not installed on this computer.");
            }

            var application = Activator.CreateInstance(type);
            try
            {
                dynamic visible = application;
                visible.Visible = true;
            }
            catch
            {
                // Outlook has no Visible property; drafts are
                // displayed by the draft session itself.
            }

            return application;
        }

        private void RequireAllowedArguments(
            IDictionary<string, object> arguments,
            string name)
        {
            ISet<string> allowed;
            if (string.Equals(
                name,
                CrossAppToolCatalog.CreateEmailDraft,
                StringComparison.Ordinal))
            {
                allowed = EmailArguments;
            }
            else if (
                string.Equals(
                    name,
                    WorkbookToolCatalog.WriteDraftSheet,
                    StringComparison.Ordinal) ||
                string.Equals(
                    name,
                    CrossAppToolCatalog.SendToExcel,
                    StringComparison.Ordinal))
            {
                allowed = SheetArguments;
            }
            else if (string.Equals(
                name,
                WorkbookToolCatalog.WriteCells,
                StringComparison.Ordinal))
            {
                allowed = CellsArguments;
            }
            else if (string.Equals(
                name,
                WorkbookToolCatalog.WriteSelectionOutput,
                StringComparison.Ordinal))
            {
                allowed = SelectionOutputArguments;
            }
            else if (string.Equals(
                name,
                WorkbookToolCatalog.WriteKoreanTranslations,
                StringComparison.Ordinal))
            {
                allowed = KoreanOutputArguments;
            }
            else if (
                string.Equals(
                    name,
                    WordToolCatalog.WriteDraftDocument,
                    StringComparison.Ordinal) ||
                string.Equals(
                    name,
                    CrossAppToolCatalog.SendToWord,
                    StringComparison.Ordinal))
            {
                allowed = DocumentArguments;
            }
            else
            {
                allowed = SlideArguments;
            }

            var unexpected = arguments.Keys
                .FirstOrDefault(key => !allowed.Contains(key));
            if (unexpected != null)
            {
                throw new InvalidOperationException(
                    "Unexpected draft argument: " + unexpected);
            }
        }

        private MailboxToolResult Error(
            string callId,
            OneShotDraftAuthorization authorization,
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
                        },
                        {
                            "permission_consumed",
                            authorization != null &&
                            authorization.IsConsumed
                        }
                    }),
                "[" + code + "] " +
                TextBoundary.PlainText(message, 600));
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using Scribble.Chat;
using Scribble.Security;

namespace Scribble.Office
{
    public sealed partial class DocumentDraftHost
    {
        private MailboxToolResult ExecuteAnalysisWorkbookDraft(
            string callId, IDictionary<string, object> arguments,
            OneShotDraftAuthorization authorization)
        {
            if (_hostKind != "excel" ||
                !string.Equals(Environment.GetEnvironmentVariable(
                    AnalysisDocumentPilot.FeatureFlag), "1",
                    StringComparison.Ordinal))
                return Error(callId, authorization,
                    "ANALYSIS_PILOT_DISABLED",
                    "This typed report route is available only in the development pilot.");
            if (_taskContext == null || authorization == null ||
                !authorization.CanCreate)
                return Error(callId, authorization,
                    "DRAFT_PERMISSION_NOT_AVAILABLE",
                    "The task needs the user's explicit draft instruction.");
            if (arguments.Keys.Except(new[] { "analysis_id", "title" },
                    StringComparer.Ordinal).Any())
                return Error(callId, authorization,
                    "ANALYSIS_DRAFT_ARGUMENTS_INVALID",
                    "Supply only analysis_id and an optional title; the host builds rows and formulas.");
            AnalysisArtifact artifact;
            AnalysisDocumentPlan plan;
            int formulaCount;
            try
            {
                artifact = _taskContext.LoadAnalysis();
                if (artifact == null ||
                    !string.Equals(ToolArguments.GetString(arguments,
                            "analysis_id", string.Empty),
                        artifact.AnalysisId, StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        "ANALYSIS_PLAN_BINDING_INVALID");
                var title = ToolArguments.GetString(arguments, "title",
                    "Verified analysis").Trim();
                if (title.Length == 0 || title.Length > 120 ||
                    title.StartsWith("=", StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        "ANALYSIS_WORKBOOK_TITLE_INVALID");
                plan = new AnalysisDocumentPlan
                {
                    AnalysisId = artifact.AnalysisId,
                    WorkbookTitle = title,
                    WorkbookRows = AnalysisWorkbookPlanBuilder.Build(
                        artifact)
                };
                formulaCount = AnalysisDocumentCompiler.Compile(artifact,
                    plan, false).ExpectedFormulaFacts.Count;
                OfficeTaskBinding.Validate(_taskContext.State, "excel",
                    _hostApplication);
                AnalysisWorkbookSourceGuard.Validate(_hostApplication,
                    artifact);
            }
            catch (Exception exception)
            {
                return Error(callId, authorization,
                    "ANALYSIS_DRAFT_PREFLIGHT_FAILED", exception.Message);
            }
            if (!authorization.TryConsume())
                return Error(callId, authorization,
                    "DRAFT_PERMISSION_NOT_AVAILABLE",
                    "The task's draft call budget is exhausted.");
            try
            {
                var status = AnalysisDocumentPilot.WriteWorkbook(
                    _hostApplication, artifact, plan);
                authorization.MarkCreated();
                return new MailboxToolResult(callId,
                    _serializer.Serialize(new
                    {
                        ok = true, saved = false,
                        analysis_id = artifact.AnalysisId,
                        verified_live_formulas = formulaCount,
                        status
                    }), status);
            }
            catch (Exception exception)
            {
                return Error(callId, authorization,
                    "ANALYSIS_DRAFT_FAILED", exception.Message);
            }
        }
    }
}

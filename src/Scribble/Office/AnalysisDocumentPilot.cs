using System;
using System.Collections.Generic;
using System.Linq;

namespace Scribble.Office
{
    // Development-only bridge to the existing Office writers. A caller must
    // provide a new, disposable destination; no model-facing tool uses this
    // bridge until the native, review and recovery gates have passed.
    public static class AnalysisDocumentPilot
    {
        public const string FeatureFlag = "SCRIBBLE_ANALYSIS_PILOT";

        public static string WriteWorkbook(object excelApplication,
            AnalysisArtifact artifact, AnalysisDocumentPlan plan)
        {
            RequireEnabled();
            var compiled = AnalysisDocumentCompiler.Compile(artifact, plan);
            var rows = compiled.WorkbookRows.Select(row =>
                (IReadOnlyList<string>)row).ToList();
            return WorkbookDraftWriter.WriteDraftSheet(excelApplication,
                compiled.WorkbookTitle, rows);
        }

        public static string WritePresentation(object powerPointApplication,
            AnalysisArtifact artifact, AnalysisDocumentPlan plan)
        {
            RequireEnabled();
            var compiled = AnalysisDocumentCompiler.Compile(artifact, plan);
            var slides = PresentationDraftWriter.ParseSlides(
                compiled.Slides.Cast<object>().ToArray());
            return PresentationDraftWriter.AddDraftSlides(powerPointApplication,
                slides, null, true);
        }

        private static void RequireEnabled()
        {
            if (!string.Equals(Environment.GetEnvironmentVariable(FeatureFlag),
                "1", StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "ANALYSIS_PILOT_DISABLED: The analysis writer is restricted to the development pilot.");
        }
    }
}

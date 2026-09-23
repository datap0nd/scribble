using System;
using System.Collections.Generic;
using System.Globalization;
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
            dynamic application = excelApplication;
            dynamic workbook = application.ActiveWorkbook;
            if (workbook == null)
                throw new InvalidOperationException(
                    "ANALYSIS_PILOT_DESTINATION_MISSING: Open a disposable source workbook first.");
            string[] before = SheetNames((object)workbook);
            var rows = compiled.WorkbookRows.Select(row =>
                (IReadOnlyList<string>)row).ToList();
            var status = WorkbookDraftWriter.WriteDraftSheet(excelApplication,
                compiled.WorkbookTitle, rows, null, false, (object)workbook);
            string[] after = SheetNames((object)workbook);
            var added = after.Except(before,
                StringComparer.OrdinalIgnoreCase).ToArray();
            if (added.Length != 1 || !added[0].StartsWith(
                WorkbookDraftWriter.DraftSheetName, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(
                    "ANALYSIS_PILOT_DRAFT_IDENTITY_INVALID: Expected one new marked draft sheet.");
            dynamic sheet = workbook.Worksheets[added[0]];
            var facts = artifact.Facts.ToDictionary(fact => fact.FactId,
                StringComparer.Ordinal);
            foreach (var expected in compiled.ExpectedFormulaFacts)
            {
                dynamic cell = sheet.Range(expected.Key);
                var formula = Convert.ToString(cell.Formula,
                    CultureInfo.InvariantCulture) ?? string.Empty;
                object native = cell.Value2;
                decimal actual = 0m;
                if (!formula.StartsWith("=", StringComparison.Ordinal) ||
                    native == null ||
                    !decimal.TryParse(Convert.ToString(native,
                        CultureInfo.InvariantCulture), NumberStyles.Float,
                        CultureInfo.InvariantCulture, out actual))
                    throw new InvalidOperationException(
                        "ANALYSIS_PILOT_FORMULA_UNRESOLVED: " + expected.Key);
                var wanted = AnalysisContract.Decimal(facts[expected.Value]);
                var tolerance = facts[expected.Value].Unit == "currency"
                    ? 0.005m : 0.000001m;
                if (Math.Abs(actual - wanted) > tolerance)
                    throw new InvalidOperationException(
                        "ANALYSIS_PILOT_FORMULA_MISMATCH: " + expected.Key +
                        " expected " + wanted.ToString(CultureInfo.InvariantCulture) +
                        " but Excel returned " + actual.ToString(CultureInfo.InvariantCulture));
            }
            return status + " Verified " + compiled.ExpectedFormulaFacts.Count +
                " live formula result(s) against the analysis.";
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

        private static string[] SheetNames(dynamic workbook)
        {
            var names = new List<string>();
            foreach (dynamic sheet in workbook.Worksheets)
                names.Add(Convert.ToString(sheet.Name,
                    CultureInfo.InvariantCulture));
            return names.ToArray();
        }
    }
}

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;

namespace Scribble.Office
{
    // Model-authored narrative and layout only. Workbook rows and all live
    // formulas are derived by the host from the retained analysis.
    public static class AnalysisSlidePlanContract
    {
        public const int MaxArgumentCharacters = 60000;

        public static AnalysisDocumentPlan Parse(AnalysisArtifact artifact,
            string json)
        {
            AnalysisContract.Serialize(artifact);
            if (string.IsNullOrWhiteSpace(json) ||
                json.Length > MaxArgumentCharacters)
                throw new InvalidOperationException(
                    "ANALYSIS_PLAN_PAYLOAD_INVALID");
            var serializer = new JavaScriptSerializer
                { MaxJsonLength = MaxArgumentCharacters };
            IDictionary<string, object> raw;
            try { raw = serializer.DeserializeObject(json) as
                IDictionary<string, object>; }
            catch (ArgumentException)
            {
                throw new InvalidOperationException(
                    "ANALYSIS_PLAN_PAYLOAD_INVALID");
            }
            CheckKeys(raw, "plan", "AnalysisId", "WorkbookTitle", "Slides");
            if (!string.Equals(Value(raw, "AnalysisId"),
                    artifact.AnalysisId, StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "ANALYSIS_PLAN_BINDING_INVALID");
            var slides = Items(raw, "Slides", 1,
                PresentationDraftWriter.MaxDraftSlides);
            foreach (var slide in slides)
            {
                var fields = Map(slide, "slide");
                CheckKeys(fields, "slide", "Id", "Layout", "Title",
                    "Subtitle", "Takeaway", "TableHeaders", "TableRows",
                    "Chart", "Cards");
                TextParts(fields, "Subtitle");
                TextParts(fields, "Takeaway");
                object rows;
                if (fields.TryGetValue("TableRows", out rows))
                    foreach (var row in Array(rows, 40, "TableRows"))
                    {
                        var map = Map(row, "table row");
                        CheckKeys(map, "table row", "Cells");
                        foreach (var cell in Items(map, "Cells", 1, 30))
                        {
                            var cellMap = Map(cell, "table cell");
                            CheckKeys(cellMap, "table cell", "Text", "FactId",
                                "Formula", "ExpectedFactId");
                            object formula;
                            object expectedFact;
                            if ((cellMap.TryGetValue("Formula", out formula) &&
                                    formula != null) ||
                                (cellMap.TryGetValue("ExpectedFactId",
                                    out expectedFact) && expectedFact != null))
                                throw new InvalidOperationException(
                                    "ANALYSIS_PLAN_FORMULA_FORBIDDEN");
                        }
                    }
                object chart;
                if (fields.TryGetValue("Chart", out chart) && chart != null)
                {
                    var map = Map(chart, "chart");
                    CheckKeys(map, "chart", "Type", "Title",
                        "Categories", "Series");
                    foreach (var series in Items(map, "Series", 1, 8))
                        CheckKeys(Map(series, "chart series"),
                            "chart series", "Name", "FactIds");
                }
                object cards;
                if (fields.TryGetValue("Cards", out cards))
                    foreach (var card in Array(cards, 8, "Cards"))
                    {
                        var map = Map(card, "card");
                        CheckKeys(map, "card", "Heading", "Points");
                        TextParts(map, "Points");
                    }
            }
            AnalysisDocumentPlan plan;
            try { plan = serializer.Deserialize<AnalysisDocumentPlan>(json); }
            catch (ArgumentException)
            {
                throw new InvalidOperationException(
                    "ANALYSIS_PLAN_PAYLOAD_INVALID");
            }
            plan.WorkbookRows = AnalysisWorkbookPlanBuilder.Build(artifact);
            AnalysisDocumentCompiler.Compile(artifact, plan);
            return plan;
        }

        private static void TextParts(IDictionary<string, object> parent,
            string key)
        {
            object raw;
            if (!parent.TryGetValue(key, out raw)) return;
            foreach (var part in Array(raw, 24, key))
                CheckKeys(Map(part, key + " item"), key + " item",
                    "Text", "FactId", "IncludeUnit");
        }

        private static IDictionary<string, object> Map(object value,
            string path)
        {
            var map = value as IDictionary<string, object>;
            if (map == null)
                throw new InvalidOperationException(
                    "ANALYSIS_PLAN_SHAPE_INVALID: " + path);
            return map;
        }

        private static object[] Items(IDictionary<string, object> parent,
            string key, int minimum, int maximum)
        {
            object value;
            if (!parent.TryGetValue(key, out value))
                throw new InvalidOperationException(
                    "ANALYSIS_PLAN_SHAPE_INVALID: " + key);
            var items = Array(value, maximum, key);
            if (items.Length < minimum)
                throw new InvalidOperationException(
                    "ANALYSIS_PLAN_SHAPE_INVALID: " + key);
            return items;
        }

        private static object[] Array(object value, int maximum,
            string path)
        {
            var items = value as IList;
            if (items == null || items.Count > maximum)
                throw new InvalidOperationException(
                    "ANALYSIS_PLAN_SHAPE_INVALID: " + path);
            return items.Cast<object>().ToArray();
        }

        private static string Value(IDictionary<string, object> map,
            string key)
        {
            object value;
            return map.TryGetValue(key, out value) ? value as string : null;
        }

        private static void CheckKeys(IDictionary<string, object> map,
            string path, params string[] allowed)
        {
            if (map == null || map.Keys.Except(allowed,
                    StringComparer.Ordinal).Any())
                throw new InvalidOperationException(
                    "ANALYSIS_PLAN_FIELD_UNSUPPORTED: " + path);
        }
    }
}

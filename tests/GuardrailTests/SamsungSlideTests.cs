using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Web.Script.Serialization;
using Scribble.Office;

namespace GuardrailTests
{
    internal static class SamsungSlideTests
    {
        public static void LayoutsAndOverflow()
        {
            var json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            var previews = new List<object>();
            foreach (var layout in SamsungSlideDesign.Layouts)
            {
                foreach (var region in SamsungSlideDesign.Regions(layout))
                    if (!SamsungSlideDesign.InBounds(region)) throw new Exception("Out-of-bounds recipe " + layout);
                var slide = new Dictionary<string, object> { { "layout", layout }, { "title", "Performance review" }, { "subtitle", "Demand supports the plan" } };
                if (new[] { "cards", "scorecard", "roadmap", "stack", "action_list" }.Contains(layout))
                    slide["cards"] = new[] { new { heading = "Prepare", points = new[] { "Review evidence", "This week" } }, new { heading = "Execute", points = new[] { "Apply changes", "Next week" } } };
                else if (!new[] { "cover", "divider", "closing" }.Contains(layout))
                    slide["bullets"] = new[] { "Review the evidence", "Confirm the next action" };
                var inspected = SamsungPresentationReview.InspectPlan(json.Serialize(new[] { slide }));
                previews.Add(inspected);
                var plan = json.Serialize(inspected);
                if (!plan.Contains("[Scribble draft]")) throw new Exception("Missing draft marker " + layout);
            }
            var rows = Enumerable.Range(1, 40).Select(i => new[] { "Item " + i, "100" }).ToArray();
            var denseRows = Enumerable.Range(1, 20).Select(i => new[] { "Item " + i, "100", "100", "100", "100", "100", "100" }).ToArray();
            var dense = SamsungPresentationReview.InspectPlan(json.Serialize(new[] { new { title = "Specification comparison", layout = "matrix", subtitle = "Compare the complete specification", table = new { headers = new[] { "Item", "A", "B", "C", "D", "E", "F" }, rows = denseRows } } }));
            if (((IEnumerable)dense).Cast<object>().Count() != 1) throw new Exception("The reference 20-row comparison should fit on one dense slide.");
            previews.Add(dense);
            var twoPane = SamsungPresentationReview.InspectPlan(json.Serialize(new[] { new { title = "Operational plan", layout = "two_pane", subtitle = "Align actions with targets", cards = new[] { new { heading = "Actions", points = new[] { "Confirm channel coverage", "Review the weekly forecast" } } }, table = new { headers = new[] { "Segment", "Target" }, rows = new[] { new[] { "Enterprise", "100" }, new[] { "Consumer", "200" } } }, secondary_table = new { headers = new[] { "Channel", "Share" }, rows = new[] { new[] { "Direct", "40%" }, new[] { "Partner", "60%" } } } } }));
            var paneJson = json.Serialize(twoPane);
            if (!paneJson.Contains("#4F81BD") || !paneJson.Contains("#F2F2F2")) throw new Exception("The two-pane recipe lost its heading and container.");
            previews.Add(twoPane);
            var scorecard = SamsungPresentationReview.InspectPlan(json.Serialize(new[] { new { title = "June performance", layout = "scorecard", subtitle = "Revenue softened while cost held flat", cards = new[] {
                new { heading = "June revenue", points = new[] { "EUR 82,992", "2.95% below May" } },
                new { heading = "June cost", points = new[] { "EUR 36,714", "Flat versus May" } },
                new { heading = "Gross margin", points = new[] { "55.76%", "1.32 points below May" } } } } }));
            var scorecardJson = json.Serialize(scorecard);
            if (!scorecardJson.Contains("EUR 82,992") || !scorecardJson.Contains("\"size\":34") || !scorecardJson.Contains("#596674"))
                throw new Exception("The scorecard recipe lost its prominent KPI hierarchy.");
            var scorecardPage = ((IEnumerable)json.DeserializeObject(scorecardJson)).Cast<Dictionary<string, object>>().Single();
            var scorecardElements = ((IEnumerable)scorecardPage["elements"]).Cast<Dictionary<string, object>>().ToArray();
            if (scorecardElements.Where(e => new[] { "JUNE REVENUE", "JUNE COST", "GROSS MARGIN" }.Contains(Convert.ToString(e["text"])))
                    .Any(e => Convert.ToDouble(e["size"]) < 14))
                throw new Exception("Scorecard KPI labels must meet the native 14-point body minimum.");
            var palette = new HashSet<string>(new[] { "#4F81BD", "#5B9BD5", "#41719C", "#F2F2F2", "#C00000", "#00B050",
                "#FFFFFF", "#000000", "#7F7F7F", "#A6A6A6", "#202A35", "#596674", "#D7DDE3", "#D4D4D4" }, StringComparer.OrdinalIgnoreCase);
            if (scorecardElements.Any(e => new[] { Convert.ToString(e["fill"]), Convert.ToString(e["color"]) }
                    .Any(color => !string.IsNullOrEmpty(color) && !palette.Contains(color))))
                throw new Exception("Scorecard colors must stay inside the supplied Samsung palette.");
            var executiveTable = SamsungPresentationReview.InspectPlan(json.Serialize(new[] { new {
                title = "June results", layout = "table", subtitle = "West leads on margin",
                table = new { headers = new[] { "Group", "Revenue", "Margin" }, rows = new[] {
                    new[] { "North", "19,219", "57.95%" }, new[] { "South", "22,675", "52.43%" },
                    new[] { "East", "19,054", "49.05%" }, new[] { "West", "22,044", "63.09%" },
                    new[] { "All groups", "82,992", "55.76%" } } } } }));
            var executivePage = ((IEnumerable)json.DeserializeObject(json.Serialize(executiveTable))).Cast<Dictionary<string, object>>().Single();
            var executiveGrid = ((IEnumerable)executivePage["elements"]).Cast<Dictionary<string, object>>()
                .Single(e => Convert.ToInt32(e["tableRows"]) == 5);
            if (Convert.ToDouble(executiveGrid["height"]) < 260)
                throw new Exception("A short executive table must occupy enough vertical space to avoid a small floating spreadsheet strip.");
            var evidenceCards = SamsungPresentationReview.InspectPlan(json.Serialize(new[] { new {
                title = "Data quality", layout = "cards", subtitle = "Source coverage is complete",
                cards = new[] {
                    new { heading = "Coverage", points = new[] { "144 rows across six months", "24 rows in June" } },
                    new { heading = "Integrity", points = new[] { "No blank revenue or cost cells", "Each ID counted once" } },
                    new { heading = "Method", points = new[] { "Margin uses aggregate totals" } }
                } } }));
            var evidenceJson = json.Serialize(evidenceCards);
            if (!evidenceJson.Contains("Coverage") || !evidenceJson.Contains("Integrity") ||
                !evidenceJson.Contains("Method") || !evidenceJson.Contains("#202A35"))
                throw new Exception("Evidence cards must render as distinct readable typographic columns.");
            var evidencePage = ((IEnumerable)json.DeserializeObject(evidenceJson)).Cast<Dictionary<string, object>>().Single();
            if (((IEnumerable)evidencePage["elements"]).Cast<Dictionary<string, object>>()
                .Any(e => new[] { Convert.ToString(e["fill"]), Convert.ToString(e["color"]) }
                    .Any(color => !string.IsNullOrEmpty(color) && !palette.Contains(color))))
                throw new Exception("Evidence-card colors must stay inside the supplied Samsung palette.");
            SamsungAuthoringPolicy.ValidateVisualDesign(new[] { new Dictionary<string, object> { { "title", "June performance" }, { "layout", "scorecard" }, { "cards", new object[] { new { heading = "Revenue", points = new[] { "82,992" } }, new { heading = "Margin", points = new[] { "55.76%" } } } } } });
            ExpectFailure(() => SamsungAuthoringPolicy.ValidateVisualDesign(new[] { new Dictionary<string, object> { { "title", "June performance" }, { "layout", "bullets" }, { "purpose", "analytical" }, { "bullets", new[] { "Revenue 82,992", "Cost 36,714", "Margin 55.76%" } } } }));
            ExpectFailure(() => SamsungAuthoringPolicy.ValidateVisualDesign(new[] { new Dictionary<string, object> { { "title", "Data Quality and Methodology" }, { "layout", "bullets" }, { "purpose", "methodology" }, { "bullets", new[] { "144 source rows", "0 missing inputs", "24 rows per period", "Each ID counted once", "Rates use aggregate totals", "No imputation" } } } }));
            ExpectFailure(() => SamsungAuthoringPolicy.ValidateVisualDesign(new[] { new Dictionary<string, object> { { "title", "Data Quality and Methodology" }, { "layout", "action_list" }, { "purpose", "explanatory" }, { "bullets", new[] { "144 source rows", "0 missing inputs", "24 rows per period", "No duplicates" } } } }));
            ExpectFailure(() => SamsungAuthoringPolicy.ValidateVisualDesign(new[] { new Dictionary<string, object> { { "title", "Data Quality and Methodology" }, { "layout", "bullets" }, { "bullets", new[] { "144 source rows", "0 missing inputs", "24 rows per period", "No duplicates" } }, { "cards", new object[] { new { heading = "Coverage" }, new { heading = "Integrity" } } } } }));
            var otherSlideVerdict = "{\"approved\":false,\"issues\":\"Slide 2 is missing.\",\"findings\":[{\"slide_id\":\"slide-2\",\"severity\":\"blocker\",\"type\":\"coverage\",\"correction\":\"Add the period comparison slide.\"}]}";
            if (!SamsungAuthoringPolicy.OnlyOtherSlideCoverageBlockers(otherSlideVerdict, "headline") ||
                SamsungAuthoringPolicy.OnlyOtherSlideCoverageBlockers(otherSlideVerdict, "slide-2"))
                throw new Exception("A single-slide reviewer must not block on an unrelated planned slide.");
            var revisionType = typeof(SamsungAuthoringPolicy).Assembly.GetType("Scribble.Office.PresentationRevision", true);
            var overflow = revisionType.GetMethod("NativeTextOverflows",
                System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic);
            if ((bool)overflow.Invoke(null, new object[] { "", 18f, 20f, 4f, 100f }) ||
                !(bool)overflow.Invoke(null, new object[] { "Visible", 18f, 20f, 4f, 100f }))
                throw new Exception("Empty accent shapes must not fail native text-fit validation.");
            ExpectFailure(() => SamsungAuthoringPolicy.ValidateDeckVisualDesign(new[] {
                new Dictionary<string, object> { { "title", "One" }, { "layout", "bullets" }, { "purpose", "explanatory" }, { "bullets", new[] { "Context" } } },
                new Dictionary<string, object> { { "title", "Two" }, { "layout", "bullets" }, { "purpose", "explanatory" }, { "bullets", new[] { "Context" } } },
                new Dictionary<string, object> { { "title", "Three" }, { "layout", "table" }, { "table", new Dictionary<string, object> { { "rows", new[] { new[] { "A", "1" } } } } } }
            }));
            const string captionText = "June 2026 records only. Margin uses aggregate revenue and cost, never an average of row rates.";
            var captionPlan = SamsungPresentationReview.InspectPlan(json.Serialize(new[] {
                new { title = "June results", subtitle = "Revenue held firm", layout = "table", caption = captionText, unit = "(EUR; % where shown)",
                    table = new { headers = new[] { "Group", "Revenue EUR" }, rows = new[] { new[] { "North", "19,219" } } } } }));
            var captioned = (IEnumerable)json.DeserializeObject(json.Serialize(captionPlan));
            var captionPage = (Dictionary<string, object>)captioned.Cast<object>().Single();
            var caption = ((IEnumerable)captionPage["elements"]).Cast<Dictionary<string, object>>()
                .Single(element => Convert.ToString(element["text"]) == captionText);
            if (Convert.ToDouble(caption["size"]) < 14 || Convert.ToDouble(caption["minimum"]) < 14 || Convert.ToDouble(caption["width"]) < 610)
                throw new Exception("Samsung analytical captions must remain readable at the native presentation minimum.");
            const string chartCaveat = "June source workbook, Summary!B4:C5";
            var chartPlan = SamsungPresentationReview.InspectPlan(json.Serialize(new[] { new {
                title = "Revenue by region", subtitle = "West leads", layout = "chart", unit = "EUR",
                caption = "Native editable chart; single primary series; value axis from 0",
                footnote = "Primary values only; value axis begins at zero; " + chartCaveat,
                sources = "June source workbook", chart = new {
                    categories = new[] { "North", "South", "West" },
                    series = new[] { new { name = "Revenue", values = new[] { 10, 12, 16 } } }
                } } }));
            var chartPage = ((IEnumerable)json.DeserializeObject(json.Serialize(chartPlan)))
                .Cast<Dictionary<string, object>>().Single();
            var chartElements = ((IEnumerable)chartPage["elements"]).Cast<Dictionary<string, object>>().ToArray();
            var visibleCopy = string.Join(" ", chartElements.Select(e => Convert.ToString(e["text"])));
            if (visibleCopy.Contains("Native editable chart") || visibleCopy.Contains("single primary series") ||
                visibleCopy.Contains("value axis") || visibleCopy.Contains("Primary values only"))
                throw new Exception("Chart implementation instructions leaked onto the audience-facing canvas.");
            if (!visibleCopy.Contains(chartCaveat) || !visibleCopy.Contains("June source workbook"))
                throw new Exception("Filtering chart instructions removed a factual source or caveat.");
            if (!chartElements.Any(e => Convert.ToString(e["text"]) == "EUR" && Convert.ToDouble(e["size"]) >= 11))
                throw new Exception("Chart units must be legible in the native presentation.");
            var pages = (IEnumerable)SamsungPresentationReview.InspectPlan(json.Serialize(new[] { new { title = "Data", layout = "matrix", subtitle = "Review every row", table = new { headers = new[] { "Item", "Value" }, rows } } }));
            previews.Add(pages);
            File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SamsungPlans.json"), json.Serialize(previews));
            var data = (IEnumerable)json.DeserializeObject(json.Serialize(pages));
            var covered = 0; var count = 0;
            foreach (Dictionary<string, object> page in data)
            {
                count++;
                foreach (Dictionary<string, object> element in (IEnumerable)page["elements"])
                    covered += Convert.ToInt32(element["tableRows"]);
            }
            if (count < 2 || covered != 40) throw new Exception("Table pagination dropped or duplicated rows.");
            ExpectFailure(() => SamsungSlideDesign.Fit(new string('X', 5000), "Arial", new RectangleF(0, 0, 100, 20), 18, 14));
            ExpectFailure(() => SamsungPresentationReview.InspectPlan("[{\"title\":\"Bad chart\",\"chart\":{\"categories\":[\"A\",\"B\"],\"series\":[{\"name\":\"Sales\",\"values\":[1]}]}}]"));
            if (SamsungSlideDesign.FontFor("한글", "Arial") != "Malgun Gothic") throw new Exception("Korean font mapping lost.");
            if (!SamsungSlideDesign.SameOwner("ABCDEF", "abcdef") || SamsungSlideDesign.SameOwner("", "")) throw new Exception("PowerPoint tag normalization broke ownership checks.");
        }
        public static void EvidenceAndNumbers()
        {
            SamsungPresentationReview.ValidatePlan(new[] { "intro", "evidence", "decision" }, new[] { "evidence" }, new[] { "intro" });
            ExpectFailure(() => SamsungPresentationReview.ValidatePlan(new[] { "intro", "intro" }, new[] { "intro" }, new string[0]));
            ExpectFailure(() => SamsungPresentationReview.ValidatePlan(new[] { "intro", "decision" }, new[] { "decision" }, new string[0]));
            ExpectFailure(() => SamsungPresentationReview.ValidatePlan(new[] { "intro" }, new[] { "intro" }, new[] { "intro" }));
            var json = new JavaScriptSerializer();
            var slide = new { title = "Sales increased 20%", subtitle = "Sales increased 20%", sources = "Report, page 1", evidence = "Sales increased 20%." };
            SamsungPresentationReview.ValidateEvidence(json.Serialize(slide), "Source: Sales increased 20%.");
            SamsungPresentationReview.ValidateEvidence(json.Serialize(new { title = "Value 1000", subtitle = "Rate 0.4", sources = "Report", evidence = "Value 1,000.0; rate .4" }), "Value 1,000.0; rate .4");
            SamsungPresentationReview.ValidateEvidence(json.Serialize(new { title = "Next actions", subtitle = "Prepare the launch", sources = "Report", evidence = "Prepare the launch. Review results.",
                bullets = new[] { "1. Prepare the launch", "2. Review results" } }), "Prepare the launch.\n  Review results.");
            ExpectFailure(() => SamsungPresentationReview.ValidateEvidence(json.Serialize(new { title = "Next actions", subtitle = "Prepare", sources = "Report", evidence = "Prepare the launch.",
                bullets = new[] { "2026. Launch with 500 units" } }), "Prepare the launch."));
            ExpectFailure(() => SamsungPresentationReview.ValidateEvidence(json.Serialize(new { layout = "cover", title = "500 units", evidence = "Launch" }), "Launch"));
            ExpectFailure(() => SamsungPresentationReview.ValidateEvidence(json.Serialize(slide), "No such result"));
            ExpectFailure(() => SamsungPresentationReview.ValidateEvidence(json.Serialize(new { title = "Sales increased 30%", subtitle = "Growth", sources = "Report", evidence = "Sales increased 20%." }), "Sales increased 20%."));
        }
        private static void ExpectFailure(Action action)
        { try { action(); } catch (InvalidOperationException) { return; } throw new Exception("Expected a concrete source/layout blocker."); }
    }
}

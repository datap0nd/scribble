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
            var comparison = SamsungPresentationReview.InspectPlan(json.Serialize(new[] { new {
                title = "Two period comparison", layout = "two_pane",
                table = new { headers = new[] { "Metric", "May", "June" }, rows = new[] {
                    new[] { "Revenue EUR", "85,519", "82,992" },
                    new[] { "Cost EUR", "36,702", "36,714" } } },
                chart = new { type = "column", title = "Revenue EUR",
                    categories = new[] { "May", "June" }, series = new[] {
                        new { name = "Revenue EUR", values = new[] { 85519, 82992 } } } }
            } }));
            var comparisonPage = ((IEnumerable)json.DeserializeObject(json.Serialize(comparison)))
                .Cast<Dictionary<string, object>>().Single();
            var comparisonElements = ((IEnumerable)comparisonPage["elements"])
                .Cast<Dictionary<string, object>>().ToArray();
            var comparisonTable = comparisonElements.Single(e =>
                Convert.ToInt32(e["tableRows"]) == 2);
            var comparisonChart = comparisonElements.Single(e =>
                Convert.ToBoolean(e["chart"]));
            if (Convert.ToDouble(comparisonChart["width"]) < 400 ||
                Convert.ToDouble(comparisonChart["height"]) < 260 ||
                Convert.ToDouble(comparisonTable["height"]) > 250 ||
                Convert.ToDouble(comparisonTable["width"]) >
                    Convert.ToDouble(comparisonChart["width"]))
                throw new Exception("A table/chart comparison must give the native chart and table balanced, legible regions.");
            previews.Add(comparison);
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
            var assurance = SamsungPresentationReview.InspectPlan(json.Serialize(new[] { new {
                title = "What the ledger supports", layout = "cards",
                cards = new[] { new { heading = "Evidence",
                    points = new[] { "Verified workbook range and live formulas" } } }
            } }));
            var assuranceElements = ((IEnumerable)((IEnumerable)json.DeserializeObject(
                json.Serialize(assurance))).Cast<Dictionary<string, object>>()
                .Single()["elements"]).Cast<Dictionary<string, object>>().ToArray();
            if (assuranceElements.Any(e => Convert.ToString(e["fill"]) ==
                    SamsungSlideDesign.Gray) ||
                !assuranceElements.Any(e => Convert.ToString(e["text"]) ==
                    "Verified workbook range and live formulas" &&
                    Convert.ToDouble(e["size"]) >= 24))
                throw new Exception("A short single-card assurance needs a prominent statement rather than a sparse gray form panel.");
            previews.Add(assurance);
            var evidenceJson = json.Serialize(evidenceCards);
            if (!evidenceJson.Contains("Coverage") || !evidenceJson.Contains("Integrity") ||
                !evidenceJson.Contains("Method") || !evidenceJson.Contains("#202A35"))
                throw new Exception("Evidence cards must render as distinct readable typographic columns.");
            var evidencePage = ((IEnumerable)json.DeserializeObject(evidenceJson)).Cast<Dictionary<string, object>>().Single();
            var evidenceElements = ((IEnumerable)evidencePage["elements"]).Cast<Dictionary<string, object>>().ToArray();
            var cardPanels = evidenceElements.Where(e => Convert.ToString(e["fill"]) == SamsungSlideDesign.Gray &&
                Convert.ToDouble(e["height"]) > 100).ToArray();
            if (cardPanels.Length != 3 || cardPanels.Any(e => Convert.ToDouble(e["height"]) > 260))
                throw new Exception("Short evidence cards must not leave a half-empty full-height gray panel.");
            var qualityPlan = SamsungPresentationReview.InspectPlan(json.Serialize(new[] { new {
                title = "Data quality and evidence limits", layout = "cards", subtitle = "Each ID counted once",
                cards = new[] {
                    new { heading = "Coverage", points = new[] { "Ledger: 144 records, 2026-01 to 2026-06", "History through May", "Four groups", "Revenue and cost in EUR" } },
                    new { heading = "Integrity", points = new[] { "144 distinct RowIDs", "0 duplicates", "0 blank revenue cells", "0 blank cost cells" } },
                    new { heading = "Limits", points = new[] { "May reconciles: 85,519 / 36,702", "June missing from History", "June uses Ledger", "Never set blanks to zero" } }
                } } }));
            var qualityPage = ((IEnumerable)json.DeserializeObject(json.Serialize(qualityPlan)))
                .Cast<Dictionary<string, object>>().Single();
            var qualityElements = ((IEnumerable)qualityPage["elements"]).Cast<Dictionary<string, object>>().ToArray();
            if (qualityElements.Count(e => Convert.ToDouble(e["size"]) >= 26 &&
                    System.Text.RegularExpressions.Regex.IsMatch(Convert.ToString(e["text"]), @"\d")) < 2)
                throw new Exception("Numeric evidence cards need at least two prominent source-backed metrics.");
            var incidentalNumber = SamsungPresentationReview.InspectPlan(json.Serialize(new[] { new {
                title = "Data-quality limits", layout = "cards", subtitle = "Ledger is complete",
                cards = new[] {
                    new { heading = "Coverage", points = new[] { "24 rows in June", "Four groups" } },
                    new { heading = "Data integrity", points = new[] { "Revenue and Cost complete", "WB01-0104 Units/UnitCost row full (33 / 50)" } },
                    new { heading = "Method", points = new[] { "Rates use aggregates" } },
                    new { heading = "Limits", points = new[] { "Planned follow-ups are not results" } }
                } } }));
            var incidentalElements = ((IEnumerable)((IEnumerable)json.DeserializeObject(json.Serialize(incidentalNumber)))
                .Cast<Dictionary<string, object>>().Single()["elements"]).Cast<Dictionary<string, object>>();
            if (incidentalElements.Any(e => Convert.ToDouble(e["size"]) >= 26 &&
                    new[] { "33", "50", "0104" }.Contains(Convert.ToString(e["text"]))))
                throw new Exception("An incidental source-row value was promoted as the Data integrity KPI.");
            if (((IEnumerable)evidencePage["elements"]).Cast<Dictionary<string, object>>()
                .Any(e => new[] { Convert.ToString(e["fill"]), Convert.ToString(e["color"]) }
                    .Any(color => !string.IsNullOrEmpty(color) && !palette.Contains(color))))
                throw new Exception("Evidence-card colors must stay inside the supplied Samsung palette.");
            SamsungAuthoringPolicy.ValidateVisualDesign(new[] { new Dictionary<string, object> { { "title", "June performance" }, { "layout", "scorecard" }, { "cards", new object[] { new { heading = "Revenue", points = new[] { "82,992" } }, new { heading = "Margin", points = new[] { "55.76%" } } } } } });
            SamsungAuthoringPolicy.ValidateVisualDesign(new[] { new Dictionary<string, object> {
                { "title", "Source-backed cover" }, { "layout", "cover" },
                { "sources", new string('S', 120) }, { "footnote", "Fictional source" } } });
            ExpectFailure(() => SamsungAuthoringPolicy.ValidateVisualDesign(new[] { new Dictionary<string, object> {
                { "title", "Source-backed cover" }, { "layout", "cover" },
                { "sources", "WB01" }, { "footnote", new string('F', 120) } } }));
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
                    footnote = "Rates use aggregate revenue and cost, never row averages; each RowID is counted once; missing inputs remain unknown rather than zero; the visible numbers are rounded only for display.",
                    sources = "Source: WB01 Ledger",
                    table = new { headers = new[] { "Group", "Revenue EUR" }, rows = new[] { new[] { "North", "19,219" } } } } }));
            var captioned = (IEnumerable)json.DeserializeObject(json.Serialize(captionPlan));
            var captionPage = (Dictionary<string, object>)captioned.Cast<object>().Single();
            var caption = ((IEnumerable)captionPage["elements"]).Cast<Dictionary<string, object>>()
                .Single(element => Convert.ToString(element["text"]) == captionText);
            if (Convert.ToDouble(caption["size"]) < 14 || Convert.ToDouble(caption["minimum"]) < 14 || Convert.ToDouble(caption["width"]) < 610)
                throw new Exception("Samsung analytical captions must remain readable at the native presentation minimum.");
            var captionElements = ((IEnumerable)captionPage["elements"]).Cast<Dictionary<string, object>>().ToArray();
            if (!captionElements.Any(e => Convert.ToString(e["text"]) == "Source: WB01 Ledger") ||
                captionElements.Any(e => Convert.ToString(e["text"]).Contains("full evidence in speaker notes")))
                throw new Exception("A long caveat must not displace the audience-facing citation with workflow copy.");
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
            if (SamsungAuthoringPolicy.AudienceChartTitle("Revenue EUR by month (zero-based axis)") != "Revenue EUR by month")
                throw new Exception("Chart construction checks leaked into the native audience-facing title.");
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
            // Six source passages can exceed the old 12k evidence limit even
            // though the native slide itself remains compact and well formed.
            SamsungPresentationReview.InspectPlan(json.Serialize(new[] { new {
                title = "Monthly comparison", layout = "cover", evidence = new string('E', 18000) } }));
            if (SamsungSlideDesign.FontFor("한글", "Arial") != "Malgun Gothic") throw new Exception("Korean font mapping lost.");
            if (SamsungSlideDesign.FontFor("Planned ≠ completed", "Samsung Sharp Sans Bold") != "Arial")
                throw new Exception("PowerPoint comparison glyph can render as a hash in the title font.");
            if (!SamsungSlideDesign.SameOwner("ABCDEF", "abcdef") || SamsungSlideDesign.SameOwner("", "")) throw new Exception("PowerPoint tag normalization broke ownership checks.");
        }
        public static void EvidenceAndNumbers()
        {
            var citation = PresentationInspection.CitationTextFromCaptured(new Dictionary<string, object> {
                { "shapes", new object[] {
                    new Dictionary<string, object> { { "text", "June revenue eur: 82,992.\r\rCost EUR: 36,714." } },
                    new Dictionary<string, object> { { "table", new object[] {
                        new object[] { new Dictionary<string, object> { { "text", "Group" } }, new Dictionary<string, object> { { "text", "Revenue EUR" } } },
                        new object[] { new Dictionary<string, object> { { "text", "North" } }, new Dictionary<string, object> { { "text", "19,219" } } } } } } } },
                { "notes", "January–June 2026" } });
            if (!citation.Contains("June revenue eur: 82,992.") || !citation.Contains("Cost EUR: 36,714.") ||
                !citation.Contains("Group\tRevenue EUR\nNorth\t19,219") || citation.Contains("\\r"))
                throw new Exception("Native PowerPoint citation text lost decoded source values or table associations.");
            SamsungEvidence.ValidateClaims(new Dictionary<string, object> { { "claims", new object[] {
                new Dictionary<string, object> { { "text", "June revenue EUR 82,992" },
                    { "evidence", "June revenue eur: 82,992.\r\rCost EUR: 36,714." },
                    { "label", "Revenue EUR" }, { "unit", "EUR" }, { "period", "2026-06" } } } } }, citation);
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

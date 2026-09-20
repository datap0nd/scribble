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
            SamsungAuthoringPolicy.ValidateVisualDesign(new[] { new Dictionary<string, object> { { "title", "June performance" }, { "layout", "scorecard" }, { "cards", new object[] { new { heading = "Revenue", points = new[] { "82,992" } }, new { heading = "Margin", points = new[] { "55.76%" } } } } } });
            ExpectFailure(() => SamsungAuthoringPolicy.ValidateVisualDesign(new[] { new Dictionary<string, object> { { "title", "June performance" }, { "layout", "bullets" }, { "purpose", "analytical" }, { "bullets", new[] { "Revenue 82,992", "Cost 36,714", "Margin 55.76%" } } } }));
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

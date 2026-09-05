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
                if (new[] { "cards", "roadmap", "stack", "action_list" }.Contains(layout))
                    slide["cards"] = new[] { new { heading = "Prepare", points = new[] { "Review evidence", "This week" } }, new { heading = "Execute", points = new[] { "Apply changes", "Next week" } } };
                else if (!new[] { "cover", "divider", "closing" }.Contains(layout))
                    slide["bullets"] = new[] { "Review the evidence", "Confirm the next action" };
                var inspected = SamsungPresentationReview.InspectPlan(json.Serialize(new[] { slide }));
                previews.Add(inspected);
                var plan = json.Serialize(inspected);
                if (!plan.Contains("[Scribble draft]")) throw new Exception("Missing draft marker " + layout);
            }
            var rows = Enumerable.Range(1, 40).Select(i => new[] { "Item " + i, "100" }).ToArray();
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
            ExpectFailure(() => SamsungPresentationReview.ValidateEvidence(json.Serialize(slide), "No such result"));
            ExpectFailure(() => SamsungPresentationReview.ValidateEvidence(json.Serialize(new { title = "Sales increased 30%", subtitle = "Growth", sources = "Report", evidence = "Sales increased 20%." }), "Sales increased 20%."));
        }
        private static void ExpectFailure(Action action)
        { try { action(); } catch (InvalidOperationException) { return; } throw new Exception("Expected a concrete source/layout blocker."); }
    }
}

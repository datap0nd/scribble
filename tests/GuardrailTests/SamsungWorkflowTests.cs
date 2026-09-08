using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;
using Scribble.Chat;
using Scribble.Office;

namespace GuardrailTests
{
    internal static class SamsungWorkflowTests
    {
        private static void Check(bool value, string message) { if (!value) throw new Exception(message); }
        private static void Reject(Action action)
        { try { action(); } catch (InvalidOperationException) { return; } throw new Exception("Expected rejection."); }
        internal static void Evidence()
        {
            var json = new JavaScriptSerializer();
            const string source = "Sales Q1 100 units. Sales Q2 120 units.";
            var calc = new Dictionary<string, object> {
                { "label", "Sales growth" }, { "operation", "growth_percent" }, { "result", 20 }, { "unit", "%" }, { "decimals", 1 },
                { "operands", new object[] {
                    new Dictionary<string, object> { { "value", 120 }, { "label", "Sales" }, { "unit", "units" }, { "period", "Q2" }, { "evidence", "Sales Q2 120 units." } },
                    new Dictionary<string, object> { { "value", 100 }, { "label", "Sales" }, { "unit", "units" }, { "period", "Q1" }, { "evidence", "Sales Q1 100 units." } }
                } }
            };
            var slide = new Dictionary<string, object> { { "title", "Sales" }, { "subtitle", "Sales rose 20%" }, { "evidence", source }, { "sources", "Supplied report" }, { "calculations", new[] { calc } } };
            SamsungPresentationReview.ValidateEvidence(json.Serialize(slide), source);
            calc["result"] = 25; Reject(() => SamsungPresentationReview.ValidateEvidence(json.Serialize(slide), source));
            calc["result"] = 20; calc["unit"] = "units"; Reject(() => SamsungPresentationReview.ValidateEvidence(json.Serialize(slide), source)); calc["unit"] = "%";
            var operand = (Dictionary<string, object>)((object[])calc["operands"])[0]; operand["period"] = "Q3";
            Reject(() => SamsungPresentationReview.ValidateEvidence(json.Serialize(slide), source)); operand["period"] = "Q2";
            var factual = new Dictionary<string, object> { { "content_kind", "fact" }, { "source_spans", new[] { "source1" } } };
            Check(!SamsungPresentationReview.PrepareSampleEvidence(factual, "Use sample data for the example slide"), "Sample mode contaminated a factual slide.");
            var sample = new Dictionary<string, object> { { "content_kind", "sample" } };
            Check(SamsungPresentationReview.PrepareSampleEvidence(sample, "Use sample data: 100 units"), "Explicit sample mode rejected.");
            Check(!SamsungPresentationReview.PrepareSampleEvidence(new Dictionary<string, object>(), "Do not use sample data"), "Negative sample instruction ignored.");
            SamsungPresentationReview.ValidateEvidence(json.Serialize(new { title = "Definitions", purpose = "explanatory", evidence = "Sales means units shipped.", sources = "Glossary" }), "Sales means units shipped.");
            Reject(() => SamsungPresentationReview.ValidateEvidence(json.Serialize(new { title = "Sales", evidence = source, sources = "Report" }), source));
        }
        internal static void PolicyAndCompletion()
        {
            var definition = PresentationToolCatalog.DraftDefinition();
            Check(definition.function.description.Contains(SamsungAuthoringPolicy.Instructions), "Generation policy differs from shared policy.");
            Check(!definition.function.description.Contains("takeaway sentences as titles"), "Conflicting title rules remain.");
            var key = SamsungAuthoringPolicy.CacheKey("model", "endpoint", "slide", "evidence");
            Check(key != SamsungAuthoringPolicy.CacheKey("model", "endpoint", "slide", "changed evidence"), "Evidence did not invalidate review.");
            Check(!SamsungAuthoringPolicy.Approved("{\"approved\":true,\"findings\":[{\"severity\":\"blocker\"}]}"), "Blocker approved.");
            Check(!SamsungAuthoringPolicy.Approved("not json"), "Invalid review approved.");
            var state = new DurableTaskState { EnumerationComplete = true, PresentationReviewRequired = true };
            Check(!state.CanComplete(false), "Deck completed without a final review receipt.");
            state.PresentationReviewReceipt = "reviewed"; Check(state.CanComplete(false), "Valid final receipt rejected.");
            var briefs = new object[] { new Dictionary<string, object> { { "id", "a" }, { "purpose", "explanation" }, { "message", "Definitions" }, { "layout", "bullets" }, { "required_content", new[] { "definition" } } } };
            SamsungAuthoringPolicy.ValidateBriefs(briefs, new[] { "a" });
            Reject(() => SamsungAuthoringPolicy.ValidateBriefs(briefs, new[] { "b" }));
        }
        internal static void ChartGapsAndAnnotations()
        {
            var json = new JavaScriptSerializer();
            var slides = new[] { new { title = "Sales", subtitle = "A missing value is unknown", layout = "chart",
                chart = new { categories = new[] { "Q1", "Q2", "Q3" }, series = new[] { new { name = "Sales", values = new double?[] { 100, null, 120 } } } } } };
            var call = new ChatToolCall { id = "gaps", function = new ChatToolCallFunction { name = PresentationToolCatalog.AddDraftSlides, arguments = json.Serialize(new { slides }) } };
            Check(ToolContractValidator.Validate(call, PresentationToolCatalog.DraftDefinition()).Count == 0, "A native chart gap failed schema validation.");
            SamsungPresentationReview.InspectPlan(json.Serialize(slides));
            call.function.arguments = call.function.arguments.Replace("null", "\"unknown\"");
            Check(ToolContractValidator.Validate(call, PresentationToolCatalog.DraftDefinition()).Count > 0, "A string chart value bypassed nullable number validation.");
            var nullable = (Dictionary<string, object>)GeminiCodeAssistGateway.SanitizeSchema(new Dictionary<string, object> { { "type", new[] { "number", "null" } } });
            Check((string)nullable["type"] == "NUMBER" && (bool)nullable["nullable"], "Gemini lost the nullable numeric contract.");
            var annotated = new { title = "Comparison", subtitle = "Review the supporting row", layout = "dual_visual",
                table = new { headers = new[] { "Item", "Units" }, rows = new[] { new[] { "A", "100" } } },
                secondary_table = new { headers = new[] { "Item", "Units" }, rows = new[] { new[] { "B", "120" } } },
                annotations = new[] { new { target = "secondary_table", row = 1, column = 2 } } };
            var rendered = json.Serialize(SamsungPresentationReview.InspectPlan(json.Serialize(new[] { annotated })));
            Check(rendered.Contains("\"hollow\":true"), "Secondary evidence annotation was omitted.");
            Reject(() => SamsungPresentationReview.InspectPlan(json.Serialize(new[] { annotated }).Replace("\"row\":1", "\"row\":9")));
        }
    }
}

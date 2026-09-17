using System;
using System.Collections;
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
            const string aggregateSource = "June 2026 Gross margin 55.76%.";
            var directClaim = new Dictionary<string, object> {
                { "text", "Gross margin was 55.76%." }, { "label", "Gross margin" }, { "unit", "%" }, { "period", "June 2026" }, { "evidence", aggregateSource }
            };
            var directSlide = new Dictionary<string, object> { { "title", "June margin" }, { "subtitle", "Gross margin was 55.76%" },
                { "evidence", aggregateSource }, { "sources", "Management view" }, { "claims", new[] { directClaim } } };
            SamsungPresentationReview.ValidateEvidence(json.Serialize(directSlide), aggregateSource);
            const string isoPeriodSource = "2026-06 Gross margin 55.76%.";
            directClaim["evidence"] = isoPeriodSource;
            directSlide["evidence"] = isoPeriodSource;
            SamsungPresentationReview.ValidateEvidence(json.Serialize(directSlide), isoPeriodSource);
            directClaim["evidence"] = aggregateSource;
            directSlide["evidence"] = aggregateSource;
            const string adjacentHeaderSource = "Atlas Components | January–June 2026\nJune revenue eur: 82,992.";
            var adjacentClaim = new Dictionary<string, object> {
                { "text", "June revenue was 82,992 EUR." }, { "label", "Revenue EUR" }, { "unit", "EUR" },
                { "period", "June 2026" }, { "evidence", "June revenue eur: 82,992." }
            };
            var adjacentSlide = new Dictionary<string, object> { { "title", "June revenue" },
                { "subtitle", "June revenue was 82,992 EUR" }, { "evidence", adjacentHeaderSource },
                { "sources", "Management view" }, { "claims", new[] { adjacentClaim } } };
            SamsungPresentationReview.ValidateEvidence(json.Serialize(adjacentSlide), adjacentHeaderSource);
            directClaim["period"] = "May 2026";
            Reject(() => SamsungPresentationReview.ValidateEvidence(json.Serialize(directSlide), aggregateSource));
            try { SamsungPresentationReview.ValidateEvidence(json.Serialize(directSlide), aggregateSource); throw new Exception("Expected actionable association rejection."); }
            catch (InvalidOperationException ex) { Check(ex.Message.Contains("Gross margin was 55.76%") && ex.Message.Contains("May 2026") && ex.Message.Contains("June 2026 Gross margin 55.76%"), "Claim association failure did not identify the claim, missing period and rejected passage together."); }
            directClaim["period"] = "June 2026"; directClaim["evidence"] = "Gross margin was 55.76%.";
            try { SamsungPresentationReview.ValidateEvidence(json.Serialize(directSlide), aggregateSource); throw new Exception("Expected exact-passage rejection."); }
            catch (InvalidOperationException ex) { Check(ex.Message.Contains("exact verbatim") && ex.Message.Contains("Gross margin was 55.76%"), "Rejected association did not return actionable exact-passage guidance."); }
            const string auditTable = "Metric\tMay (History)\tJune (Ledger)\nRevenue EUR\t85519\t82992\nCost EUR\t36702\t36714";
            var repairedClaim = new Dictionary<string, object> {
                { "text", "June revenue was 82,992 EUR." }, { "label", "Revenue EUR" }, { "unit", "EUR" },
                { "period", "June (Ledger)" }, { "evidence", "2026-06\t82992\t36714" }
            };
            var repairedSlide = new Dictionary<string, object> { { "title", "June revenue" },
                { "subtitle", "June revenue was 82,992 EUR" }, { "evidence", auditTable },
                { "sources", "WB01" }, { "claims", new[] { repairedClaim } } };
            try { SamsungPresentationReview.ValidateEvidence(json.Serialize(repairedSlide), auditTable); throw new Exception("Expected fabricated-row rejection."); }
            catch (InvalidOperationException ex) {
                Check(ex.Message.Contains("nearby host-verified passage") &&
                    ex.Message.Contains("Metric May (History) June (Ledger) Revenue EUR 85519 82992 Cost EUR 36702 36714") &&
                    ex.Message.Contains("split comparisons into one claim per period"),
                    "Rejected table citation did not return the nearest verified header-and-row block.");
            }
            var factual = new Dictionary<string, object> { { "content_kind", "fact" }, { "source_spans", new[] { "source1" } } };
            Check(!SamsungPresentationReview.PrepareSampleEvidence(factual, "Use sample data for the example slide"), "Sample mode contaminated a factual slide.");
            var sample = new Dictionary<string, object> { { "content_kind", "sample" } };
            Check(SamsungPresentationReview.PrepareSampleEvidence(sample, "Use sample data: 100 units"), "Explicit sample mode rejected.");
            Check(!SamsungPresentationReview.PrepareSampleEvidence(new Dictionary<string, object>(), "Do not use sample data"), "Negative sample instruction ignored.");
            SamsungPresentationReview.ValidateEvidence(json.Serialize(new { title = "Definitions", purpose = "explanatory", evidence = "Sales means units shipped.", sources = "Glossary" }), "Sales means units shipped.");
            Reject(() => SamsungPresentationReview.ValidateEvidence(json.Serialize(new { title = "Sales", evidence = source, sources = "Report" }), source));
        }
        // XA01 regression: a gross margin is derived from two cited operands and
        // recomputed by the host; it is neither an unsupported number nor a
        // question for the user.
        internal static void DerivedMarginCalculations()
        {
            var json = new JavaScriptSerializer();
            const string auditTable = "Metric\tMay\tJune\nRevenue EUR\t85519\t82992\nCost EUR\t36702\t36714";
            const string revenueRow = "Metric\tMay\tJune\nRevenue EUR\t85519\t82992";
            var revenue = new Dictionary<string, object> { { "value", 82992 }, { "label", "Revenue EUR" }, { "unit", "EUR" }, { "period", "June" }, { "evidence", revenueRow } };
            var cost = new Dictionary<string, object> { { "value", 36714 }, { "label", "Cost EUR" }, { "unit", "EUR" }, { "period", "June" }, { "evidence", auditTable } };
            var margin = new Dictionary<string, object> {
                { "label", "June gross margin" }, { "operation", "margin_percent" }, { "result", 55.76 }, { "unit", "%" }, { "decimals", 2 },
                { "operands", new object[] { revenue, cost } }
            };
            var slide = new Dictionary<string, object> { { "title", "June margin" }, { "subtitle", "June gross margin was 55.76% on revenue of 82992 EUR" },
                { "evidence", auditTable }, { "sources", "WB01 Scribble Draft" }, { "calculations", new[] { margin } } };
            SamsungPresentationReview.ValidateEvidence(json.Serialize(slide), auditTable);

            var definition = PresentationToolCatalog.DraftDefinition();
            Check(json.Serialize(definition.function.parameters).Contains("margin_percent"), "The slide schema does not expose margin_percent.");
            Check(SamsungAuthoringPolicy.Instructions.Contains("margin_percent") && SamsungAuthoringPolicy.Instructions.Contains("not a clarification question"),
                "Authoring policy does not route derived percentages to calculations.");

            margin["operands"] = new object[] { cost, revenue };
            try { SamsungPresentationReview.ValidateEvidence(json.Serialize(slide), auditTable); throw new Exception("Expected reversed operands to fail host arithmetic."); }
            catch (InvalidOperationException ex) { Check(ex.Message.Contains("SLIDE_CALCULATION_MISMATCH"), "Reversed margin operands were accepted: " + ex.Message); }
            margin["operands"] = new object[] { revenue, cost };
            margin["unit"] = "EUR";
            Reject(() => SamsungPresentationReview.ValidateEvidence(json.Serialize(slide), auditTable));
            margin["unit"] = "%";
            margin["result"] = 55.8;
            Reject(() => SamsungPresentationReview.ValidateEvidence(json.Serialize(slide), auditTable));
            margin["result"] = 55.76;

            slide.Remove("calculations");
            try { SamsungPresentationReview.ValidateEvidence(json.Serialize(slide), auditTable); throw new Exception("Expected an unsupported derived value to be rejected."); }
            catch (InvalidOperationException ex)
            {
                Check(ex.Message.Contains("SLIDE_NUMBERS_UNVERIFIED") && ex.Message.Contains("55.76") &&
                    ex.Message.Contains("margin_percent") && ex.Message.Contains("never a question for the user"),
                    "Unsupported derived value did not point to the calculations contract: " + ex.Message);
            }
        }

        // XA01 regression: after the cover slide was written, the model paused the
        // deck to ask how to handle gross margin. The settled request cannot be
        // reopened mid-deck; the host closes ask_user and names the remaining IDs.
        internal static void ClarificationClosesAfterDeckWriteStarts()
        {
            var json = new JavaScriptSerializer();
            var root = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "scribble-ask-gate-" + Guid.NewGuid().ToString("N"));
            try
            {
                var objective = "Use the draft worksheet to produce four new native editable Samsung MD PowerPoint slides: headline, period comparison, June group analysis and data-quality limits.";
                var request = DocumentChatRequestFactory.Create("model", "excel", "Workbook", new ChatTurn[0], objective, true);
                var task = new TaskContextManager(request, "excel", objective, new TaskCheckpointStore(root));
                var ask = new ChatToolCall { id = "ask-1", type = "function", function = new ChatToolCallFunction { name = PromptHelperTool.Name,
                    arguments = json.Serialize(new { question = "How should I handle the gross margin percentages in the deck?", reason = "The host verifies only source numbers.",
                        options = new[] { new { label = "Omit them", description = "Leave margins out." }, new { label = "Show them", description = "Display the margins." } } }) } };
                Check(task.ValidateArguments(ask) == null, "A clarification before any slide was written was rejected.");

                var plan = new[] { "cover", "comparison", "groups", "limits" };
                task.State.HostData["samsung_plan"] = json.Serialize(plan);
                foreach (var id in plan) task.State.ExpectedSourceIds.Add("ppt:" + id);
                task.State.Batches.Add(new TaskBatchResult { Id = "ppt:cover", CoveredSourceIds = new List<string> { "ppt:cover" }, Output = "Source and rendered review passed" });
                var rejected = task.ValidateArguments(ask);
                Check(rejected != null, "ask_user interrupted a deck whose first slide was already written.");
                var payload = json.Deserialize<Dictionary<string, object>>(rejected.Content);
                var message = Convert.ToString(payload["message"]);
                Check(Convert.ToString(payload["error_code"]) == "CLARIFICATION_AFTER_DELIVERABLE_STARTED" && !Convert.ToBoolean(payload["permission_consumed"]),
                    "Mid-deck clarification was not closed with the expected code: " + rejected.Content);
                Check(message.Contains("comparison, groups, limits") && message.Contains("margin_percent") && message.Contains("final response"),
                    "Closed clarification did not name the remaining slide IDs and the calculation route: " + message);

                foreach (var id in plan.Skip(1))
                    task.State.Batches.Add(new TaskBatchResult { Id = "ppt:" + id, CoveredSourceIds = new List<string> { "ppt:" + id }, Output = "Source and rendered review passed" });
                Check(task.ValidateArguments(ask) == null, "The clarification gate outlived the finished deck.");
            }
            finally { if (System.IO.Directory.Exists(root)) System.IO.Directory.Delete(root, true); }
        }

        // The complete citation stays in the speaker notes; a cover shows one short
        // pointer instead of a dense footer of file names.
        internal static void StructuralFootersStayShort()
        {
            var json = new JavaScriptSerializer();
            var sources = string.Join("; ", Enumerable.Range(1, 3).Select(i => "WB01 Ledger sheet rows 2-145 read receipt " + i + " (2026-06 audit)"));
            Check(sources.Length > 110 && sources.Length <= 240, "Fixture footer length does not exercise the structural limit.");
            var cover = new { id = "cover", title = "Atlas Components June review", subtitle = "Operations leadership", layout = "cover", sources };
            var analytical = new { id = "limits", title = "Data-quality limits", subtitle = "One June cost is missing", layout = "bullets", bullets = new[] { "One June cost is missing" }, sources };
            var rendered = json.Serialize(SamsungPresentationReview.InspectPlan(json.Serialize(new object[] { cover, analytical })));
            var texts = System.Text.RegularExpressions.Regex.Matches(rendered, "\"text\":\"((?:[^\"\\\\]|\\\\.)*)\"").Cast<System.Text.RegularExpressions.Match>()
                .Select(m => m.Groups[1].Value).ToArray();
            Check(texts.Any(t => t.EndsWith("full source references in speaker notes.") && t.Length <= 150), "Cover footer was not shortened to a pointer.");
            Check(texts.Contains(sources), "Analytical footer under 240 characters lost its complete visible citation.");
            Check(texts.Count(t => t == sources) == 1, "The cover still rendered the dense footer.");
        }

        internal static void PolicyAndCompletion()
        {
            var definition = PresentationToolCatalog.DraftDefinition();
            Check(definition.function.description.Contains(SamsungAuthoringPolicy.Instructions), "Generation policy differs from shared policy.");
            Check(!definition.function.description.Contains("takeaway sentences as titles"), "Conflicting title rules remain.");
            Check(SamsungAuthoringPolicy.Instructions.Contains("exact verbatim passage") && SamsungAuthoringPolicy.FactReview.Contains("compatible rounding"), "Generation and review prompts do not protect exact source values from paraphrase or rounding false positives.");
            Check(SamsungAuthoringPolicy.Instructions.Contains("Revenue EUR 85519 82992") && SamsungAuthoringPolicy.Instructions.Contains("include its headers"), "Table claim instructions need an explicit ambiguous-evidence counterexample.");
            Check(SamsungAuthoringPolicy.Instructions.Contains("never invent a reformatted row") && SamsungAuthoringPolicy.Instructions.Contains("one claim per period"), "Generation policy does not prevent synthetic comparison citations.");
            var key = SamsungAuthoringPolicy.CacheKey("model", "endpoint", "slide", "evidence");
            Check(key != SamsungAuthoringPolicy.CacheKey("model", "endpoint", "slide", "changed evidence"), "Evidence did not invalidate review.");
            Check(!SamsungAuthoringPolicy.Approved("{\"approved\":true,\"findings\":[{\"severity\":\"blocker\"}]}"), "Blocker approved.");
            Check(!SamsungAuthoringPolicy.Approved("not json"), "Invalid review approved.");
            Check(!SamsungAuthoringPolicy.WellFormedReview("{\"approved\":true,\"issues\":\"Instruction says \"four slides\"\",\"findings\":[]}"), "Malformed quoted review was treated as valid JSON.");
            Check(SamsungAuthoringPolicy.WellFormedReview("{\"approved\":true,\"issues\":\"\",\"findings\":[]}"), "A strict empty approved review was rejected.");
            Check(SamsungAuthoringPolicy.OutlineReview.Contains("internal stable identifiers") && SamsungAuthoringPolicy.ReviewContract.Contains("keep issues under 240"), "Outline review contract must prevent verbose invalid approval loops.");
            var state = new DurableTaskState { EnumerationComplete = true, PresentationReviewRequired = true };
            Check(!state.CanComplete(false), "Deck completed without a final review receipt.");
            state.PresentationReviewReceipt = "reviewed"; Check(state.CanComplete(false), "Valid final receipt rejected.");
            var briefs = new object[] { new Dictionary<string, object> { { "id", "a" }, { "purpose", "explanation" }, { "message", "Definitions" }, { "layout", "bullets" }, { "required_content", new[] { "definition" } } } };
            SamsungAuthoringPolicy.ValidateBriefs(briefs, new[] { "a" });
            Reject(() => SamsungAuthoringPolicy.ValidateBriefs(briefs, new[] { "b" }));
            var factualBrief = new object[] { new Dictionary<string, object> { { "id", "a" }, { "purpose", "analysis" }, { "message", "Finding" }, { "layout", "bullets" }, { "required_content", new[] { "finding" } }, { "source_spans", new string[0] } } };
            var factualSlide = new[] { new Dictionary<string, object> { { "id", "a" }, { "title", "Finding" }, { "layout", "bullets" }, { "content_kind", "fact" } } };
            Reject(() => SamsungAuthoringPolicy.ValidateSourceSpanCoverage(factualBrief, factualSlide, true));
            factualBrief[0] = new Dictionary<string, object> { { "id", "a" }, { "purpose", "analysis" }, { "message", "Finding" }, { "layout", "bullets" }, { "required_content", new[] { "finding" } }, { "source_spans", new[] { "span:0" } } };
            factualSlide[0]["source_spans"] = new[] { "span:0" };
            SamsungAuthoringPolicy.ValidateSourceSpanCoverage(factualBrief, factualSlide, true);

            var scopeRoot = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                "scribble-slide-scope-" + System.Guid.NewGuid().ToString("N"));
            try
            {
                var objective = "Use this workbook to produce four new native editable Samsung MD PowerPoint slides.";
                var scopedRequest = DocumentChatRequestFactory.Create("model", "excel", "Workbook", new ChatTurn[0],
                    objective, true);
                var scopedTask = new TaskContextManager(scopedRequest, "excel", objective,
                    new TaskCheckpointStore(scopeRoot));
                Check(scopedTask.State.RequiredPresentationSlides == 4, "Adjectives between the requested count and slides lost the exact slide count.");
                Check(scopedRequest.tools.Any(tool => tool.function.name == CrossAppToolCatalog.SendToPowerPoint) &&
                    !scopedRequest.tools.Any(tool => tool.function.name == WorkbookToolCatalog.WriteDraftSheet ||
                        tool.function.name == CrossAppToolCatalog.SendToWord ||
                        tool.function.name == CrossAppToolCatalog.CreateEmailDraft),
                    "An exact slide deliverable exposed an unrelated document write surface.");
            }
            finally { if (System.IO.Directory.Exists(scopeRoot)) System.IO.Directory.Delete(scopeRoot, true); }
        }

        internal static void ReadReceiptsExposeSourceSpans()
        {
            var root = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "scribble-source-spans-" + Guid.NewGuid().ToString("N"));
            try
            {
                var request = new ChatCompletionRequest { model = "local-model", messages = new List<object> { new ChatCompletionInputMessage { role = "user", content = "Read evidence" } } };
                var task = new TaskContextManager(request, "outlook", "Read evidence", new TaskCheckpointStore(root));
                var call = new ChatToolCall { id = "read-1", function = new ChatToolCallFunction { name = "read_messages", arguments = "{}" } };
                var result = new MailboxToolResult(call.id, "{\"content\":\"Revenue AED 420000\"}", "Read message");
                task.AfterTool(call, result);
                var parsed = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(result.Content);
                var ids = ((IEnumerable)parsed["source_spans"]).Cast<object>().Select(Convert.ToString).ToArray();
                Check(ids.Length == 1 && ids[0].Contains(":"), "Read receipt did not expose its host-issued source span.");
                var before = task.Sources.Spans().Count;
                task.Sources.CaptureRead(call, result);
                Check(task.Sources.Spans().Count == before, "Replaying an enriched receipt created citation-only source spans.");
            }
            finally { if (System.IO.Directory.Exists(root)) System.IO.Directory.Delete(root, true); }
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

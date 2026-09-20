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
        // XA01: gross margin is absent from the workbook audit as a percentage.
        // It must be accepted only as a host-recomputed calculation from the
        // exact Revenue and Cost operands, never from model prose.
        internal static void DerivedMarginCalculation()
        {
            var json = new JavaScriptSerializer();
            const string audit = "Metric\tMay (2026-05)\tJune (2026-06)\nRevenue EUR\t85519\t82992\nCost EUR\t36702\t36714\nGross margin\t0.570832212724658\t0.55762001156738";
            const string revenueBlock = "Metric\tMay (2026-05)\tJune (2026-06)\nRevenue EUR\t85519\t82992";
            const string costBlock = "Metric\tMay (2026-05)\tJune (2026-06)\nRevenue EUR\t85519\t82992\nCost EUR\t36702\t36714";
            Func<string, decimal, decimal, decimal, Dictionary<string, object>> margin = (period, revenue, cost, result) => new Dictionary<string, object> {
                { "label", "Gross margin" }, { "operation", "margin_percent" }, { "result", result }, { "unit", "%" }, { "decimals", 2 },
                { "operands", new object[] {
                    new Dictionary<string, object> { { "value", revenue }, { "label", "Revenue EUR" }, { "unit", "EUR" }, { "period", period }, { "evidence", revenueBlock } },
                    new Dictionary<string, object> { { "value", cost }, { "label", "Cost EUR" }, { "unit", "EUR" }, { "period", period }, { "evidence", costBlock } }
                } }
            };
            var june = margin("June (2026-06)", 82992m, 36714m, 55.76m);
            var may = margin("May (2026-05)", 85519m, 36702m, 57.08m);
            var slide = new Dictionary<string, object> { { "title", "Gross margin" }, { "subtitle", "Gross margin moved from 57.08% to 55.76%" },
                { "evidence", audit }, { "sources", "WB01 audit" }, { "calculations", new[] { may, june } } };
            SamsungPresentationReview.ValidateEvidence(json.Serialize(slide), audit);

            // The source fraction 0.5576... never verifies a displayed percentage.
            var unsupported = new Dictionary<string, object>(slide); unsupported.Remove("calculations");
            try { SamsungPresentationReview.ValidateEvidence(json.Serialize(unsupported), audit); throw new Exception("A derived percentage was accepted from prose."); }
            catch (InvalidOperationException ex) { Check(ex.Message.StartsWith("SLIDE_NUMBERS_UNVERIFIED") && ex.Message.Contains("55.76") && ex.Message.Contains("57.08"), "Unsupported derived percentages were not both reported."); }

            june["result"] = 55.8m; Reject(() => SamsungPresentationReview.ValidateEvidence(json.Serialize(slide), audit)); june["result"] = 55.76m;
            june["unit"] = "EUR"; Reject(() => SamsungPresentationReview.ValidateEvidence(json.Serialize(slide), audit)); june["unit"] = "%";
            var operands = (object[])june["operands"];
            june["operands"] = new[] { operands[1], operands[0] };
            Reject(() => SamsungPresentationReview.ValidateEvidence(json.Serialize(slide), audit));
            june["operands"] = operands;
            ((Dictionary<string, object>)operands[1])["value"] = 36000m;
            Reject(() => SamsungPresentationReview.ValidateEvidence(json.Serialize(slide), audit));
            ((Dictionary<string, object>)operands[1])["value"] = 36714m;
            SamsungPresentationReview.ValidateEvidence(json.Serialize(slide), audit);

            // A host-computed decrease may be displayed as its size ("fell 2.95%")
            // or with a typographic minus; both are the recomputed -2.95.
            var growth = new Dictionary<string, object> {
                { "label", "Revenue change" }, { "operation", "growth_percent" }, { "result", -2.95m }, { "unit", "%" }, { "decimals", 2 },
                { "operands", new object[] {
                    new Dictionary<string, object> { { "value", 82992m }, { "label", "Revenue EUR" }, { "unit", "EUR" }, { "period", "June (2026-06)" }, { "evidence", revenueBlock } },
                    new Dictionary<string, object> { { "value", 85519m }, { "label", "Revenue EUR" }, { "unit", "EUR" }, { "period", "May (2026-05)" }, { "evidence", revenueBlock } }
                } }
            };
            slide["calculations"] = new[] { may, june, growth };
            slide["subtitle"] = "Revenue fell 2.95% (−2.95%) while gross margin moved from 57.08% to 55.76%";
            SamsungPresentationReview.ValidateEvidence(json.Serialize(slide), audit);
            slide["subtitle"] = "Revenue fell 2.96% while gross margin moved from 57.08% to 55.76%";
            Reject(() => SamsungPresentationReview.ValidateEvidence(json.Serialize(slide), audit));

            // A footnote that is plainly the source line is the citation.
            var footnoted = new Dictionary<string, object> { { "footnote", " Source: WB01 Ledger (live SUMIFS)" } };
            SamsungPresentationReview.AdoptFootnoteCitation(footnoted);
            Check((string)footnoted["sources"] == "Source: WB01 Ledger (live SUMIFS)" && !footnoted.ContainsKey("footnote"), "A source-line footnote was not adopted as the citation.");
            var qualified = new Dictionary<string, object> { { "footnote", "Axis starts at 0." } };
            SamsungPresentationReview.AdoptFootnoteCitation(qualified);
            Check(!qualified.ContainsKey("sources") && qualified.ContainsKey("footnote"), "A qualifying footnote was mistaken for a citation.");
            var cited = new Dictionary<string, object> { { "sources", "WB01" }, { "footnote", "Source: other" } };
            SamsungPresentationReview.AdoptFootnoteCitation(cited);
            Check((string)cited["sources"] == "WB01" && cited.ContainsKey("footnote"), "An existing citation was replaced.");
            slide["subtitle"] = "Revenue fell 2.95% while gross margin moved from 57.08% to 55.76%";
            var uncited = new Dictionary<string, object>(slide); uncited.Remove("sources");
            try { SamsungPresentationReview.ValidateEvidence(json.Serialize(uncited), audit); throw new Exception("A factual slide without a citation was accepted."); }
            catch (InvalidOperationException ex) { Check(ex.Message.StartsWith("SLIDE_CITATION_REQUIRED") && ex.Message.Contains("sources string"), "The missing-citation rejection does not name the sources field."); }

            Check(SamsungAuthoringPolicy.Instructions.Contains("margin_percent") && SamsungAuthoringPolicy.Instructions.Contains("never through ask_user"), "Authoring policy does not route derived values to calculations.");
            Check(SamsungEvidence.DerivedValueGuidance.Contains("margin_percent") && SamsungEvidence.DerivedValueGuidance.Contains("Do not call ask_user"), "Derived-value recovery guidance is incomplete.");
            var draft = json.Serialize(PresentationToolCatalog.DraftDefinition());
            Check(draft.Contains("margin_percent"), "The draft tool schema does not expose the margin calculation.");
        }

        // XA01 run 2026-09-17: the audit table heads its columns May and June
        // while the required chart categories are 2026-05 and 2026-06, and the
        // model cites bare data rows for calculation operands.
        internal static void PeriodLabelsAndOperandHeaders()
        {
            var json = new JavaScriptSerializer();
            const string audit = "Metric\tMay\tJune\nRevenue EUR\t85519\t82992\nCost EUR\t36702\t36714";
            const string corpus = audit + "\nRowID\tPeriod\tRevenueEUR\n1\t2026-05\t2538\n2\t2026-06\t3990";
            Func<string, decimal, decimal, Dictionary<string, object>> margin = (period, revenue, cost) => new Dictionary<string, object> {
                { "label", "Gross margin " + period }, { "operation", "margin_percent" }, { "result", period == "2026-06" ? 55.76m : 57.08m }, { "unit", "%" }, { "decimals", 2 },
                { "operands", new object[] {
                    new Dictionary<string, object> { { "value", revenue }, { "label", "Revenue EUR" }, { "unit", "EUR" }, { "period", period }, { "evidence", "Revenue EUR\t85519\t82992" } },
                    new Dictionary<string, object> { { "value", cost }, { "label", "Cost EUR" }, { "unit", "EUR" }, { "period", period }, { "evidence", "Cost EUR\t36702\t36714" } }
                } }
            };
            var may = margin("2026-05", 85519m, 36702m);
            var slide = new Dictionary<string, object> {
                { "title", "Revenue by month (EUR)" }, { "subtitle", "Gross margin moved from 57.08% in 2026-05 to 55.76% in 2026-06" }, { "layout", "chart" },
                { "chart", new Dictionary<string, object> { { "title", "Revenue (EUR)" }, { "categories", new[] { "2026-05", "2026-06" } },
                    { "series", new object[] { new Dictionary<string, object> { { "name", "Revenue EUR" }, { "values", new[] { 85519, 82992 } } } } } } },
                { "claims", new object[] { new Dictionary<string, object> { { "text", "June revenue EUR 82,992" }, { "label", "Revenue EUR" }, { "unit", "EUR" },
                    { "period", "2026-06" }, { "evidence", "Revenue EUR\t85519\t82992" } } } },
                { "calculations", new object[] { may, margin("2026-06", 82992m, 36714m) } },
                { "evidence", audit }, { "sources", "WB01 audit" } };
            SamsungPresentationReview.ValidateEvidence(json.Serialize(slide), corpus);

            // A header line plus a later whole row of the same block resolves to
            // the contiguous verified block; a partial or invented row does not.
            var claims = (object[])slide["claims"];
            var stitched = (Dictionary<string, object>)claims[0];
            stitched["text"] = "June cost EUR 36,714"; stitched["label"] = "Cost EUR"; stitched["evidence"] = "Metric\tMay\tJune\nCost EUR\t36702\t36714";
            SamsungPresentationReview.ValidateEvidence(json.Serialize(slide), corpus);
            stitched["evidence"] = "Metric\tMay\tJune\nCost EUR\t36702";
            Reject(() => SamsungPresentationReview.ValidateEvidence(json.Serialize(slide), corpus));
            stitched["evidence"] = "Metric\tMay\tJune\n2026-06\t82992\t36714";
            Reject(() => SamsungPresentationReview.ValidateEvidence(json.Serialize(slide), corpus));
            stitched["text"] = "June revenue EUR 82,992"; stitched["label"] = "Revenue EUR"; stitched["evidence"] = "Revenue EUR\t85519\t82992";

            // XA01 run 2026-09-20: read_grouped_totals emits compact source
            // headers (RevenueEUR), but slide labels remain human-readable
            // (Revenue EUR). The host may expand a bare exact row to that
            // adjacent header; it must not accept a different metric.
            const string grouped = "Period\tGroup\tRows\tRevenueEUR\tCostEUR\tBlank or non-numeric cells\n2026-06\tNorth\t6\t19219\t8082\t0";
            var groupedClaim = new Dictionary<string, object> {
                { "text", "North June revenue 19,219 EUR" }, { "label", "Revenue EUR" }, { "unit", "EUR" },
                { "period", "2026-06" }, { "evidence", "2026-06\tNorth\t6\t19219\t8082\t0" }
            };
            var groupedSlide = new Dictionary<string, object> { { "claims", new object[] { groupedClaim } } };
            SamsungEvidence.ValidateClaims(groupedSlide, grouped);
            Check(((string)groupedClaim["evidence"]).Contains("RevenueEUR"),
                "A compact host header was not attached to its exact grouped-total row.");
            groupedClaim["label"] = "Profit EUR";
            Reject(() => SamsungEvidence.ValidateClaims(groupedSlide, grouped));

            // A label the task never read is still refused, as is a quantity.
            slide["subtitle"] = "Gross margin was 55.76% in 2026-06 against a 2026-07 plan";
            try { SamsungPresentationReview.ValidateEvidence(json.Serialize(slide), corpus); throw new Exception("An unread period label was accepted."); }
            catch (InvalidOperationException ex) { Check(ex.Message.StartsWith("SLIDE_NUMBERS_UNVERIFIED") && ex.Message.Contains("2026"), "An unread period label was not reported as unverified."); }
            slide["subtitle"] = "Gross margin was 55.76% in 2026-06 on 4,100 units";
            Reject(() => SamsungPresentationReview.ValidateEvidence(json.Serialize(slide), corpus));
            slide["subtitle"] = "Gross margin moved from 57.08% in 2026-05 to 55.76% in 2026-06";

            // The month header cannot stand in for a different year.
            const string lastYear = "Metric\tMay 2025\tJune 2025\nRevenue EUR\t85519\t82992\nCost EUR\t36702\t36714";
            var dated = new Dictionary<string, object>(slide); dated["evidence"] = lastYear;
            Reject(() => SamsungPresentationReview.ValidateEvidence(json.Serialize(dated), lastYear + "\n" + corpus));

            may["result"] = 57.18m;
            try { SamsungPresentationReview.ValidateEvidence(json.Serialize(slide), corpus); throw new Exception("A wrong margin was accepted."); }
            catch (InvalidOperationException ex) { Check(ex.Message.StartsWith("SLIDE_CALCULATION_MISMATCH") && ex.Message.Contains("57.08") && ex.Message.Contains("57.18"), "The mismatch did not return the host-computed value."); }
            may["result"] = 57.08m;
            ((Dictionary<string, object>)((object[])may["operands"])[0])["period"] = "2026-04";
            try { SamsungPresentationReview.ValidateEvidence(json.Serialize(slide), corpus); throw new Exception("An operand for an absent column was accepted."); }
            catch (InvalidOperationException ex) { Check(ex.Message.StartsWith("SLIDE_OPERAND_ASSOCIATION") && ex.Message.Contains("period '2026-04'") && ex.Message.Contains("Metric May June"), "The operand rejection does not name the missing period and a verified block."); }
            Check(SamsungAuthoringPolicy.FactReview.Contains("month name alone"), "The fact reviewer is not told how period labels were verified.");
        }

        // After the first verified slide, a clarification that the tool contract
        // already answers is resolved by the host; a persistent one reaches the user.
        internal static void ClarificationDeferral()
        {
            var root = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "scribble-clarify-" + Guid.NewGuid().ToString("N"));
            try
            {
                var objective = "Use this workbook to produce four new native editable Samsung MD PowerPoint slides.";
                var request = DocumentChatRequestFactory.Create("model", "excel", "Workbook", new ChatTurn[0], objective, true);
                var task = new TaskContextManager(request, "excel", objective, new TaskCheckpointStore(root));
                var ask = new ChatToolCall { id = "ask-1", function = new ChatToolCallFunction { name = PromptHelperTool.Name,
                    arguments = "{\"question\":\"How should I handle the gross margin percentages?\",\"options\":[{\"label\":\"Omit\",\"description\":\"Leave out\"},{\"label\":\"Keep\",\"description\":\"Show\"}]}" } };
                Check(task.DeferClarification(ask) == null, "A clarification before any slide was written must reach the user.");
                foreach (var id in new[] { "cover", "compare", "groups", "limits" }) task.State.ExpectedSourceIds.Add("ppt:" + id);
                task.State.Batches.Add(new TaskBatchResult { Id = "ppt:cover", CoveredSourceIds = new List<string> { "ppt:cover" } });
                var read = new ChatToolCall { id = "read-1", function = new ChatToolCallFunction { name = "read_range", arguments = "{}" } };
                Check(task.DeferClarification(read) == null, "Only ask_user can be deferred.");
                for (var attempt = 0; attempt < TaskContextManager.MaxDeferredClarifications; attempt++)
                {
                    var deferred = task.DeferClarification(ask);
                    Check(deferred != null && deferred.Outcome.Failed && deferred.Outcome.ErrorCode == "TASK_CLARIFICATION_DEFERRED" &&
                        deferred.Outcome.PermissionConsumed == false, "A mid-deliverable clarification was not deferred as a failed, permission-free result.");
                    Check(deferred.Content.Contains("compare") && deferred.Content.Contains("groups") && deferred.Content.Contains("limits") &&
                        !deferred.Content.Contains("\"cover\"") && deferred.Content.Contains("margin_percent"),
                        "The deferral did not name the remaining planned slides and the calculation contract.");
                }
                Check(task.DeferClarification(ask) == null, "A persistent clarification must eventually reach the user.");
                var decisions = task.State.OriginalDecisions.Count;
                Check(decisions == 1, "A host deferral was recorded as a user decision.");

                var done = new TaskContextManager(DocumentChatRequestFactory.Create("model", "excel", "Workbook", new ChatTurn[0], objective, true),
                    "excel", objective, new TaskCheckpointStore(root));
                done.State.ExpectedSourceIds.Add("ppt:cover");
                done.State.Batches.Add(new TaskBatchResult { Id = "ppt:cover", CoveredSourceIds = new List<string> { "ppt:cover" } });
                Check(done.DeferClarification(ask) == null, "A completed deck has nothing left to continue.");
            }
            finally { if (System.IO.Directory.Exists(root)) System.IO.Directory.Delete(root, true); }
        }

        // After a slide write stops part-way, a revised payload is refused as a
        // bounded tool error that hands back the original arguments, instead
        // of ending the task with a fatal uncertain-write exception.
        internal static void InterruptedWriteRedirectsToOriginalPayload()
        {
            var root = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "scribble-write-recovery-" + Guid.NewGuid().ToString("N"));
            try
            {
                var objective = "Use this workbook to produce four new native editable Samsung MD PowerPoint slides.";
                var request = DocumentChatRequestFactory.Create("model", "excel", "Workbook", new ChatTurn[0], objective, true);
                var task = new TaskContextManager(request, "excel", objective, new TaskCheckpointStore(root));
                var revised = new ChatToolCall { id = "call-2", function = new ChatToolCallFunction { name = CrossAppToolCatalog.SendToPowerPoint, arguments = "{\"slides\":[{\"id\":\"b\"}]}" } };
                Check(task.RecoverableWriteConflict(revised, true) == null, "A task without an interrupted write was redirected.");
                task.State.HostData["samsung_pending"] = new JavaScriptSerializer().Serialize(new Dictionary<string, object> {
                    { "Owner", "o" }, { "Input", "different-input-hash" }, { "ToolCall", "call-1" }, { "Arguments", "{\"slides\":[{\"id\":\"a\"}]}" } });
                task.State.Writes.Add(new TaskWriteRecord { Id = "tool:call-1", Status = "uncertain" });
                Check(task.RecoverableWriteConflict(revised, false) == null, "A read-only call was treated as a write conflict.");
                for (var attempt = 0; attempt < TaskContextManager.MaxWriteRecoveryRedirects; attempt++)
                {
                    var redirected = task.RecoverableWriteConflict(revised, true);
                    Check(redirected != null && redirected.Outcome.Failed && redirected.Outcome.ErrorCode == "SLIDE_RECOVERY_INPUT_CHANGED" &&
                        redirected.Outcome.PermissionConsumed == false && redirected.Content.Contains("original_arguments") &&
                        redirected.Content.Contains("\\\"id\\\":\\\"a\\\""), "A revised payload after an interrupted write was not redirected to the original arguments.");
                }
                Check(task.RecoverableWriteConflict(revised, true) == null, "Write-recovery redirects must stay bounded.");
            }
            finally { if (System.IO.Directory.Exists(root)) System.IO.Directory.Delete(root, true); }
        }

        // XA01 workbook: the model planned one group-table layout and emitted another.
        internal static void DraftFormulaAssociations()
        {
            var requiredPrompt = "Use live Excel formulas linked to the source observations, not pasted constants. Use linked formulas for B4:C5.";
            var pasted = new[] { new[] { "Metric", "May", "June" }, new[] { "Revenue EUR", "0", "0" }, new[] { "Cost EUR", "0", "0" } };
            var requiredIssues = DraftFormulaAssociation.ValidatePromptRequirements(requiredPrompt, pasted);
            Check(requiredIssues.Count == 4 && requiredIssues.All(issue => issue.Contains("live formula")),
                "Pasted answer constants were accepted where the prompt named required live-formula cells.");
            var generalPrompt = "Use live Excel formulas linked to the source observations for calculated outputs, not pasted answer constants.";
            requiredIssues = DraftFormulaAssociation.ValidatePromptRequirements(generalPrompt, pasted);
            Check(requiredIssues.Count == 4 && requiredIssues.All(issue => issue.Contains("pasted numeric constant")),
                "Pasted answer constants were accepted under the general live-formula requirement.");
            var linked = new[] { new[] { "Metric", "May", "June" },
                new[] { "Revenue EUR", "=SUMIF(Ledger!A:A,\"2026-05\",Ledger!I:I)", "=SUMIF(Ledger!A:A,\"2026-06\",Ledger!I:I)" },
                new[] { "Cost EUR", "=SUMIF(Ledger!A:A,\"2026-05\",Ledger!J:J)", "=SUMIF(Ledger!A:A,\"2026-06\",Ledger!J:J)" } };
            Check(DraftFormulaAssociation.ValidatePromptRequirements(requiredPrompt, linked).Count == 0,
                "Correct source-linked formulas were rejected.");
            Check(DraftFormulaAssociation.ValidatePromptRequirements(generalPrompt, linked).Count == 0,
                "Correct source-linked formulas were rejected under the general requirement.");
            var draftSchema = new JavaScriptSerializer().Serialize(WorkbookToolCatalog.DraftDefinition());
            Check(draftSchema.Contains("never put 0") && draftSchema.Contains("pasted answer constant"),
                "The Excel draft schema does not warn the model against formula placeholders.");
            const string may = "=SUMIF(Ledger!$B$2:$B$145,\"2026-05\",Ledger!$I$2:$I$145)";
            const string june = "=SUMIF(Ledger!$B$2:$B$145,\"2026-06\",Ledger!$I$2:$I$145)";
            Func<string, string, string> group = (column, name) => "=SUMIFS(Ledger!$" + column + "$2:$" + column + "$145,Ledger!$B$2:$B$145,\"2026-06\",Ledger!$D$2:$D$145,\"" + name + "\")";
            Func<string[], string[][]> sheet = tail => new[] {
                new[] { "Metric", "May (2026-05)", "June (2026-06)" },
                new[] { "Revenue EUR", may, june },
                new[] { "Cost EUR", may.Replace("$I$", "$J$"), june.Replace("$I$", "$J$") },
                new[] { "Gross profit EUR", "=B4-B5", "=C4-C5" },
                new[] { "Gross margin", "=(B4-B5)/B4", "=(C4-C5)/C4" },
                new[] { "Observations", "=SUMPRODUCT((Ledger!$B$2:$B$145=\"2026-05\")*(Ledger!$I$2:$I$145<>\"\"))", "=SUMPRODUCT((Ledger!$B$2:$B$145=\"2026-06\")*(Ledger!$I$2:$I$145<>\"\"))" },
                new string[0],
                new[] { "June by group (source 'Ledger', June 2026 rows 122-145)" },
                new[] { "Group", "Revenue EUR", "Cost EUR", "Gross margin" },
                new[] { "North", group("I", "North"), group("J", "North"), tail[0] },
                new[] { "South", group("I", "South"), group("J", "South"), tail[1] },
                new[] { "East", group("I", "East"), group("J", "East"), tail[2] },
                new[] { "West", group("I", "West"), group("J", "West"), tail[3] },
                new[] { "Total (check)", tail[4], tail[5], tail[6] }
            };
            var captured = DraftFormulaAssociation.Validate(sheet(new[] {
                "=(B11-B12)/B11", "=(B13-B14)/B13", "=(B15-B16)/B15", "=(B17-B18)/B17", "=SUM(B11:B14)", "=SUM(B15:B18)", "=(B19-B20)/B19" }));
            foreach (var cell in new[] { "D12", "D13", "D14", "D15", "B16", "C16", "D16" })
                Check(captured.Any(issue => issue.StartsWith(cell + " ", StringComparison.Ordinal)), "The shifted formula in " + cell + " was not reported.");
            Check(captured.Count == 7, "A correct audit formula was reported as misassociated: " + string.Join(" | ", captured));
            Check(captured.Any(issue => issue.Contains("B12:B15")), "The total repair did not name the exact data block.");
            var message = DraftFormulaAssociation.RepairMessage(captured);
            Check(message.StartsWith("DRAFT_FORMULA_ASSOCIATION") && message.Contains("rows[i] is sheet row i+3") && message.Contains("no draft permission was consumed"),
                "The formula repair message does not restate the deterministic layout.");

            var repaired = DraftFormulaAssociation.Validate(sheet(new[] {
                "=(B12-C12)/B12", "=(B13-C13)/B13", "=(B14-C14)/B14", "=(B15-C15)/B15", "=SUM(B12:B15)", "=SUM(C12:C15)", "=(B16-C16)/B16" }));
            Check(repaired.Count == 0, "Correct same-row rates and totals were rejected: " + string.Join(" | ", repaired));
            // A share of a common total is a filled column anchored on one cell.
            var shares = DraftFormulaAssociation.Validate(sheet(new[] {
                "=B12/B16", "=B13/B16", "=B14/B16", "=B15/$B$16", "=SUM(B12:B15)", "=SUM(C12:C15)", "=B16/B16" }));
            Check(shares.Count == 0, "A share-of-total column was rejected: " + string.Join(" | ", shares));
            // XA01 run 2026-09-17: a derived-metric list under the audit table
            // reads across the table's rows; it has no record row to ignore.
            var derived = DraftFormulaAssociation.Validate(new[] {
                new[] { "Metric", "May", "June" }, new[] { "Revenue EUR", "85519", "82992" }, new[] { "Cost EUR", "36702", "36714" }, new string[0],
                new[] { "Revenue change EUR", "=C4-B4" }, new[] { "Revenue change", "=(C4-B4)/B4" },
                new[] { "Cost change EUR", "=C5-B5" }, new[] { "Cost change", "=(C5-B5)/B5" },
                new[] { "Columns summed", "=SUM(C4:C5)" }, new[] { "Rows summed", "=SUM(B4:C4)" } });
            Check(derived.Count == 0, "A derived-metric list under a table was rejected: " + string.Join(" | ", derived));
            // A numeric year header directly above the data is not a missed row.
            var years = DraftFormulaAssociation.Validate(new[] {
                new[] { "Region", "2025" }, new[] { "North", "10" }, new[] { "South", "12" }, new[] { "Total", "=SUM(B4:B5)" } });
            Check(years.Count == 0, "A total under a numeric year header was rejected: " + string.Join(" | ", years));
            var quoted = DraftFormulaAssociation.Validate(new[] {
                new[] { "Quarter", "Units" }, new[] { "Q1", "10" }, new[] { "Q2", "=IF(A4=\"Q1\",B4,0)" }, new[] { "Label", "=LOG10(B4)" } });
            Check(quoted.Count == 0, "Function formulas or quoted cell-like text were misread as arithmetic references.");
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

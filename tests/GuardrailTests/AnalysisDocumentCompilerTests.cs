using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Script.Serialization;
using Scribble.Chat;
using Scribble.Office;

namespace GuardrailTests
{
    internal static class AnalysisDocumentCompilerTests
    {
        public static void DeckToolAcceptsOnlyFactReferencedPlan()
        {
            var definition = CrossAppToolCatalog.AnalysisDeckDefinition();
            var json = new JavaScriptSerializer();
            var valid = new ChatToolCall
            {
                id = "typed-deck", type = "function",
                function = new ChatToolCallFunction
                {
                    name = CrossAppToolCatalog.SendToPowerPoint,
                    arguments = json.Serialize(new
                    {
                        AnalysisId = "host-issued-id",
                        Slides = new[] { new
                        {
                            Id = "headline", Layout = "scorecard",
                            Title = "Verified trend",
                            Subtitle = new[] { new { FactId = "fact-id" } }
                        } }
                    })
                }
            };
            Check(ToolContractValidator.Validate(valid, definition).Count == 0,
                "The fact-referenced deck schema rejected its minimal plan.");
            var injected = new ChatToolCall
            {
                id = "typed-deck-injection", type = "function",
                function = new ChatToolCallFunction
                {
                    name = CrossAppToolCatalog.SendToPowerPoint,
                    arguments = valid.function.arguments.Replace(
                        "\"FactId\":\"fact-id\"",
                        "\"FactId\":\"fact-id\",\"Formula\":\"=1\"")
                }
            };
            Check(ToolContractValidator.Validate(injected, definition)
                    .Any(error => error.Contains("Formula")),
                "The model-facing deck schema accepted an authored formula.");
        }

        public static void SharedAnalysisAllowsOneDraftPerDestination()
        {
            var prior = Environment.GetEnvironmentVariable(
                AnalysisDocumentPilot.FeatureFlag);
            var root = Path.Combine(Path.GetTempPath(),
                "scribble-analysis-write-scope-" +
                Guid.NewGuid().ToString("N"));
            try
            {
                Environment.SetEnvironmentVariable(
                    AnalysisDocumentPilot.FeatureFlag, "1");
                var request = new ChatCompletionRequest
                {
                    model = "offline-test",
                    messages = new List<object>()
                };
                var task = new TaskContextManager(request, "excel",
                    "Create a workbook and deck from this analysis",
                    new TaskCheckpointStore(root));
                task.State.AnalysisContractVersion =
                    AnalysisContract.Version;
                task.State.AnalysisArtifactEvidenceId =
                    "host-owned-evidence-id";
                Func<string, string, ChatToolCall> call = (id, name) =>
                    new ChatToolCall
                    {
                        id = id, type = "function",
                        function = new ChatToolCallFunction
                        {
                            name = name, arguments = "{}"
                        }
                    };
                var workbook = call("workbook", 
                    WorkbookToolCatalog.WriteDraftSheet);
                task.BeforeTool(workbook, true);
                task.AfterTool(workbook, new MailboxToolResult(
                    workbook.id, "{\"ok\":true}", "Workbook draft"));
                var deck = call("deck",
                    CrossAppToolCatalog.SendToPowerPoint);
                task.BeforeTool(deck, true);
                task.AfterTool(deck, new MailboxToolResult(
                    deck.id, "{\"ok\":true}", "Deck draft"));
                var duplicateRejected = false;
                try { task.BeforeTool(call("duplicate",
                    WorkbookToolCatalog.WriteDraftSheet), true); }
                catch (InvalidOperationException error)
                { duplicateRejected = error.Message.Contains(
                    "already completed"); }
                Check(duplicateRejected,
                    "A second analysis workbook bypassed its write receipt.");
            }
            finally
            {
                Environment.SetEnvironmentVariable(
                    AnalysisDocumentPilot.FeatureFlag, prior);
                if (Directory.Exists(root)) Directory.Delete(root, true);
            }
        }

        public static void PilotRequiresExplicitFeatureFlag()
        {
            var prior = Environment.GetEnvironmentVariable(
                AnalysisDocumentPilot.FeatureFlag);
            try
            {
                Environment.SetEnvironmentVariable(
                    AnalysisDocumentPilot.FeatureFlag, null);
                var blocked = false;
                try { AnalysisDocumentPilot.WriteWorkbook(null, null, null); }
                catch (InvalidOperationException error)
                { blocked = error.Message.Contains("ANALYSIS_PILOT_DISABLED"); }
                Check(blocked, "The development pilot writer was available by default.");
            }
            finally
            {
                Environment.SetEnvironmentVariable(
                    AnalysisDocumentPilot.FeatureFlag, prior);
            }
        }

        public static void OneAnalysisSuppliesWorkbookAndFourSlides()
        {
            var locator = new SourceLocator
            {
                Kind = "excel_range", SourceInstanceId = "WB01",
                WorksheetIdentity = "Ledger", Range = "A1:L145"
            };
            var snapshot = AnalysisContract.CreateSnapshot("WB01",
                "excel_workbook", "revision-1", "complete_range",
                "recalculated", new[] { locator }, new TableDataset[0]);
            var mayRevenue = Fact(snapshot, locator, "RevenueEUR", "85519",
                "2026-05");
            var juneRevenue = Fact(snapshot, locator, "RevenueEUR", "82992",
                "2026-06");
            var mayCost = Fact(snapshot, locator, "CostEUR", "36702",
                "2026-05");
            var juneCost = Fact(snapshot, locator, "CostEUR", "36714",
                "2026-06");
            var southRevenue = Fact(snapshot, locator, "SouthRevenueEUR",
                "22675", "2026-06");
            var westRevenue = Fact(snapshot, locator, "WestRevenueEUR",
                "22044", "2026-06");
            VerifiedFact juneMargin;
            var marginCalculation = AnalysisCalculator.Calculate(
                AnalysisCalculator.Margin,
                new[] { juneRevenue, juneCost }, "GrossMargin",
                "2026-06", out juneMargin);
            var artifact = AnalysisContract.CreateArtifact(new[] { snapshot },
                new[] { mayRevenue, juneRevenue, mayCost, juneCost,
                    southRevenue, westRevenue, juneMargin },
                new[] { marginCalculation }, new string[0], new string[0]);

            var plan = new AnalysisDocumentPlan
            {
                AnalysisId = artifact.AnalysisId,
                WorkbookTitle = "Atlas Components audit",
                WorkbookRows = new List<AnalysisPlanRow>
                {
                    Row(Text("Metric"), Text("May"), Text("June")),
                    Row(Text("Revenue EUR"), Formula(
                        "=SUMIF(Ledger!$B$2:$B$145,\"2026-05\",Ledger!$I$2:$I$145)",
                        mayRevenue), Formula(
                        "=SUMIF(Ledger!$B$2:$B$145,\"2026-06\",Ledger!$I$2:$I$145)",
                        juneRevenue)),
                    Row(Text("Cost EUR"), Formula(
                        "=SUMIF(Ledger!$B$2:$B$145,\"2026-05\",Ledger!$J$2:$J$145)",
                        mayCost), Formula(
                        "=SUMIF(Ledger!$B$2:$B$145,\"2026-06\",Ledger!$J$2:$J$145)",
                        juneCost))
                },
                Slides = new List<AnalysisPlanSlide>
                {
                    new AnalysisPlanSlide
                    {
                        Id = "headline", Layout = "scorecard",
                        Title = "June sales audit",
                        Subtitle = new List<AnalysisPlanText>
                        {
                            Word("June revenue "), Value(juneRevenue),
                            Word(" versus May "), Value(mayRevenue)
                        },
                        Cards = new List<AnalysisPlanCard>
                        {
                            new AnalysisPlanCard
                            {
                                Heading = "Revenue EUR",
                                Points = new List<AnalysisPlanText> { Value(juneRevenue) }
                            },
                            new AnalysisPlanCard
                            {
                                Heading = "Gross margin",
                                Points = new List<AnalysisPlanText> { Value(juneMargin) }
                            }
                        }
                    },
                    new AnalysisPlanSlide
                    {
                        Id = "periods", Layout = "two_pane",
                        Title = "May and June comparison",
                        TableHeaders = new List<string> { "Metric", "May", "June" },
                        TableRows = new List<AnalysisPlanRow>
                        {
                            Row(Text("Revenue EUR"), Ref(mayRevenue), Ref(juneRevenue)),
                            Row(Text("Cost EUR"), Ref(mayCost), Ref(juneCost))
                        },
                        Chart = new AnalysisPlanChart
                        {
                            Type = "column", Title = "Revenue EUR by month",
                            Categories = new List<string> { "2026-05", "2026-06" },
                            Series = new List<AnalysisPlanSeries>
                            {
                                new AnalysisPlanSeries
                                {
                                    Name = "Revenue EUR",
                                    FactIds = new List<string>
                                    { mayRevenue.FactId, juneRevenue.FactId }
                                }
                            }
                        }
                    },
                    new AnalysisPlanSlide
                    {
                        Id = "groups", Layout = "table",
                        Title = "June operating groups",
                        TableHeaders = new List<string> { "Group", "Revenue EUR" },
                        TableRows = new List<AnalysisPlanRow>
                        {
                            Row(Text("South"), Ref(southRevenue)),
                            Row(Text("West"), Ref(westRevenue))
                        }
                    },
                    new AnalysisPlanSlide
                    {
                        Id = "limits", Layout = "cards",
                        Title = "Evidence boundaries",
                        Cards = new List<AnalysisPlanCard>
                        {
                            new AnalysisPlanCard
                            {
                                Heading = "Source",
                                Points = new List<AnalysisPlanText>
                                { Word("Workbook ledger and verified totals") }
                            }
                        }
                    }
                }
            };

            var authored = new JavaScriptSerializer().Serialize(plan);
            Check(!authored.Contains("82992") && !authored.Contains("85519") &&
                !authored.Contains("22675"),
                "The hand-authored plan duplicated numerical facts.");
            var compiled = AnalysisDocumentCompiler.Compile(artifact, plan);
            Check(compiled.WorkbookRows[1][1].StartsWith("=SUMIF(") &&
                compiled.ExpectedFormulaFacts["B4"] == mayRevenue.FactId &&
                compiled.ExpectedFormulaFacts["C4"] == juneRevenue.FactId &&
                compiled.Slides.Count == 4,
                "The shared analysis did not retain live formula expectations and four slide plans.");
            var periodSlide = compiled.Slides[1];
            var table = (Dictionary<string, object>)periodSlide["table"];
            var rows = (object[])table["rows"];
            var chart = (Dictionary<string, object>)periodSlide["chart"];
            var series = (Dictionary<string, object>)((object[])chart["series"])[0];
            var values = (object[])series["values"];
            Check(((string[])rows[0])[1] == "85,519" &&
                ((string[])rows[0])[2] == "82,992" &&
                Convert.ToDouble(values[0]) == 85519d &&
                Convert.ToDouble(values[1]) == 82992d &&
                (string)periodSlide["sources"] == "WB01 / Ledger!A1:L145" &&
                !periodSlide.ContainsKey("footnote"),
                "Workbook, table, chart, and citations did not resolve from the same facts.");
            Check(((string)compiled.Slides[0]["subtitle"]).Contains("82,992") &&
                ((string)compiled.Slides[0]["sources"]).Contains("WB01"),
                "Headline prose or citation lost its verified fact binding.");

            plan.Slides[0].Subtitle.Add(Word(" and 9% growth"));
            var numericProseRejected = false;
            try { AnalysisDocumentCompiler.Compile(artifact, plan); }
            catch (InvalidOperationException error)
            { numericProseRejected = error.Message.Contains("NUMERIC_LITERAL_UNVERIFIED"); }
            Check(numericProseRejected,
                "A valid fact reference laundered an unsupported numeric claim.");
            plan.Slides[0].Subtitle.RemoveAt(plan.Slides[0].Subtitle.Count - 1);

            plan.Slides[0].Subtitle.Add(Word(" up 9 units"));
            RejectNumericProse(() => AnalysisDocumentCompiler.Compile(artifact, plan),
                "A one-digit number in authored slide prose bypassed fact binding.");
            plan.Slides[0].Subtitle.RemoveAt(plan.Slides[0].Subtitle.Count - 1);
            plan.Slides[0].Title = "June 82992 audit";
            RejectNumericProse(() => AnalysisDocumentCompiler.Compile(artifact, plan),
                "A numerical title bypassed fact binding.");
            plan.Slides[0].Title = "June sales audit";
            plan.Slides[0].Cards[0].Heading = "Revenue 82992";
            RejectNumericProse(() => AnalysisDocumentCompiler.Compile(artifact, plan),
                "A numerical card heading bypassed fact binding.");
            plan.Slides[0].Cards[0].Heading = "Revenue EUR";

            plan.Slides[1].TableRows[0].Cells[1] = Text("85,519");
            var numericTableRejected = false;
            try { AnalysisDocumentCompiler.Compile(artifact, plan); }
            catch (InvalidOperationException error)
            { numericTableRejected = error.Message.Contains("NUMERIC_LITERAL_UNVERIFIED"); }
            Check(numericTableRejected,
                "A slide table accepted an author-supplied numerical value.");
            plan.Slides[1].TableRows[0].Cells[1] = Ref(mayRevenue);

            plan.Slides[1].Chart.Categories[0] = "2026-04";
            var wrongPeriodRejected = false;
            try { AnalysisDocumentCompiler.Compile(artifact, plan); }
            catch (InvalidOperationException error)
            { wrongPeriodRejected = error.Message.Contains("CHART_PERIOD_MISMATCH"); }
            Check(wrongPeriodRejected,
                "A chart relabeled a verified May value as April.");
            plan.Slides[1].Chart.Categories[0] = "2026-05";

            plan.AnalysisId = "stale-analysis";
            var rejected = false;
            try { AnalysisDocumentCompiler.Compile(artifact, plan); }
            catch (InvalidOperationException error)
            { rejected = error.Message.Contains("PLAN_BINDING_INVALID"); }
            Check(rejected, "A stale plan was reused against another analysis revision.");

            plan.AnalysisId = artifact.AnalysisId;
            artifact.Assumptions.Add("source revision changed after planning");
            var changedArtifactRejected = false;
            try { AnalysisDocumentCompiler.Compile(artifact, plan); }
            catch (InvalidOperationException error)
            { changedArtifactRejected = error.Message.Contains("ANALYSIS_ID_MISMATCH"); }
            Check(changedArtifactRejected,
                "A mutated analysis retained a stale host-issued identity.");
        }

        private static VerifiedFact Fact(SourceSnapshot snapshot,
            SourceLocator locator, string metric, string value,
            string period)
        {
            return AnalysisContract.CreateObservedFact(snapshot.SnapshotId,
                metric, AnalysisContract.DecimalValue, value, value,
                "currency", "EUR", period,
                new Dictionary<string, string>(),
                new[] { locator }, AnalysisContract.Verified);
        }

        private static AnalysisPlanRow Row(params AnalysisPlanCell[] cells)
        { return new AnalysisPlanRow { Cells = cells.ToList() }; }

        private static AnalysisPlanCell Text(string value)
        { return new AnalysisPlanCell { Text = value }; }

        private static AnalysisPlanCell Ref(VerifiedFact fact)
        { return new AnalysisPlanCell { FactId = fact.FactId }; }

        private static AnalysisPlanCell Formula(string formula,
            VerifiedFact expected)
        { return new AnalysisPlanCell { Formula = formula,
            ExpectedFactId = expected.FactId }; }

        private static AnalysisPlanText Word(string value)
        { return new AnalysisPlanText { Text = value }; }

        private static AnalysisPlanText Value(VerifiedFact fact)
        { return new AnalysisPlanText { FactId = fact.FactId }; }

        private static void Check(bool value, string message)
        { if (!value) throw new Exception(message); }

        private static void RejectNumericProse(Action action, string message)
        {
            try { action(); }
            catch (InvalidOperationException error)
            {
                if (error.Message.Contains("NUMERIC_LITERAL_UNVERIFIED")) return;
            }
            throw new Exception(message);
        }
    }
}

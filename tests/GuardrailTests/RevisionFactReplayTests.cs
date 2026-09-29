using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web.Script.Serialization;
using Scribble.Office;

namespace GuardrailTests
{
    internal static class RevisionFactReplayTests
    {
        internal static void SeparateTableMetricAvailableWithPrimaryOnlyChart()
        {
            var locator = new SourceLocator {
                Kind = "excel_cell", SourceInstanceId = "replay-source",
                WorksheetIdentity = "Ledger", Cell = "D7"
            };
            Func<string, string, string, VerifiedFact> make =
                (metric, value, currency) =>
                    AnalysisContract.CreateObservedFact(
                        "snapshot-replay", metric,
                        AnalysisContract.DecimalValue, value, value,
                        currency == null ? string.Empty : "currency",
                        currency, "2026-06",
                        new Dictionary<string, string>(),
                        new[] { locator }, AnalysisContract.Verified);
            var primary = make("Items processed", "126", null);
            var secondary = make("Asset value GBP", "8450", "GBP");
            var bindingType = typeof(AnalysisContract).Assembly.GetType(
                "Scribble.Office.PresentationDraftCopy+MonthlyChartBinding",
                true);
            var binding = Activator.CreateInstance(bindingType, true);
            bindingType.GetField("Facts").SetValue(binding,
                new WorkbookMonthlyChartFacts.BoundSeries {
                    Names = new[] { "Items processed" },
                    Facts = new[] { primary }
                });
            bindingType.GetField("ContextFacts").SetValue(binding,
                new WorkbookMonthlyChartFacts.BoundSeries {
                    Names = new[] { "Items processed",
                        "Asset value GBP" },
                    Facts = new[] { primary, secondary }
                });
            var bindings = Array.CreateInstance(bindingType, 1);
            bindings.SetValue(binding, 0);
            var catalogType = typeof(AnalysisContract).Assembly.GetType(
                "Scribble.Office.RevisionFactCatalog", true);
            var fromBindings = catalogType.GetMethod("FromBindings",
                BindingFlags.Static | BindingFlags.NonPublic);
            var publicFacts = catalogType.GetMethod("PublicFacts",
                BindingFlags.Instance | BindingFlags.NonPublic);
            if (fromBindings == null || publicFacts == null)
                throw new Exception("Workbook repair fact catalog is missing.");
            var catalog = fromBindings.Invoke(null, new object[] {
                bindings
            });
            var visible = new JavaScriptSerializer().Serialize(
                publicFacts.Invoke(catalog, null));
            if (!visible.Contains("Items processed") ||
                !visible.Contains("Asset value GBP") ||
                !visible.Contains(secondary.FactId))
                throw new Exception("The separate table metric was omitted " +
                    "when the chart requested only its primary series.");
        }

        internal static void FactReferencesAndHistoricalResponses()
        {
            var locator = new SourceLocator {
                Kind = "excel_range",
                SourceInstanceId = "replay-source",
                WorksheetIdentity = "Ledger",
                Range = "H8:H145",
                Cell = "H145"
            };
            var fact = AnalysisContract.CreateObservedFact(
                "snapshot-replay", "Revenue EUR",
                AnalysisContract.DecimalValue, "82992", "82992",
                "currency", "EUR", "2026-06",
                new Dictionary<string, string>(),
                new[] { locator }, AnalysisContract.Verified);
            var change = AnalysisContract.CreateObservedFact(
                "snapshot-replay", "Revenue EUR change",
                AnalysisContract.DecimalValue, "-0.03", "-0.03",
                "ratio", null, "2026-06",
                new Dictionary<string, string>(),
                new[] { locator }, AnalysisContract.Verified);
            var catalogType = typeof(AnalysisContract).Assembly.GetType(
                "Scribble.Office.RevisionFactCatalog", true);
            var constructor = catalogType.GetConstructor(
                BindingFlags.Instance | BindingFlags.NonPublic, null,
                new[] { typeof(IEnumerable<VerifiedFact>) }, null);
            if (constructor == null)
                throw new Exception("Revision FactId catalog is missing.");
            var catalog = constructor.Invoke(new object[] {
                new[] { fact, change }
            });
            var render = catalogType.GetMethod("Render",
                BindingFlags.Instance | BindingFlags.NonPublic |
                BindingFlags.Public);
            var bind = catalogType.GetMethod("BindOperations",
                BindingFlags.Instance | BindingFlags.NonPublic |
                BindingFlags.Public);
            if (render == null || bind == null)
                throw new Exception("Revision FactId binder is missing.");
            var token = "[[fact:" + fact.FactId + ":";
            var rendered = (string)render.Invoke(catalog,
                new object[] { token + "metric]] " + token +
                    "value]] " + token + "unit]] in " + token +
                    "period]] (" + token + "locator]])" });
            if (rendered != "Revenue EUR 82,992 EUR in June 2026 " +
                "(Ledger!H8:H145)")
                throw new Exception("Revision text was not host rendered " +
                    "from its FactId.");
            Reject(render, catalog, "Revenue EUR 82,992 in June 2026",
                "REVISION_FACT_LITERAL_UNBOUND");
            Reject(render, catalog, "Revenue improved",
                "REVISION_FACT_LITERAL_UNBOUND");
            Reject(render, catalog, "Results in EUR",
                "REVISION_FACT_LITERAL_UNBOUND");
            if ((string)render.Invoke(catalog,
                new object[] { "A period change does not establish a cause." }) !=
                "A period change does not establish a cause.")
                throw new Exception("Ordinary prose was rejected as a metric label.");
            Reject(render, catalog, "[[fact:fact_000000000000000000000000:value]]",
                "REVISION_FACT_ID_UNKNOWN");
            try
            {
                bind.Invoke(catalog, new object[] { new object[] {
                    new Dictionary<string, object> {
                        { "kind", "replace_slide" },
                        { "slide", new Dictionary<string, object> {
                            { "sources", "Revenue EUR 82,992" }
                        } }
                    },
                    new Dictionary<string, object> {
                        { "kind", "replace_text" },
                        { "text", "June revenue" }
                    }
                } });
                throw new Exception("Visible citation text bypassed FactId binding.");
            }
            catch (TargetInvocationException error)
            {
                var message = error.InnerException?.Message ?? string.Empty;
                if (!message.StartsWith(
                        "REVISION_FACT_LITERAL_UNBOUND:",
                        StringComparison.Ordinal) ||
                    !message.Contains("operations[0].slide.sources") ||
                    !message.Contains("operations[1].text")) throw;
            }

            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "Fixtures", "pp01-model-replay.jsonl");
            var serializer = new JavaScriptSerializer {
                MaxJsonLength = 16000000
            };
            var runs = new HashSet<string>(StringComparer.Ordinal);
            var responses = 0;
            var empty = 0;
            var revisions = 0;
            var safelyRejected = 0;
            foreach (var line in File.ReadLines(path))
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                var record = serializer.DeserializeObject(line) as
                    Dictionary<string, object>;
                if (record == null ||
                    Convert.ToInt32(record["http_status"]) != 200 ||
                    Convert.ToString(record["response_sha256"]).Length !=
                        64)
                    throw new Exception("A historical response is invalid.");
                responses++;
                runs.Add(Convert.ToString(record["run"]));
                if ((bool)record["empty"]) empty++;
                foreach (var raw in (object[])record["tool_calls"])
                {
                    var tool = (Dictionary<string, object>)raw;
                    if (Convert.ToString(tool["name"]) !=
                        "revise_slides") continue;
                    revisions++;
                    Dictionary<string, object> args;
                    try
                    {
                        args = serializer.DeserializeObject(
                            Convert.ToString(tool["arguments"])) as
                            Dictionary<string, object>;
                    }
                    catch (ArgumentException)
                    {
                        safelyRejected++;
                        continue;
                    }
                    object operations;
                    if (args == null ||
                        !args.TryGetValue("operations", out operations) ||
                        !(operations is object[]))
                    {
                        safelyRejected++;
                        continue;
                    }
                    try
                    {
                        bind.Invoke(catalog, new[] { operations });
                    }
                    catch (TargetInvocationException error)
                    {
                        var failure = error.InnerException as
                            InvalidOperationException;
                        if (failure == null ||
                            !failure.Message.StartsWith(
                                "REVISION_FACT_",
                                StringComparison.Ordinal))
                            throw new Exception("Historical revision " +
                                "escaped typed preflight: " +
                                error.InnerException?.Message, error);
                        safelyRejected++;
                    }
                }
            }
            if (runs.Count != 16 || responses != 158 ||
                revisions != 73 || empty != 11 ||
                safelyRejected < 50)
                throw new Exception("The offline PP01 response replay is " +
                    "incomplete or skipped its unsafe proposals: " +
                    runs.Count + "/" + responses + "/" + revisions +
                    "/" + empty + "/" + safelyRejected);
        }

        internal static void NarrowPilotRevisionSchema()
        {
            var definition = Scribble.Chat.PresentationToolCatalog
                .PilotRevisionDefinition();
            var root = (Dictionary<string, object>)
                definition.function.parameters;
            var rootProperties = (Dictionary<string, object>)
                root["properties"];
            var operations = (Dictionary<string, object>)
                rootProperties["operations"];
            var operation = (Dictionary<string, object>)
                operations["items"];
            var fields = (Dictionary<string, object>)
                operation["properties"];
            var slide = (Dictionary<string, object>)fields["slide"];
            var slideFields = (Dictionary<string, object>)
                slide["properties"];
            if (fields.ContainsKey("series") ||
                fields.ContainsKey("category") ||
                fields.ContainsKey("new_index") ||
                slideFields.ContainsKey("chart") ||
                slideFields.ContainsKey("secondary_chart") ||
                slideFields.ContainsKey("footnote") ||
                slideFields.ContainsKey("sources") ||
                slideFields.ContainsKey("evidence") ||
                !slideFields.ContainsKey("cards") ||
                !slideFields.ContainsKey("table"))
                throw new Exception("Pilot revision exposes unsupported " +
                    "model-owned chart, layout or source metadata.");
        }

        private static void Reject(MethodInfo render, object catalog,
            string text, string expected)
        {
            try { render.Invoke(catalog, new object[] { text }); }
            catch (TargetInvocationException error)
            {
                if (error.InnerException?.Message == expected) return;
                throw;
            }
            throw new Exception("Revision FactId binder accepted " +
                "unbound or unknown text.");
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using Scribble.Chat;
using Scribble.Office;

namespace GuardrailTests
{
    internal static class AnalysisContractTests
    {
        private static void Check(bool condition, string message)
        {
            if (!condition) throw new Exception(message);
        }

        public static void ExplicitTableBindingsIssueOnlyVerifiedFacts()
        {
            Func<TableDataset, SourceSnapshot> snapshot = table =>
                AnalysisContract.CreateSnapshot("workbook-a",
                    "excel_workbook", "revision-1", "complete_range",
                    "literal_values", new[] { Locator("workbook-a",
                        "Ledger", "A1:D3", "") }, new[] { table });
            var binding = new AnalysisTableBinding
            {
                TableId = "ledger", PeriodHeader = "Period",
                DimensionHeaders = new List<string> { "Group" },
                Metrics = new List<AnalysisMetricColumnBinding>
                {
                    new AnalysisMetricColumnBinding { Header = "RevenueEUR",
                        Metric = "RevenueEUR", Unit = "currency",
                        Currency = "EUR" },
                    new AnalysisMetricColumnBinding { Header = "CostEUR",
                        Metric = "CostEUR", Unit = "currency",
                        Currency = "EUR" }
                }
            };
            var artifact = AnalysisTableArtifactBuilder.Build(
                snapshot(MappedTable()), binding);
            Check(artifact.Facts.Count == 4 &&
                artifact.Facts.Single(fact => fact.Metric == "RevenueEUR" &&
                    fact.Period == "2026-06").Value == "82992" &&
                artifact.Facts.Single(fact => fact.Metric == "CostEUR" &&
                    fact.Period == "2026-05").Locators.Single().Cell == "D2" &&
                artifact.Facts.All(fact => fact.Status ==
                    AnalysisContract.Verified &&
                    fact.Dimensions["Group"] == "North"),
                "Typed table binding lost period, dimension, source cell, or value.");
            var zero = MappedTable();
            var zeroCell = zero.Cells.Single(cell => cell.Reference == "C3");
            zeroCell.Value = "0";
            zeroCell.RawValue = "0";
            zeroCell.DisplayText = "0";
            Check(AnalysisTableArtifactBuilder.Build(snapshot(zero), binding)
                .Facts.Single(fact => fact.Metric == "RevenueEUR" &&
                    fact.Period == "2026-06").Value == "0",
                "A verified zero was treated as a missing metric.");
            var mislabeled = new AnalysisTableBinding
            {
                TableId = "ledger", PeriodHeader = "Period",
                Metrics = new List<AnalysisMetricColumnBinding>
                {
                    new AnalysisMetricColumnBinding { Header = "RevenueEUR",
                        Metric = "CostEUR", Unit = "currency",
                        Currency = "EUR" }
                }
            };
            RejectTableBinding(() => AnalysisTableArtifactBuilder.Build(
                snapshot(MappedTable()), mislabeled),
                "ANALYSIS_TABLE_BINDING_INVALID");
            binding.Metrics[0].Currency = "USD";
            RejectTableBinding(() => AnalysisTableArtifactBuilder.Build(
                snapshot(MappedTable()), binding),
                "ANALYSIS_TABLE_BINDING_INVALID");
            binding.Metrics[0].Currency = "EUR";
            binding.PeriodHeader = "Group";
            RejectTableBinding(() => AnalysisTableArtifactBuilder.Build(
                snapshot(MappedTable()), binding),
                "ANALYSIS_TABLE_BINDING_INVALID");
            binding.PeriodHeader = "Period";
            var duplicate = MappedTable();
            duplicate.Cells.Single(cell => cell.Reference == "A3").Value =
                "2026-05";
            RejectTableBinding(() => AnalysisTableArtifactBuilder.Build(
                snapshot(duplicate), binding), "ANALYSIS_TABLE_FACT_DUPLICATE");
            var cachedFormula = MappedTable();
            var formulaCell = cachedFormula.Cells.Single(cell =>
                cell.Reference == "C3");
            formulaCell.Formula = "=SUM(C2:C2)";
            formulaCell.Status = AnalysisContract.Unresolved;
            RejectTableBinding(() => AnalysisTableArtifactBuilder.Build(
                snapshot(cachedFormula), binding),
                "ANALYSIS_TABLE_VALUE_UNVERIFIED");
            var blank = MappedTable();
            var blankCell = blank.Cells.Single(cell => cell.Reference == "D3");
            blankCell.ValueType = AnalysisContract.MissingValue;
            blankCell.Value = string.Empty;
            blankCell.Status = AnalysisContract.Unresolved;
            RejectTableBinding(() => AnalysisTableArtifactBuilder.Build(
                snapshot(blank), binding),
                "ANALYSIS_TABLE_VALUE_UNVERIFIED");
            var ambiguous = MappedTable();
            ambiguous.Cells.Single(cell => cell.Reference == "D1").Value =
                "RevenueEUR";
            RejectTableBinding(() => AnalysisTableArtifactBuilder.Build(
                snapshot(ambiguous), binding),
                "ANALYSIS_TABLE_HEADER_AMBIGUOUS");
            var partial = AnalysisContract.CreateSnapshot("workbook-a",
                "excel_workbook", "revision-1", "partial_page",
                "literal_values", new[] { Locator("workbook-a",
                    "Ledger", "A1:D3", "") }, new[] { MappedTable() });
            RejectTableBinding(() => AnalysisTableArtifactBuilder.Build(
                partial, binding), "ANALYSIS_TABLE_SOURCE_UNSUPPORTED");
        }

        private static void RejectTableBinding(Action action, string code)
        {
            try { action(); }
            catch (InvalidOperationException error)
            {
                if (error.Message.Contains(code)) return;
                throw;
            }
            throw new Exception("Expected table binding failure " + code);
        }

        private static TableDataset MappedTable()
        {
            return new TableDataset
            {
                TableId = "ledger", Name = "Ledger", Rows = 3,
                Columns = 4, Cells = new List<DatasetCell>
                {
                    Cell(0, 0, "A1", AnalysisContract.TextValue, "Period"),
                    Cell(0, 1, "B1", AnalysisContract.TextValue, "Group"),
                    Cell(0, 2, "C1", AnalysisContract.TextValue, "RevenueEUR"),
                    Cell(0, 3, "D1", AnalysisContract.TextValue, "CostEUR"),
                    Cell(1, 0, "A2", AnalysisContract.TextValue, "2026-05"),
                    Cell(1, 1, "B2", AnalysisContract.TextValue, "North"),
                    Cell(1, 2, "C2", AnalysisContract.DecimalValue, "85519"),
                    Cell(1, 3, "D2", AnalysisContract.DecimalValue, "36702"),
                    Cell(2, 0, "A3", AnalysisContract.TextValue, "2026-06"),
                    Cell(2, 1, "B3", AnalysisContract.TextValue, "North"),
                    Cell(2, 2, "C3", AnalysisContract.DecimalValue, "82992"),
                    Cell(2, 3, "D3", AnalysisContract.DecimalValue, "36714")
                }
            };
        }

        public static void SnapshotIdentityInvalidationAndSerialization()
        {
            var table = Table();
            var locatorA = Locator("workbook-a", "Ledger", "A1:D2", "");
            var first = AnalysisContract.CreateSnapshot(
                "workbook-a",
                "excel_workbook",
                "revision-1",
                "complete_range",
                "recalculated",
                new[] { locatorA },
                new[] { table });
            var repeat = AnalysisContract.CreateSnapshot(
                "workbook-a",
                "excel_workbook",
                "revision-1",
                "complete_range",
                "recalculated",
                new[] { locatorA },
                new[] { Table() });
            var identicalOtherSource = AnalysisContract.CreateSnapshot(
                "workbook-b",
                "excel_workbook",
                "revision-1",
                "complete_range",
                "recalculated",
                new[] { Locator("workbook-b", "Ledger", "A1:D2", "") },
                new[] { Table() });
            var changedRevision = AnalysisContract.CreateSnapshot(
                "workbook-a",
                "excel_workbook",
                "revision-2",
                "complete_range",
                "recalculated",
                new[] { locatorA },
                new[] { Table() });

            Check(first.ContentHash == repeat.ContentHash &&
                first.SnapshotId == repeat.SnapshotId,
                "A stable source revision did not produce a stable snapshot binding.");
            Check(first.ContentHash == identicalOtherSource.ContentHash &&
                first.SnapshotId != identicalOtherSource.SnapshotId,
                "Identical workbook content collapsed two source instances.");
            Check(first.SnapshotId != changedRevision.SnapshotId &&
                !AnalysisContract.IsCurrent(first, "workbook-a", "revision-2",
                    changedRevision.ContentHash),
                "A changed source revision did not invalidate the prior snapshot.");
            Check(table.Cells.Count == 8 &&
                table.Cells.Single(cell => cell.Reference == "B2").ValueType ==
                    AnalysisContract.MissingValue,
                "A blank table cell was dropped instead of retaining its column position.");

            var revenue = Fact(first, "RevenueEUR", "82992", "EUR");
            var cost = Fact(first, "CostEUR", "36714", "EUR");
            VerifiedFact margin;
            var calculation = AnalysisCalculator.Calculate(
                AnalysisCalculator.Margin,
                new[] { revenue, cost },
                "GrossMargin",
                "2026-06",
                out margin);
            var artifact = AnalysisContract.CreateArtifact(
                new[] { first },
                new[] { revenue, cost, margin },
                new[] { calculation },
                new[] { "Revenue and cost use the same period." },
                new string[0]);
            var serialized = AnalysisContract.Serialize(artifact);
            var restored = AnalysisContract.Deserialize(serialized);
            Check(restored.AnalysisId == artifact.AnalysisId &&
                restored.Facts.Single(fact => fact.Metric == "GrossMargin")
                    .Value.StartsWith("0.557", StringComparison.Ordinal) &&
                restored.Snapshots[0].Tables[0].Cells.Count == 8,
                "The typed analysis did not survive persistence without value loss.");

            var root = Path.Combine(Path.GetTempPath(),
                "scribble-analysis-contract-" + Guid.NewGuid().ToString("N"));
            try
            {
                var request = new ChatCompletionRequest
                {
                    model = "offline-test",
                    messages = new List<object>
                    {
                        new ChatCompletionInputMessage
                        {
                            role = "user",
                            content = "Persist a typed analysis"
                        }
                    }
                };
                var task = new TaskContextManager(request, "excel",
                    "Persist a typed analysis", new TaskCheckpointStore(root));
                var evidenceId = task.PersistAnalysis(artifact);
                var loaded = task.LoadAnalysis();
                Check(task.State.AnalysisContractVersion ==
                        AnalysisContract.Version &&
                    task.State.AnalysisArtifactEvidenceId == evidenceId &&
                    loaded != null && loaded.AnalysisId == artifact.AnalysisId,
                    "Task persistence did not bind the typed analysis version and protected evidence.");

                var sources = new TaskSources(task);
                var firstText = sources.Add("Attached document", "Same source text", "attachment:0");
                var secondText = sources.Add("Attached document", "Same source text", "attachment:1");
                var repeatedFirst = sources.Add("Attached document", "Same source text", "attachment:0");
                var changedFirst = sources.Add("Attached document", "Changed source text", "attachment:0");
                Check(firstText.Count == 1 && secondText.Count == 1 &&
                    firstText[0] != secondText[0] &&
                    repeatedFirst.SequenceEqual(firstText) &&
                    changedFirst[0] != firstText[0] &&
                    sources.Resolve(firstText) == "Same source text" &&
                    sources.Resolve(secondText) == "Same source text" &&
                    sources.Spans().Count == 3,
                    "Identical attachments lost distinct provenance, or changed content reused an old source span.");
            }
            finally
            {
                if (Directory.Exists(root)) Directory.Delete(root, true);
            }

            restored.Facts.Single(fact => fact.Metric == "RevenueEUR").Value =
                "99999";
            var tamperRejected = false;
            try { AnalysisContract.Serialize(restored); }
            catch (InvalidOperationException error)
            {
                tamperRejected = error.Message.Contains("FACT_ID_MISMATCH");
            }
            Check(tamperRejected,
                "A fact value changed in place without invalidating its host-issued ID.");
        }

        public static void DeterministicCalculationsPreserveAuthority()
        {
            var snapshot = AnalysisContract.CreateSnapshot(
                "workbook-calculations",
                "excel_workbook",
                "revision-1",
                "complete_range",
                "recalculated",
                new[] { Locator("workbook-calculations", "Ledger", "A1:D2", "") },
                new[] { Table() });
            var june = Fact(snapshot, "RevenueEUR", "82992", "EUR");
            var may = Fact(snapshot, "RevenueEUR", "85519", "EUR");
            var groupA = Fact(snapshot, "RevenueEUR", "22675", "EUR");
            var groupB = Fact(snapshot, "RevenueEUR", "22044", "EUR");

            VerifiedFact difference;
            AnalysisCalculator.Calculate(AnalysisCalculator.Difference,
                new[] { june, may }, "RevenueChange", "2026-06", out difference);
            VerifiedFact growth;
            AnalysisCalculator.Calculate(AnalysisCalculator.Growth,
                new[] { june, may }, "RevenueGrowth", "2026-06", out growth);
            Check(difference.Value == "-2527" &&
                growth.Value.StartsWith("-0.0295", StringComparison.Ordinal),
                "Host arithmetic did not preserve the exact May-to-June direction.");

            var ranking = AnalysisCalculator.RankDescending(
                new[] { groupB, groupA });
            Check(ranking.Single(item => item.FactId == groupA.FactId).Rank == 1 &&
                ranking.Single(item => item.FactId == groupB.FactId).Rank == 2,
                "Deterministic ranking accepted the incorrect 22,044 > 22,675 objection.");

            var zero = Fact(snapshot, "RevenueEUR", "0", "EUR");
            var zeroRejected = false;
            try
            {
                VerifiedFact ignored;
                AnalysisCalculator.Calculate(AnalysisCalculator.Ratio,
                    new[] { june, zero }, "Ratio", "2026-06", out ignored);
            }
            catch (InvalidOperationException error)
            {
                zeroRejected = error.Message.Contains("ZERO_DENOMINATOR");
            }
            Check(zeroRejected, "A zero denominator produced a verified ratio.");

            var usd = Fact(snapshot, "Revenue", "100", "USD");
            var mismatchRejected = false;
            try
            {
                VerifiedFact ignored;
                AnalysisCalculator.Calculate(AnalysisCalculator.Sum,
                    new[] { june, usd }, "MixedRevenue", "2026-06", out ignored);
            }
            catch (InvalidOperationException error)
            {
                mismatchRejected = error.Message.Contains("UNIT_MISMATCH");
            }
            Check(mismatchRejected,
                "A mixed-currency calculation was promoted to a verified fact.");

            var missing = AnalysisContract.CreateObservedFact(
                snapshot.SnapshotId,
                "RevenueEUR",
                AnalysisContract.MissingValue,
                string.Empty,
                string.Empty,
                "currency",
                "EUR",
                "2026-07",
                new Dictionary<string, string>(),
                snapshot.Locators,
                AnalysisContract.Unresolved);
            var missingRejected = false;
            try
            {
                VerifiedFact ignored;
                AnalysisCalculator.Calculate(AnalysisCalculator.Sum,
                    new[] { june, missing }, "RevenueEUR", "2026-H2", out ignored);
            }
            catch (InvalidOperationException error)
            {
                missingRejected = error.Message.Contains("NOT_VERIFIED_NUMERIC");
            }
            Check(missingRejected,
                "A missing source value was silently treated as zero.");

            var values = new object[2, 4]
            {
                { "RowID", "Period", "Revenue", "Approved" },
                { "R1", 45808d, 82992d, true }
            };
            var formulas = new object[2, 4]
            {
                { "RowID", "Period", "Revenue", "Approved" },
                { "R1", 45808d, "=SUM(Source!C2:C10)", true }
            };
            var formats = new object[2, 4]
            {
                { "General", "General", "General", "General" },
                { "General", "yyyy-mm", "#,##0", "General" }
            };
            var displayed = new object[2, 4]
            {
                { "RowID", "Period", "Revenue", "Approved" },
                { "R1", "2025-05", "82,992", "TRUE" }
            };
            var typed = WorkbookTypedCapture.Capture("live", "Ledger",
                values, formulas, formats, displayed, 2, 4, 1, 1);
            var period = typed.Cells.Single(cell => cell.Reference == "B2");
            var formula = typed.Cells.Single(cell => cell.Reference == "C2");
            var approved = typed.Cells.Single(cell => cell.Reference == "D2");
            Check(period.ValueType == AnalysisContract.DateValue &&
                period.RawValue == "45808" &&
                period.DisplayText == "2025-05" &&
                formula.Formula == "=SUM(Source!C2:C10)" &&
                formula.RawValue == "82992" &&
                formula.Status == AnalysisContract.Unresolved &&
                approved.ValueType == AnalysisContract.BooleanValue &&
                approved.Value == "true",
                "Typed workbook capture lost raw values, display text, formulas, formats, or booleans.");

            var aggregateValues = new object[3, 3]
            {
                { "Group", "Period", "RevenueEUR" },
                { "A", "2026-06", 10d },
                { "A", "2026-06", 20d }
            };
            var aggregateFormulas = new object[3, 3]
            {
                { null, null, null },
                { null, null, null },
                { null, null, "=SUM(Source!C2:C10)" }
            };
            var aggregate = WorkbookTypedCapture.Capture(
                "typed-groups", "Ledger", aggregateValues,
                aggregateFormulas, "General", null, 3, 3, 1, 1);
            var totals = WorkbookGroupedTotals.Compute(aggregate,
                new[] { "Group" }, new[] { "RevenueEUR" },
                "Period", "2026-06");
            Check(totals.SourceRows == 2 && totals.MatchedRows == 2 &&
                totals.Groups == 1 && totals.SkippedCells == 1 &&
                totals.Table.Contains("2026-06\tA\t2\t10\t1"),
                "Typed grouped totals trusted an unresolved formula cache or lost its skipped-cell disclosure.");
        }

        public static void OpenXmlCaptureRetainsTypedCells()
        {
            var root = Path.Combine(Path.GetTempPath(),
                "scribble-openxml-analysis-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try
            {
                var path = Path.Combine(root, "typed.xlsx");
                using (var archive = ZipFile.Open(path, ZipArchiveMode.Create))
                {
                    Entry(archive, "xl/workbook.xml",
                        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
                        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets>" +
                        "<sheet name=\"Ledger\" sheetId=\"2\" r:id=\"rId2\"/>" +
                        "<sheet name=\"Notes\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
                    Entry(archive, "xl/_rels/workbook.xml.rels",
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"rId1\" Target=\"worksheets/sheet1.xml\" " +
                        "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\"/>" +
                        "<Relationship Id=\"rId2\" Target=\"worksheets/sheet2.xml\" " +
                        "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\"/>" +
                        "</Relationships>");
                    Entry(archive, "xl/sharedStrings.xml",
                        "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
                        "<si><t>Period</t></si><si><t>Revenue</t></si><si><t>Approved</t></si></sst>");
                    Entry(archive, "xl/styles.xml",
                        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
                        "<numFmts count=\"1\"><numFmt numFmtId=\"164\" formatCode=\"#,##0 &quot;EUR&quot;\"/></numFmts>" +
                        "<cellXfs count=\"3\"><xf numFmtId=\"0\"/><xf numFmtId=\"14\"/><xf numFmtId=\"164\"/></cellXfs>" +
                        "</styleSheet>");
                    Entry(archive, "xl/worksheets/sheet1.xml",
                        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>" +
                        "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Source notes</t></is></c></row>" +
                        "</sheetData></worksheet>");
                    Entry(archive, "xl/worksheets/sheet2.xml",
                        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>" +
                        "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c><c r=\"B1\" t=\"s\"><v>1</v></c>" +
                        "<c r=\"D1\" t=\"s\"><v>2</v></c></row>" +
                        "<row r=\"2\"><c r=\"A2\" s=\"1\"><v>45808</v></c>" +
                        "<c r=\"B2\" s=\"2\"><f>SUM(Source!B2:B10)</f><v>82992</v></c>" +
                        "<c r=\"C2\"/><c r=\"D2\" t=\"b\"><v>1</v></c></row>" +
                        "<row r=\"3\"><c r=\"A3\" t=\"b\"><v>maybe</v></c></row>" +
                        "</sheetData></worksheet>");
                }
                var snapshot = OpenXmlWorkbookSnapshotReader.Capture(path,
                    "attachment:typed", "revision-1", CancellationToken.None);
                var table = snapshot.Tables[0];
                var date = table.Cells.Single(cell => cell.Reference == "A2");
                var formula = table.Cells.Single(cell => cell.Reference == "B2");
                var blank = table.Cells.Single(cell => cell.Reference == "C2");
                var approved = table.Cells.Single(cell => cell.Reference == "D2");
                var malformed = table.Cells.Single(cell => cell.Reference == "A3");
                Check(snapshot.Tables.Count == 2 && table.Name == "Ledger" &&
                    snapshot.Tables[1].Name == "Notes" && table.Rows == 3 &&
                    table.Columns == 4 && date.ValueType == AnalysisContract.DateValue &&
                    date.RawValue == "45808" && date.NumberFormat == "m/d/yy" &&
                    formula.Formula == "=SUM(Source!B2:B10)" &&
                    formula.RawCellType == "n" && formula.RawValue == "82992" &&
                    formula.NumberFormat.Contains("EUR") &&
                    formula.Status == AnalysisContract.Unresolved &&
                    blank.ValueType == AnalysisContract.MissingValue &&
                    approved.ValueType == AnalysisContract.BooleanValue &&
                    malformed.ValueType == AnalysisContract.ErrorValue &&
                    malformed.RawCellType == "b" &&
                    malformed.Status == AnalysisContract.Unresolved &&
                    snapshot.CalculationState == "cached_formula_values_unverified",
                    "OpenXML typed capture lost a sheet name, blank column, date serial, formula, format, boolean, malformed value, or cache status.");
            }
            finally
            {
                if (Directory.Exists(root)) Directory.Delete(root, true);
            }
        }

        private static void Entry(
            ZipArchive archive,
            string name,
            string content)
        {
            var entry = archive.CreateEntry(name);
            using (var writer = new StreamWriter(entry.Open(),
                new UTF8Encoding(false))) writer.Write(content);
        }

        private static TableDataset Table()
        {
            return new TableDataset
            {
                TableId = "ledger",
                Name = "Ledger",
                Rows = 2,
                Columns = 4,
                Cells = new List<DatasetCell>
                {
                    Cell(0, 0, "A1", AnalysisContract.TextValue, "RowID"),
                    Cell(0, 1, "B1", AnalysisContract.TextValue, "Group"),
                    Cell(0, 2, "C1", AnalysisContract.TextValue, "RevenueEUR"),
                    Cell(0, 3, "D1", AnalysisContract.TextValue, "Period"),
                    Cell(1, 0, "A2", AnalysisContract.TextValue, "R1"),
                    Cell(1, 1, "B2", AnalysisContract.MissingValue, string.Empty),
                    Cell(1, 2, "C2", AnalysisContract.DecimalValue, "0"),
                    Cell(1, 3, "D2", AnalysisContract.TextValue, "2026-06")
                }
            };
        }

        private static DatasetCell Cell(
            int row,
            int column,
            string reference,
            string valueType,
            string value)
        {
            return new DatasetCell
            {
                Row = row,
                Column = column,
                Reference = reference,
                ValueType = valueType,
                RawCellType = valueType == AnalysisContract.MissingValue
                    ? "blank"
                    : valueType == AnalysisContract.DecimalValue ||
                      valueType == AnalysisContract.IntegerValue
                        ? "number"
                        : "text",
                RawValue = value,
                Value = value,
                DisplayText = value,
                Formula = string.Empty,
                NumberFormat = "General",
                Status = valueType == AnalysisContract.MissingValue
                    ? AnalysisContract.Unresolved
                    : AnalysisContract.Verified
            };
        }

        private static SourceLocator Locator(
            string instance,
            string sheet,
            string range,
            string cell)
        {
            return new SourceLocator
            {
                Kind = "excel_range",
                SourceInstanceId = instance,
                WorksheetIdentity = sheet,
                Range = range,
                Cell = cell
            };
        }

        private static VerifiedFact Fact(
            SourceSnapshot snapshot,
            string metric,
            string value,
            string currency)
        {
            return AnalysisContract.CreateObservedFact(
                snapshot.SnapshotId,
                metric,
                AnalysisContract.DecimalValue,
                value,
                value,
                "currency",
                currency,
                "2026-06",
                new Dictionary<string, string>(),
                snapshot.Locators,
                AnalysisContract.Verified);
        }
    }
}

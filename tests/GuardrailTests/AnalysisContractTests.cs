using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
                approved.ValueType == AnalysisContract.BooleanValue &&
                approved.Value == "true",
                "Typed workbook capture lost raw values, display text, formulas, formats, or booleans.");
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

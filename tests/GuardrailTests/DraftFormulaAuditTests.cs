using System;
using System.Collections.Generic;
using System.Linq;
using Scribble.Office;

namespace GuardrailTests
{
    // The Excel draft table is written at A3, so a model that counts rows
    // loosely produces formulas that evaluate but describe the wrong cells.
    // The audit must reject those before the one-shot write permission is
    // spent, and must leave legitimate cross-row formulas alone.
    internal static class DraftFormulaAuditTests
    {
        private static void Check(bool value, string message) { if (!value) throw new Exception(message); }
        private static IReadOnlyList<IReadOnlyList<string>> Rows(params string[][] rows)
        { return rows.Select(row => (IReadOnlyList<string>)row.ToList()).ToList(); }

        internal static void RowMismatchesAreRejectedBeforeWriting()
        {
            // Sheet rows: 3 header, 4 revenue, 5 cost, 6 gross profit, 7 margin,
            // 8 spacer, 9 section header, 10-13 groups, 14 totals, 15 share.
            var shifted = Rows(
                new[] { "Metric", "May", "June" },
                new[] { "Revenue EUR", "=Ledger!F2*2", "='Scribble Draft'!C4" },
                new[] { "Cost EUR", "36702", "36714" },
                new[] { "Gross profit EUR", "=B4-B5", "=C4-C5" },
                new[] { "Gross margin", "=B6/B4", "=IF(C4=\"\",\"\",C6/C4)" },
                new[] { "" },
                new[] { "Group", "Revenue EUR", "Cost EUR", "Margin" },
                new[] { "North", "1000", "400", "=(B10-C10)/B10" },
                new[] { "South", "1200", "500", "=(B10-B11)/B10" },
                new[] { "East", "900", "300", "=(B12-B13)/B12" },
                new[] { "West", "800", "200", "=(B13-C13)/B13" },
                new[] { "Total", "=SUM(B9:B13)", "=SUM(C10:C14)", "=SUM(B10:B13)" },
                new[] { "Share North", "=B10/B$14", "=SUM(B3:B13)", "=SUM(B10:B20)" });
            var findings = DraftTableFormulaAudit.Findings(shifted);
            var text = string.Join("\n", findings);
            Check(findings.Any(f => f.StartsWith("D11 ") && f.Contains("mixes row 10") && f.Contains("'South'")), "Shifted same-row rate D11 was not reported: " + text);
            Check(findings.Any(f => f.StartsWith("D12 ") && f.Contains("mixes row 13")), "Shifted same-row rate D12 was not reported: " + text);
            Check(findings.Any(f => f.StartsWith("B14 ") && f.Contains("text cell B9")), "Total over the section header was not reported: " + text);
            Check(findings.Any(f => f.StartsWith("C14 ") && f.Contains("circular")), "Circular total C14 was not reported: " + text);
            Check(findings.Any(f => f.StartsWith("D14 ") && f.Contains("totals column B from column D") && f.Contains("=SUM(D10:D13)")), "Cross-column total D14 was not reported: " + text);
            Check(findings.Any(f => f.StartsWith("C15 ") && f.Contains("header row 3")), "Range over the header row was not reported: " + text);
            Check(findings.Any(f => f.StartsWith("D15 ") && f.Contains("B20") && f.Contains("outside")), "Reference beyond the table was not reported: " + text);
            Check(!findings.Any(f => f.StartsWith("B4 ") || f.StartsWith("C4 ") || f.StartsWith("B6 ") || f.StartsWith("C6 ") || f.StartsWith("B7 ") ||
                f.StartsWith("C7 ") || f.StartsWith("D10 ") || f.StartsWith("D13 ") || f.StartsWith("B15 ")),
                "A legitimate cross-row, anchored, sheet-qualified or same-row formula was reported: " + text);
            Check(findings.Count == 7, "Unexpected finding count " + findings.Count + ": " + text);

            try { DraftTableFormulaAudit.Require(shifted); throw new Exception("Expected the audit to reject the draft."); }
            catch (InvalidOperationException ex)
            {
                Check(ex.Message.StartsWith(DraftTableFormulaAudit.ErrorCode) && ex.Message.Contains("first data row is row 4") &&
                    ex.Message.Contains("no write permission was spent") && ex.Message.Contains("D11 ") && ex.Message.Contains("(1 more.)"),
                    "Rejection did not carry actionable, bounded guidance: " + ex.Message);
            }

            // Period rows compare with their neighbour; running figures span rows
            // with different columns; the header row is never mixed into a rate.
            var periods = Rows(
                new[] { "Period", "Revenue EUR", "Change EUR", "Running EUR" },
                new[] { "2026-05", "85519", "", "=B4" },
                new[] { "2026-06", "82992", "=B5-B4", "=D4+B5" },
                new[] { "Total", "=SUM(B4:B5)", "", "" });
            Check(DraftTableFormulaAudit.Findings(periods).Count == 0, "Legitimate period comparison was rejected: " + string.Join("\n", DraftTableFormulaAudit.Findings(periods)));
            DraftTableFormulaAudit.Require(periods);
            var headerMix = Rows(new[] { "Metric", "May", "June" }, new[] { "Revenue EUR", "1", "=B4-B3" });
            var mixed = DraftTableFormulaAudit.Findings(headerMix);
            Check(mixed.Count == 1 && mixed[0].Contains("(the header)"), "A rate mixing the header row was not reported: " + string.Join("\n", mixed));
            Check(DraftTableFormulaAudit.Findings(null).Count == 0 && DraftTableFormulaAudit.Findings(Rows(new[] { "Only", "text" })).Count == 0,
                "Formula-free drafts produced findings.");
        }
    }
}

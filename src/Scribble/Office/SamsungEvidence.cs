using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace Scribble.Office
{
    public static class SamsungEvidence
    {
        // A displayed value that is not printed in the source is never accepted
        // from model prose. This recovery text points the model at the existing
        // host-recomputed contract instead of a clarification request.
        public const string DerivedValueGuidance =
            " A percentage, margin, share, difference or total that is not printed verbatim in the cited source is a derived value. " +
            "Declare it in calculations so the host recomputes it from cited operands: percent = a/b*100, growth_percent = (a-b)/b*100, " +
            "margin_percent = (a-b)/a*100 (gross margin: operands Revenue then Cost for one period, unit %, result rounded to decimals). " +
            "A source fraction such as 0.5576 does not verify a displayed 55.76%; use the calculation. " +
            "Otherwise omit the derived value when the request does not require it. Do not call ask_user about number verification; repair the slide and continue the remaining planned IDs.";

        private static string Normalize(string value) { return Regex.Replace(value ?? "", @"\s+", " ").Trim(); }
        private static bool AssociationOccurs(string passage, string association, string key)
        {
            if (Normalize(passage).IndexOf(Normalize(association), StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (!string.Equals(key, "period", StringComparison.Ordinal)) return false;

            var expected = CanonicalPeriods(association);
            if (expected.Count != 1) return false;
            if (CanonicalPeriods(passage).Contains(expected.Single())) return true;
            return MonthHeaderOccurs(passage, expected.Single());
        }

        private static readonly string[] MonthNames = { "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December" };

        // An audit table often heads its columns "May" and "June" while the
        // deck must label them 2026-05 and 2026-06. The capitalized month name
        // identifies the column; the passage must not name a different year.
        private static bool MonthHeaderOccurs(string passage, string canonicalPeriod)
        {
            var year = canonicalPeriod.Substring(0, 4);
            var name = MonthNames[int.Parse(canonicalPeriod.Substring(5, 2), CultureInfo.InvariantCulture) - 1];
            if (!Regex.IsMatch(passage ?? "", @"(?<![A-Za-z])(?:" + name + "|" + name.Substring(0, 3) + @")(?![A-Za-z])")) return false;
            return !Regex.Matches(passage ?? "", @"(?<![0-9])(?:19|20)[0-9]{2}(?![0-9])").Cast<Match>().Any(match => match.Value != year);
        }

        // Period labels are not quantities. A displayed YYYY-MM or "June 2026"
        // is verified against every source the task has read; the numbers
        // beside it remain bound to the slide's cited evidence.
        public static string RemoveVerifiedPeriodLabels(string displayed, string taskSources)
        {
            var known = CanonicalPeriods(taskSources);
            if (known.Count == 0) return displayed ?? "";
            return PeriodPattern.Replace(displayed ?? "", match =>
            {
                var canonical = CanonicalPeriods(match.Value);
                return canonical.Count == 1 && known.Contains(canonical.Single()) ? " " : match.Value;
            });
        }

        // Models sometimes quote the exact data row but omit an adjacent period
        // header. Expand only inside the already resolved, host-verified source
        // evidence, and only to the shortest contiguous line window that contains
        // the model's exact quote plus every declared association. This cannot
        // introduce external or model-authored evidence.
        private static string ExpandClaimPassage(string evidence, string passage,
            IDictionary<string, object> claim)
        {
            var associations = new[] { "label", "unit", "period" }
                .Select(key => new { Key = key, Value = SamsungAuthoringPolicy.Text(claim, key) })
                .Where(item => !string.IsNullOrWhiteSpace(item.Value) &&
                    !string.Equals(item.Value, "not applicable", StringComparison.OrdinalIgnoreCase))
                .ToArray();
            if (associations.All(item => AssociationOccurs(passage, item.Value, item.Key))) return passage;
            var quoted = Normalize(passage);
            if (quoted.Length == 0) return passage;
            var lines = Regex.Split(evidence ?? "", @"\r\n|\n|\r");
            string best = null;
            for (var start = 0; start < lines.Length; start++)
            {
                for (var end = start; end < lines.Length && end - start < 24; end++)
                {
                    var candidate = string.Join("\n", lines.Skip(start).Take(end - start + 1));
                    var normalized = Normalize(candidate);
                    if (normalized.Length > 2400) break;
                    if (normalized.IndexOf(quoted, StringComparison.Ordinal) < 0 ||
                        associations.Any(item => !AssociationOccurs(candidate, item.Value, item.Key))) continue;
                    if (best == null || normalized.Length < Normalize(best).Length) best = candidate;
                }
            }
            return best ?? passage;
        }

        private static readonly Regex PeriodPattern = new Regex(
            @"\b(?:19|20)\d{2}[-/](?:0?[1-9]|1[0-2])\b|" +
            @"\b(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|" +
            @"Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|" +
            @"Dec(?:ember)?)\s+(?:19|20)\d{2}\b|" +
            @"\b(?:19|20)\d{2}\s+(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|" +
            @"Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|" +
            @"Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\b",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static HashSet<string> CanonicalPeriods(string value)
        {
            var result = new HashSet<string>(StringComparer.Ordinal);
            var text = Normalize(value);
            foreach (Match match in PeriodPattern.Matches(text))
            {
                DateTime parsed;
                if (DateTime.TryParse(match.Value, CultureInfo.InvariantCulture,
                    DateTimeStyles.AllowWhiteSpaces, out parsed))
                    result.Add(parsed.ToString("yyyy-MM", CultureInfo.InvariantCulture));
            }
            return result;
        }
        // A table citation is often its header line plus one later row of the
        // same block ("Metric May June" + "Cost EUR 36702 36714"), skipping
        // the rows between. Every cited line is verbatim, so the host resolves
        // it to the complete contiguous block rather than rejecting it. A
        // passage containing anything that is not a whole source line within
        // that block is still refused.
        private static string StitchedBlock(string evidence, string passage)
        {
            var wanted = Normalize(passage);
            if (wanted.Length == 0) return null;
            var lines = Regex.Split(evidence ?? "", @"\r\n|\n|\r").Select(Normalize).ToArray();
            for (var start = 0; start < lines.Length; start++)
            {
                if (lines[start].Length == 0 || !wanted.StartsWith(lines[start], StringComparison.Ordinal)) continue;
                var remaining = wanted.Substring(lines[start].Length).TrimStart();
                var last = start;
                for (var index = start + 1; index < lines.Length && index - start < 12 && remaining.Length > 0; index++)
                {
                    var line = lines[index];
                    if (line.Length == 0 || !remaining.StartsWith(line, StringComparison.Ordinal)) continue;
                    var rest = remaining.Substring(line.Length);
                    if (rest.Length > 0 && rest[0] != ' ') continue;
                    remaining = rest.TrimStart(); last = index;
                }
                if (remaining.Length == 0 && last > start)
                    return string.Join("\n", Regex.Split(evidence ?? "", @"\r\n|\n|\r").Skip(start).Take(last - start + 1));
            }
            return null;
        }

        private static string RequirePassage(string passage, string evidence)
        {
            if (!string.IsNullOrWhiteSpace(passage) && Normalize(evidence).Contains(Normalize(passage))) return passage;
            var stitched = StitchedBlock(evidence, passage);
            if (stitched != null) return stitched;
            {
                var rejected = Normalize(passage);
                if (rejected.Length > 180) rejected = rejected.Substring(0, 180) + "...";
                var suggestion = ClosestVerifiedTablePassage(evidence, passage);
                var guidance = string.IsNullOrWhiteSpace(suggestion) ? "" :
                    " A nearby host-verified passage is: \"" + suggestion +
                    "\". Copy it exactly and use its literal header labels for label, unit and period; split comparisons into one claim per period.";
                throw new InvalidOperationException("SLIDE_ASSOCIATION_UNVERIFIED: Association evidence must be an exact verbatim passage from the slide's resolved source evidence (whitespace may differ). Recopy or remove this rejected passage: \"" + rejected + "\"." + guidance);
            }
        }
        private static string ClosestVerifiedTablePassage(string evidence, string rejectedPassage)
        {
            var lines = Regex.Split(evidence ?? "", @"\r\n|\n|\r");
            var wanted = Regex.Matches(Normalize(rejectedPassage), @"(?<![A-Za-z0-9])[-+]?(?:\d+(?:[,.]\d+)*|\.\d+)")
                .Cast<Match>().Select(match => match.Value.Replace(",", ""))
                .Where(value => value.Length >= 3).Distinct(StringComparer.Ordinal).ToArray();
            if (wanted.Length == 0 || lines.Length == 0) return "";

            var bestStart = -1;
            var bestEnd = -1;
            var bestScore = 0;
            var bestLength = int.MaxValue;
            for (var start = 0; start < lines.Length; start++)
            {
                for (var end = start; end < lines.Length && end - start < 5; end++)
                {
                    var candidate = Normalize(string.Join("\n", lines.Skip(start).Take(end - start + 1))).Replace(",", "");
                    var score = wanted.Count(value => candidate.IndexOf(value, StringComparison.Ordinal) >= 0);
                    if (score == 0 || score < bestScore || (score == bestScore && candidate.Length >= bestLength)) continue;
                    bestStart = start;
                    bestEnd = end;
                    bestScore = score;
                    bestLength = candidate.Length;
                }
            }
            if (bestStart < 0) return "";

            // Include the nearest nonnumeric, multi-column header so the retry has
            // the literal period/unit labels needed for an unambiguous claim.
            for (var index = bestStart - 1; index >= 0 && bestStart - index <= 8; index--)
            {
                var cells = lines[index].Split('\t');
                if (cells.Length < 2 || cells.All(string.IsNullOrWhiteSpace)) continue;
                if (Regex.IsMatch(lines[index], @"\d")) continue;
                bestStart = index;
                break;
            }
            var result = Normalize(string.Join("\n", lines.Skip(bestStart).Take(bestEnd - bestStart + 1)));
            return result.Length <= 700 ? result : result.Substring(0, 697) + "...";
        }
        private static decimal Number(IDictionary<string, object> map, string key)
        {
            object raw; decimal value;
            if (!map.TryGetValue(key, out raw) || raw is bool || raw is string ||
                !decimal.TryParse(Convert.ToString(raw, CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture, out value))
                throw new InvalidOperationException("SLIDE_CALCULATION_INVALID: " + key + " must be a finite numeric value.");
            return value;
        }
        public static IEnumerable<string> ValidateCalculations(IDictionary<string, object> slide, string evidence)
        {
            var calculations = SamsungAuthoringPolicy.Array(slide, "calculations");
            if (calculations.Length > 24) throw new InvalidOperationException("Too many calculations for one slide.");
            foreach (var raw in calculations)
            {
                var calc = SamsungAuthoringPolicy.ReadMap(raw);
                var operands = SamsungAuthoringPolicy.Array(calc, "operands").Select(SamsungAuthoringPolicy.ReadMap).ToArray();
                if (operands.Length == 0 || operands.Length > 24) throw new InvalidOperationException("A calculation needs 1 to 24 source operands.");
                foreach (var operand in operands)
                {
                    var passage = SamsungAuthoringPolicy.Text(operand, "evidence");
                    passage = RequirePassage(passage, evidence);
                    // Like a claim, an exact data row may omit its adjacent header.
                    passage = ExpandClaimPassage(evidence, passage, operand);
                    operand["evidence"] = passage;
                    foreach (var key in new[] { "label", "unit", "period" })
                        if (string.IsNullOrWhiteSpace(SamsungAuthoringPolicy.Text(operand, key)) ||
                            !AssociationOccurs(passage,
                                SamsungAuthoringPolicy.Text(operand, key), key))
                        {
                            var cited = Normalize(passage);
                            if (cited.Length > 160) cited = cited.Substring(0, 160) + "...";
                            var nearby = ClosestVerifiedTablePassage(evidence, passage);
                            throw new InvalidOperationException("SLIDE_OPERAND_ASSOCIATION: Calculation '" + SamsungAuthoringPolicy.Text(calc, "label") +
                                "' operand " + SamsungAuthoringPolicy.Text(operand, "value") + " cites '" + cited + "' but that passage is missing " + key + " '" +
                                SamsungAuthoringPolicy.Text(operand, key) + "'. Cite one contiguous block containing the table header and the operand's row, and use the literal " +
                                key + " wording of that block (for a column headed June, period June or 2026-06 both match)." +
                                (string.IsNullOrWhiteSpace(nearby) ? "" : " A nearby host-verified passage is: \"" + nearby + "\"."));
                        }
                    var number = Number(operand, "value");
                    var found = Regex.Matches(passage, @"(?<![A-Za-z0-9])[-+]?(?:\d+(?:[,.]\d+)*|\.\d+)").Cast<Match>()
                        .Any(m => decimal.TryParse(m.Value.Replace(",", ""), NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed) && parsed == number);
                    if (!found) throw new InvalidOperationException("SLIDE_OPERAND_UNVERIFIED: Operand value is absent from its cited passage.");
                }
                var values = operands.Select(o => Number(o, "value")).ToArray();
                var operation = SamsungAuthoringPolicy.Text(calc, "operation");
                decimal result;
                try
                {
                    switch (operation)
                    {
                        case "sum": result = values.Sum(); break;
                        case "difference": RequireTwo(values); result = values[0] - values[1]; break;
                        case "ratio": RequireTwo(values); result = values[0] / values[1]; break;
                        case "percent": RequireTwo(values); result = values[0] / values[1] * 100; break;
                        case "growth_percent": RequireTwo(values); result = (values[0] - values[1]) / values[1] * 100; break;
                        // Margin on the first operand: (Revenue - Cost) / Revenue.
                        case "margin_percent": RequireTwo(values); result = (values[0] - values[1]) / values[0] * 100; break;
                        default: throw new InvalidOperationException("Unsupported calculation operation.");
                    }
                }
                catch (Exception ex) when (ex is OverflowException || ex is DivideByZeroException)
                { throw new InvalidOperationException("SLIDE_CALCULATION_INVALID: Overflow or zero baseline."); }
                var rounding = Number(calc, "decimals");
                if (rounding < 0 || rounding > 6 || rounding != decimal.Truncate(rounding)) throw new InvalidOperationException("Rounding must be 0 to 6 decimal places.");
                result = Math.Round(result, (int)rounding, MidpointRounding.AwayFromZero);
                if (result != Number(calc, "result"))
                    throw new InvalidOperationException("SLIDE_CALCULATION_MISMATCH: Calculation '" + SamsungAuthoringPolicy.Text(calc, "label") +
                        "' declares " + Number(calc, "result").ToString(CultureInfo.InvariantCulture) + " but host arithmetic on the cited operands gives " +
                        result.ToString(CultureInfo.InvariantCulture) + ". Use the host value in result and everywhere this metric is displayed.");
                var unit = SamsungAuthoringPolicy.Text(calc, "unit");
                if (string.IsNullOrWhiteSpace(unit) || ((operation == "percent" || operation == "growth_percent" || operation == "margin_percent") && unit != "%"))
                    throw new InvalidOperationException("SLIDE_CALCULATION_UNIT: Specify result units; percentage operations require %.");
                if ((operation == "ratio" || operation == "percent" || operation == "growth_percent" || operation == "margin_percent") &&
                    operands.Select(o => SamsungAuthoringPolicy.Text(o, "unit")).Distinct().Count() != 1)
                    throw new InvalidOperationException("SLIDE_CALCULATION_UNIT: Ratio/growth operands must use the same units; convert units explicitly in the source first.");
                if (operation == "growth_percent" && values[1] <= 0)
                    throw new InvalidOperationException("SLIDE_CALCULATION_BASELINE: Growth percentages require a positive baseline; describe the absolute change otherwise.");
                if (operation == "margin_percent" && values[0] <= 0)
                    throw new InvalidOperationException("SLIDE_CALCULATION_BASELINE: Margin percentages require a positive first operand (revenue); describe the absolute difference otherwise.");
                if ((operation == "sum" || operation == "difference") && operands.Any(o => SamsungAuthoringPolicy.Text(o, "unit") != unit))
                    throw new InvalidOperationException("SLIDE_CALCULATION_UNIT: Addition and subtraction require matching operand/result units.");
                yield return ((double)result).ToString("R", CultureInfo.InvariantCulture);
            }
        }
        private static void RequireTwo(decimal[] values)
        { if (values.Length != 2) throw new InvalidOperationException("This operation requires exactly two ordered operands."); }

        public static void ValidateClaims(IDictionary<string, object> slide, string evidence)
        {
            var claims = SamsungAuthoringPolicy.Array(slide, "claims");
            if (claims.Length > 24) throw new InvalidOperationException("Too many claim associations.");
            var associationFailures = new List<string>();
            foreach (var raw in claims)
            {
                var claim = SamsungAuthoringPolicy.ReadMap(raw);
                var claimText = SamsungAuthoringPolicy.Text(claim, "text");
                if (string.IsNullOrWhiteSpace(claimText)) throw new InvalidOperationException("Claim text is required.");
                var passage = SamsungAuthoringPolicy.Text(claim, "evidence");
                passage = RequirePassage(passage, evidence);
                passage = ExpandClaimPassage(evidence, passage, claim);
                claim["evidence"] = passage;
                // Semantic association is checked separately by the source reviewer.
                foreach (var key in new[] { "label", "unit", "period" })
                {
                    var association = SamsungAuthoringPolicy.Text(claim, key);
                    if (string.IsNullOrWhiteSpace(association))
                        throw new InvalidOperationException("Claim associations need label, unit and period (use 'not applicable' for qualitative claims).");
                    if (!string.Equals(association, "not applicable", StringComparison.OrdinalIgnoreCase) &&
                        !AssociationOccurs(passage, association, key))
                    {
                        var cited = Normalize(passage);
                        if (cited.Length > 120) cited = cited.Substring(0, 120) + "...";
                        associationFailures.Add(
                            "Claim '" + claimText + "' cites '" + cited +
                            "' but that passage is missing " + key +
                            " '" + association + "'.");
                    }
                }
            }
            if (associationFailures.Count > 0)
                throw new InvalidOperationException(
                    "SLIDE_CLAIM_ASSOCIATION: Every claim's one exact cited passage must contain its own value, label, unit and period. Fix all listed citations in one retry by copying a longer contiguous source block (include table headers and the claimed row): " +
                    string.Join(" ", associationFailures.Take(12)));
        }
    }
}

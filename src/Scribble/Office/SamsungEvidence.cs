using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace Scribble.Office
{
    public static class SamsungEvidence
    {
        private static string Normalize(string value) { return Regex.Replace(value ?? "", @"\s+", " ").Trim(); }
        private static void RequirePassage(string passage, string evidence)
        {
            if (string.IsNullOrWhiteSpace(passage) || !Normalize(evidence).Contains(Normalize(passage)))
                throw new InvalidOperationException("SLIDE_ASSOCIATION_UNVERIFIED: Association must cite a passage from the slide's resolved source evidence.");
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
                    RequirePassage(passage, evidence);
                    foreach (var key in new[] { "label", "unit", "period" })
                        if (string.IsNullOrWhiteSpace(SamsungAuthoringPolicy.Text(operand, key)) ||
                            !Normalize(passage).Contains(Normalize(SamsungAuthoringPolicy.Text(operand, key))))
                            throw new InvalidOperationException("SLIDE_OPERAND_ASSOCIATION: Operand label, unit and period must occur in its cited passage.");
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
                        default: throw new InvalidOperationException("Unsupported calculation operation.");
                    }
                }
                catch (Exception ex) when (ex is OverflowException || ex is DivideByZeroException)
                { throw new InvalidOperationException("SLIDE_CALCULATION_INVALID: Overflow or zero baseline."); }
                var rounding = Number(calc, "decimals");
                if (rounding < 0 || rounding > 6 || rounding != decimal.Truncate(rounding)) throw new InvalidOperationException("Rounding must be 0 to 6 decimal places.");
                result = Math.Round(result, (int)rounding, MidpointRounding.AwayFromZero);
                if (result != Number(calc, "result")) throw new InvalidOperationException("SLIDE_CALCULATION_MISMATCH: Result differs from host arithmetic.");
                var unit = SamsungAuthoringPolicy.Text(calc, "unit");
                if (string.IsNullOrWhiteSpace(unit) || ((operation == "percent" || operation == "growth_percent") && unit != "%"))
                    throw new InvalidOperationException("SLIDE_CALCULATION_UNIT: Specify result units; percentage operations require %.");
                if ((operation == "ratio" || operation == "percent" || operation == "growth_percent") &&
                    operands.Select(o => SamsungAuthoringPolicy.Text(o, "unit")).Distinct().Count() != 1)
                    throw new InvalidOperationException("SLIDE_CALCULATION_UNIT: Ratio/growth operands must use the same units; convert units explicitly in the source first.");
                if (operation == "growth_percent" && values[1] <= 0)
                    throw new InvalidOperationException("SLIDE_CALCULATION_BASELINE: Growth percentages require a positive baseline; describe the absolute change otherwise.");
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
            foreach (var raw in claims)
            {
                var claim = SamsungAuthoringPolicy.ReadMap(raw);
                if (string.IsNullOrWhiteSpace(SamsungAuthoringPolicy.Text(claim, "text"))) throw new InvalidOperationException("Claim text is required.");
                RequirePassage(SamsungAuthoringPolicy.Text(claim, "evidence"), evidence);
                // Semantic association is checked separately by the source reviewer.
                foreach (var key in new[] { "label", "unit", "period" })
                    if (string.IsNullOrWhiteSpace(SamsungAuthoringPolicy.Text(claim, key)))
                        throw new InvalidOperationException("Claim associations need label, unit and period (use 'not applicable' for qualitative claims).");
            }
        }
    }
}

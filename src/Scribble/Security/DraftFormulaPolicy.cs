using System;
using System.Text.RegularExpressions;

namespace Scribble.Security
{
    // Safety gate for model-written Excel formulas in the Scribble
    // Draft sheet. A draft cell starting with '=' may become a live
    // formula only when it stays inside the workbook: functions that
    // reach the network or native code, and references to other
    // workbook files, are rejected and the cell is written as plain
    // text instead. The draft sheet itself is still never saved.
    public static class DraftFormulaPolicy
    {
        public const int MaxFormulaCharacters = 500;

        // Function names that can call out of the workbook: web
        // requests, realtime feeds, and legacy native-code bridges.
        private static readonly string[] BlockedFunctions =
        {
            "WEBSERVICE",
            "RTD",
            "CALL",
            "REGISTER",
            "EXEC",
            "HYPERLINK"
        };

        public static bool IsAllowedFormula(string formula)
        {
            if (formula == null ||
                formula.Length < 2 ||
                formula.Length > MaxFormulaCharacters ||
                formula[0] != '=')
            {
                return false;
            }

            // External workbook references use bracketed names.
            if (formula.IndexOf('[') >= 0 ||
                formula.IndexOf(']') >= 0)
            {
                return false;
            }

            foreach (var name in BlockedFunctions)
            {
                var match = Regex.Match(
                    formula,
                    "(^|[^A-Za-z0-9_.])" + name + "\\s*\\(",
                    RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    return false;
                }
            }

            return true;
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Web.Script.Serialization;

namespace Scribble.Office
{
    public static class PresentationRevisionAcceptance
    {
        public const string ChartlessScope = "chartless-v1";
        public static string ReceiptPath { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Scribble", "PowerPointAcceptance.json"); } }
        public static string AssemblyHash()
        {
            using (var stream = File.OpenRead(typeof(PresentationRevisionAcceptance).Assembly.Location))
            using (var hash = SHA256.Create()) return BitConverter.ToString(hash.ComputeHash(stream)).Replace("-", "").ToLowerInvariant();
        }
        public static bool Enabled
        {
            get
            {
                try
                {
                    if (!File.Exists(ReceiptPath)) return false;
                    var report = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(File.ReadAllText(ReceiptPath));
                    return Supports(report, AssemblyHash(), false);
                }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is ArgumentException) { return false; }
            }
        }
        public static bool SupportsCharts
        {
            get
            {
                try
                {
                    if (!File.Exists(ReceiptPath)) return false;
                    var report = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(File.ReadAllText(ReceiptPath));
                    return Supports(report, AssemblyHash(), true);
                }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is ArgumentException) { return false; }
            }
        }

        // Acceptance certifies a capability set, not every operation merely
        // because one native route passed. Old full receipts remain valid.
        public static bool Supports(Dictionary<string, object> report,
            string assemblyHash, bool requireCharts)
        {
            if (report == null || string.IsNullOrEmpty(assemblyHash) ||
                SamsungAuthoringPolicy.Text(report, "assembly_sha256") != assemblyHash ||
                SamsungAuthoringPolicy.Text(report, "policy") != SamsungAuthoringPolicy.Version ||
                SamsungAuthoringPolicy.Text(report, "execution_kind") != "native" ||
                !IsTrue(report, "preservation_passed") || !IsTrue(report, "rollback_passed") ||
                !IsTrue(report, "revision_passed") || IsTrue(report, "powerpoint_exited")) return false;
            var scope = SamsungAuthoringPolicy.Text(report, "scope");
            if (scope == ChartlessScope)
                return !requireCharts && IsTrue(report, "chartless_operations_passed") &&
                    !IsTrue(report, "all_operations_passed");
            return string.IsNullOrEmpty(scope) && IsTrue(report, "all_operations_passed");
        }

        internal static void RequireSupportedOperations(object presentation, object[] operations)
        {
            if (SupportsCharts) return;
            dynamic deck = presentation;
            for (var index = 1; index <= (int)deck.Slides.Count; index++)
                if (PresentationInspection.ContainsNativeChart((object)deck.Slides[index]))
                    throw new InvalidOperationException("REVISION_CHART_SCOPE_UNCERTIFIED: This build supports chartless revision only. Use a verified workbook-backed draft reconstruction for native charts.");
            if (ContainsChartOperation(operations))
                throw new InvalidOperationException("REVISION_CHART_SCOPE_UNCERTIFIED: Native chart mutation or insertion has not passed this build's acceptance.");
        }

        public static bool ContainsChartOperation(object value)
        {
            var map = value as System.Collections.IDictionary;
            if (map != null)
            {
                foreach (System.Collections.DictionaryEntry entry in map)
                {
                    var key = Convert.ToString(entry.Key);
                    if (((key == "chart" || key == "secondary_chart") && entry.Value != null) ||
                        (key == "kind" && Convert.ToString(entry.Value).StartsWith("chart", StringComparison.Ordinal))) return true;
                    if (ContainsChartOperation(entry.Value)) return true;
                }
                return false;
            }
            var sequence = value as System.Collections.IEnumerable;
            if (sequence != null && !(value is string))
                foreach (var item in sequence) if (ContainsChartOperation(item)) return true;
            return false;
        }
        private static bool IsTrue(Dictionary<string, object> report, string key)
        { object value; return report.TryGetValue(key, out value) && value is bool && (bool)value; }
    }
}

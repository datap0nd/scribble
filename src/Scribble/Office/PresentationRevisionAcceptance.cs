using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Web.Script.Serialization;

namespace Scribble.Office
{
    public static class PresentationRevisionAcceptance
    {
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
                    return SamsungAuthoringPolicy.Text(report, "assembly_sha256") == AssemblyHash() &&
                        SamsungAuthoringPolicy.Text(report, "policy") == SamsungAuthoringPolicy.Version &&
                        SamsungAuthoringPolicy.Text(report, "execution_kind") == "native" &&
                        IsTrue(report, "all_operations_passed") && IsTrue(report, "preservation_passed") && IsTrue(report, "rollback_passed") && IsTrue(report, "revision_passed");
                }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is ArgumentException) { return false; }
            }
        }
        private static bool IsTrue(Dictionary<string, object> report, string key)
        { object value; return report.TryGetValue(key, out value) && value is bool && (bool)value; }
    }
}

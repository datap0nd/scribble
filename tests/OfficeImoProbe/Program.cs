using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using OfficeIMO.PowerPoint;

namespace OfficeImoProbe
{
    // Standalone compatibility and preservation spike. Not part of Scribble.sln.
    internal static class Program
    {
        private static int Main(string[] args)
        {
            if (args.Length != 2) return 2;
            var input = Path.GetFullPath(args[0]);
            var output = Path.GetFullPath(args[1]);
            Directory.CreateDirectory(output);
            var noOp = Path.Combine(output, "officeimo-noop.pptx");
            var edited = Path.Combine(output, "officeimo-edited.pptx");
            using (var deck = PowerPointPresentation.Load(input))
            {
                var features = deck.InspectFeatures();
                File.WriteAllText(Path.Combine(output, "feature-report.txt"),
                    string.Join(Environment.NewLine, FeatureLines(features)));
                deck.SaveCopy(noOp);
            }
            using (var deck = PowerPointPresentation.Load(input))
            {
                deck.ReplaceText("Two period comparison", "Period comparison");
                deck.SaveCopy(edited);
            }
            var before = Parts(input);
            var afterNoOp = Parts(noOp);
            var afterEdit = Parts(edited);
            var summary = new List<string>
            {
                "OfficeIMO.PowerPoint=3.4.2",
                "TargetFramework=net48",
                "InputSha256=" + Hash(File.ReadAllBytes(input)),
                "NoOpSha256=" + Hash(File.ReadAllBytes(noOp)),
                "EditSha256=" + Hash(File.ReadAllBytes(edited)),
                "InputParts=" + before.Count,
                "NoOpParts=" + afterNoOp.Count,
                "EditParts=" + afterEdit.Count,
                "NoOpMissingParts=" + string.Join(",", before.Keys.Except(afterNoOp.Keys).OrderBy(x => x)),
                "EditMissingParts=" + string.Join(",", before.Keys.Except(afterEdit.Keys).OrderBy(x => x)),
                "NoOpChangedParts=" + string.Join(",", before.Keys.Where(x =>
                    afterNoOp.ContainsKey(x) && before[x] != afterNoOp[x]).OrderBy(x => x)),
                "EditChangedParts=" + string.Join(",", before.Keys.Where(x =>
                    afterEdit.ContainsKey(x) && before[x] != afterEdit[x]).OrderBy(x => x))
            };
            File.WriteAllLines(Path.Combine(output, "probe-report.txt"), summary);
            Console.WriteLine(string.Join(Environment.NewLine, summary));
            // This is an observation. A missing part needs manual diagnosis;
            // do not certify OfficeIMO from a successful process exit.
            return 0;
        }

        private static Dictionary<string, string> Parts(string path)
        {
            using (var archive = ZipFile.OpenRead(path))
                return archive.Entries.ToDictionary(entry => entry.FullName,
                    entry =>
                    {
                        using (var stream = entry.Open())
                        using (var memory = new MemoryStream())
                        { stream.CopyTo(memory); return Hash(memory.ToArray()); }
                    }, StringComparer.Ordinal);
        }

        private static IEnumerable<string> FeatureLines(object report)
        {
            foreach (var property in report.GetType().GetProperties())
            {
                var value = property.GetValue(report, null);
                var items = value as System.Collections.IEnumerable;
                if (items == null || value is string)
                {
                    yield return property.Name + "=" + Convert.ToString(value);
                    continue;
                }
                foreach (var item in items)
                {
                    if (item == null) continue;
                    var type = item.GetType();
                    var name = type.GetProperty("Name");
                    var count = type.GetProperty("Count");
                    yield return property.Name + ":" +
                        (name == null ? type.Name : Convert.ToString(name.GetValue(item, null))) +
                        ":" + (count == null ? "" : Convert.ToString(count.GetValue(item, null)));
                }
            }
        }

        private static string Hash(byte[] bytes)
        {
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "")
                    .ToLowerInvariant();
        }
    }
}

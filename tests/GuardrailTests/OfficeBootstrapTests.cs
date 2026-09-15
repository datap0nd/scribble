using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Xml.Linq;
using Scribble.Chat;
using Scribble.Office;
using Scribble.Security;
using Scribble.Testing;

namespace GuardrailTests
{
    internal static class OfficeBootstrapTests
    {
        private static void Check(bool value, string message)
        { if (!value) throw new InvalidOperationException(message); }

        private static void RejectSibling(string expected)
        {
            try { TestLabOfficeConnection.ResolvePreparedSibling("Excel.Application"); }
            catch (Exception error) when (error is InvalidOperationException || error is InvalidDataException)
            {
                Check(error.Message.Contains(expected), "Attachment reached the wrong boundary: " + error.Message);
                // The production draft host must use the same rejecting path,
                // including an idle/finished lab that still owns its session.
                using (var draft = new DocumentDraftHost("powerpoint", new object()))
                {
                    var call = new ChatToolCall { id = "boundary", type = "function", function = new ChatToolCallFunction {
                        name = CrossAppToolCatalog.SendToExcel, arguments = "{\"title\":\"Boundary probe\",\"rows\":[[\"Value\"],[1]]}"
                    } };
                    var result = draft.Execute(call, new OneShotDraftAuthorization(true), true, "Create a synthetic Excel draft.");
                    Check(result.Content.Contains("DRAFT_CREATION_FAILED") && result.Content.Contains(expected),
                        "The draft host bypassed the prepared destination boundary: " + result.Content);
                }
                return;
            }
            throw new InvalidOperationException("An unverified Office destination was accepted.");
        }

        public static void UnverifiedSiblingCannotStartOffice()
        {
            // Exercise the real public resolver, stopping before any Office call.
            // Preserve expired descriptors and refuse to interfere with a live lab.
            Directory.CreateDirectory(TestLab.Root);
            using (var lease = new FileStream(Path.Combine(TestLab.Root, "suite.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))
            {
                Check(TestLab.Status() == null && TestLabSuite.Active() == null, "Close the active Test Lab before running the attachment regression.");
                var descriptors = new[] { Path.Combine(TestLab.Root, "session.bin"), Path.Combine(TestLab.Root, "suite.bin") };
                var backups = descriptors.Select(p => File.Exists(p) ? File.ReadAllBytes(p) : null).ToArray();
                var id = Guid.NewGuid().ToString("N");
                var root = Path.Combine(Path.GetTempPath(), "scribble-bootstrap-boundary-" + id);
                var kit = Path.Combine(root, "kit");
                var native = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Scribble Testcases", "Native", id);
                string runDirectory = null;
                var enabled = false;
                try
                {
                    Directory.CreateDirectory(Path.Combine(kit, "operator"));
                    File.WriteAllText(Path.Combine(kit, "operator", "cases.json"), TestLab.Serialize(new[] {
                        new LabCase { id = "XA01", host = "PowerPoint", prompt = "Attachment boundary only; never submitted.", inputs = new string[0] }
                    }));
                    File.WriteAllText(Path.Combine(kit, "operator", "mail-index.json"), "[]");
                    var files = Directory.GetFiles(kit, "*", SearchOption.AllDirectories).Select(p => new KitFile {
                        path = p.Substring(kit.Length + 1).Replace('\\', '/'), size = new FileInfo(p).Length, sha256 = TestLab.FileHash(p)
                    }).ToArray();
                    File.WriteAllText(Path.Combine(kit, "manifest.json"), TestLab.Serialize(new KitManifest { schema = 1, suite_id = "atlas-v1", files = files }));
                    using (var process = Process.GetCurrentProcess())
                    {
                        var state = new SuiteState { id = id, folder = root, fixtureSuiteId = "atlas-v1", caseId = "XA01", host = "PowerPoint",
                            pid = process.Id, processStart = process.StartTime.ToUniversalTime().Ticks, expires = DateTime.UtcNow.AddMinutes(2) };
                        TestLabSuite.Save(state); TestLab.Enable(kit); enabled = true;
                        RejectSibling("active suite case");
                        var run = TestLab.Start("XA01", "PowerPoint", true); runDirectory = TestLab.RunDirectory(run.run_id);
                        state.runId = run.run_id; TestLabSuite.Save(state);
                        RejectSibling("no prepared host receipt");
                        var office = Path.Combine(native, "office"); Directory.CreateDirectory(office);
                        var startup = Path.Combine(office, "startup-excel-" + Guid.NewGuid().ToString("N") + ".xlsx");
                        using (var source = typeof(TestLab).Assembly.GetManifestResourceStream("Scribble.Testing.Bootstrap.Excel.xlsx"))
                        using (var target = File.Create(startup)) source.CopyTo(target);
                        Action<string, int> receipt = (suiteId, pid) => File.WriteAllBytes(Path.Combine(office, "excel.binding.bin"),
                            ProtectedData.Protect(Encoding.UTF8.GetBytes(TestLab.Serialize(new {
                                schema = 1, suite_id = suiteId, host = "Excel", pid = pid, process_start = 1L,
                                executable = Path.Combine(root, "EXCEL.EXE"), startup_path = startup, startup_sha256 = TestLab.FileHash(startup)
                            })), null, DataProtectionScope.CurrentUser));
                        receipt(Guid.NewGuid().ToString("N"), process.Id);
                        RejectSibling("verified suite/startup boundary");
                        receipt(id, int.MaxValue);
                        RejectSibling("prepared process identity");
                        TestLab.Finish(true);
                        RejectSibling("active suite case");
                    }
                }
                finally
                {
                    if (enabled) TestLab.Disable();
                    for (var i = 0; i < descriptors.Length; i++)
                        if (backups[i] == null) { if (File.Exists(descriptors[i])) File.Delete(descriptors[i]); }
                        else File.WriteAllBytes(descriptors[i], backups[i]);
                    // These paths contain this test's freshly generated GUID only.
                    if (runDirectory != null && Directory.Exists(runDirectory)) Directory.Delete(runDirectory, true);
                    if (Directory.Exists(native)) Directory.Delete(native, true);
                    if (Directory.Exists(root)) Directory.Delete(root, true);
                }
            }
        }

        public static void CancelledStartupDoesNotReachOffice()
        {
            var logged = false;
            using (var cancel = new CancellationTokenSource())
            {
                cancel.Cancel();
                try
                {
                    TestLabOfficeConnection.ConnectAsync("Excel", cancel.Token, message => logged = true).GetAwaiter().GetResult();
                }
                catch (OperationCanceledException)
                {
                    Check(!logged, "A cancelled cold start performed Office setup before checking Stop.");
                    return;
                }
            }
            throw new InvalidOperationException("A cancelled cold start was not cancelled.");
        }

        public static void NeutralEmbeddedDocuments()
        {
            var assembly = typeof(TestLab).Assembly;
            foreach (var host in new[] { "Excel.xlsx", "PowerPoint.pptx" })
            using (var resource = assembly.GetManifestResourceStream("Scribble.Testing.Bootstrap." + host))
            {
                Check(resource != null, "The cold-start document is not embedded for " + host + ".");
                using (var package = new ZipArchive(resource, ZipArchiveMode.Read))
                {
                    Check(!package.Entries.Any(e => e.FullName.IndexOf("vba", StringComparison.OrdinalIgnoreCase) >= 0),
                        "A bootstrap document contains VBA.");
                    foreach (var entry in package.Entries.Where(e => e.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)))
                    using (var stream = entry.Open())
                    {
                        var relationships = XDocument.Load(stream);
                        Check(!relationships.Descendants().Any(e => string.Equals((string)e.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase)),
                            "A bootstrap document links to an external resource.");
                    }
                    if (host.StartsWith("Excel", StringComparison.Ordinal))
                    {
                        var sheets = package.Entries.Where(e => e.FullName.StartsWith("xl/worksheets/", StringComparison.Ordinal) &&
                            e.FullName.EndsWith(".xml", StringComparison.Ordinal)).ToArray();
                        Check(sheets.Length == 1, "Excel startup must provide exactly one neutral worksheet.");
                        using (var stream = sheets[0].Open())
                        {
                            var sheet = XDocument.Load(stream);
                            Check(!sheet.Descendants().Any(e => e.Name.LocalName == "v" || e.Name.LocalName == "f" || e.Name.LocalName == "is"),
                                "The Excel bootstrap includes task data or formulas.");
                        }
                    }
                    else
                    {
                        var slides = package.Entries.Where(e => e.FullName.StartsWith("ppt/slides/slide", StringComparison.Ordinal) &&
                            e.FullName.EndsWith(".xml", StringComparison.Ordinal)).ToArray();
                        Check(slides.Length == 1, "PowerPoint startup must provide one blank native document window.");
                        using (var stream = slides[0].Open())
                        {
                            var slide = XDocument.Load(stream);
                            Check(!slide.Descendants().Any(e => new[] { "sp", "pic", "graphicFrame", "cxnSp", "t" }.Contains(e.Name.LocalName)),
                                "The PowerPoint bootstrap includes drawable task content.");
                        }
                    }
                }
            }
        }
    }
}

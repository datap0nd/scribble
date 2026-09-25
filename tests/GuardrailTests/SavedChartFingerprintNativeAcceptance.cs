using System;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using System.Threading;
using System.Web.Script.Serialization;
using Scribble.Office;

namespace GuardrailTests
{
    // Explicit opt-in. Creates, saves, and deletes only its own disposable deck.
    internal static class SavedChartFingerprintNativeAcceptance
    {
        private static void Check(bool condition, string message)
        { if (!condition) throw new InvalidOperationException(message); }

        internal static int Run(string reportPath)
        {
            dynamic app = null, deck = null;
            var path = Path.Combine(Path.GetDirectoryName(
                Path.GetFullPath(reportPath)),
                "scribble-native-saved-fingerprint-" +
                Guid.NewGuid().ToString("N") + ".pptx");
            var stage = "powerpoint_start";
            var failure = string.Empty;
            var copyEvents = 0;
            var savedDeckPreserved = false;
            var chartEditDetected = false;
            var packageRejected = false;
            try
            {
                app = Activator.CreateInstance(Type.GetTypeFromProgID(
                    "PowerPoint.Application", true));
                app.Visible = -1;
                deck = app.Presentations.Add(-1);
                dynamic slide = deck.Slides.Add(1, 12);
                stage = "create_chart";
                dynamic chartShape = slide.Shapes.AddChart2(201, 51,
                    100, 100, 480, 300);
                dynamic chart = chartShape.Chart;
                if ((int)chart.SeriesCollection().Count == 0)
                    chart.SeriesCollection().NewSeries();
                dynamic series = chart.SeriesCollection(1);
                series.Name = "Boundary source";
                stage = "save_owned_fixture";
                deck.Tags.Add("ScribbleTask", "native-fingerprint-boundary");
                slide.Tags.Add("ScribbleTask", "native-fingerprint-boundary");
                deck.SaveAs(path, 24);
                Check(!string.IsNullOrEmpty(Convert.ToString(deck.Path)),
                    "The native fixture was not saved.");
                var sourceHash = Hash(File.ReadAllBytes(path));
                stage = "saved_chart_fingerprint";
                using (var watcher = new FileSystemWatcher(Path.GetTempPath(),
                    "scribble-chart-fingerprint-*.pptx"))
                {
                    watcher.Created += (sender, args) =>
                        Interlocked.Increment(ref copyEvents);
                    watcher.EnableRaisingEvents = true;
                    var before = PresentationInspection.Fingerprint(
                        (object)slide);
                    series.Name = "Boundary changed";
                    var after = PresentationInspection.Fingerprint(
                        (object)slide);
                    chartEditDetected = before != after;
                    var method = typeof(PresentationInspection).GetMethod(
                        "PackageSlideFingerprint", BindingFlags.NonPublic |
                        BindingFlags.Static);
                    Check(method != null, "Package boundary is missing.");
                    try { method.Invoke(null, new object[] { (object)slide }); }
                    catch (TargetInvocationException error)
                    {
                        packageRejected = error.InnerException is
                            InvalidOperationException &&
                            error.InnerException.Message.Contains(
                                "CHART_PACKAGE_UNSAVED_DRAFT_REQUIRED");
                    }
                    Thread.Sleep(300);
                }
                savedDeckPreserved = sourceHash == Hash(File.ReadAllBytes(path));
                Check(copyEvents == 0,
                    "A saved presentation triggered SaveCopyAs.");
                Check(chartEditDetected,
                    "A saved chart edit escaped the fingerprint.");
                Check(packageRejected,
                    "A saved deck passed the package ownership guard.");
                Check(savedDeckPreserved,
                    "Fingerprint modified the saved presentation.");
            }
            catch (Exception error) { failure = stage + ": " + error; }
            finally
            {
                if ((object)deck != null)
                    try { deck.Close(); } catch { }
                if (File.Exists(path))
                    try { File.Delete(path); } catch { }
            }
            var report = new {
                execution_kind = "native_disposable_saved_chart_fingerprint",
                copy_events = copyEvents,
                chart_edit_detected = chartEditDetected,
                package_rejected = packageRejected,
                saved_deck_preserved = savedDeckPreserved,
                passed = failure.Length == 0,
                failure
            };
            var json = new JavaScriptSerializer().Serialize(report);
            File.WriteAllText(reportPath, json);
            Console.WriteLine(json);
            return failure.Length == 0 ? 0 : 1;
        }

        private static string Hash(byte[] value)
        {
            using (var hash = SHA256.Create())
                return BitConverter.ToString(hash.ComputeHash(value));
        }
    }
}

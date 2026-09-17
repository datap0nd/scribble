using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Scribble.Outlook;

namespace Scribble.Testing
{
    // Operator-only preparation, on the runner's pumped STA. Keep every COM
    // application alive until the suite ends, including cross-app destinations.
    internal sealed class TestLabOfficeEnvironment : IDisposable
    {
        [ThreadStatic] private static TestLabOfficeEnvironment current;
        private readonly TestLabOfficeEnvironment previous;
        private readonly Dictionary<string, object> applications = new Dictionary<string, object>();
        private readonly Action<string> log;
        private readonly TestLabComMessageFilter messageFilter;
        private string suiteId;
        internal TestLabOfficeEnvironment(Action<string> log) : this(log, CancellationToken.None) { }
        internal TestLabOfficeEnvironment(Action<string> log, CancellationToken cancel)
        { this.log = log; messageFilter = new TestLabComMessageFilter(cancel); previous = current; current = this; }

        // Capture/cleanup share the same retained applications as preparation.
        // Outlook can be fully usable without publishing a ROT entry. A
        // borrowed RCW must not be released by the caller that did not acquire it.
        internal static object Borrow(string host, out bool release)
        {
            object value;
            if (current != null && current.applications.TryGetValue(host, out value)) { release = false; return value; }
            release = true;
            return Marshal.GetActiveObject(host + ".Application");
        }

        internal async Task<object> ConnectAsync(string host, CancellationToken cancel)
        {
            cancel.ThrowIfCancellationRequested();
            object value;
            if (applications.TryGetValue(host, out value))
            {
                try { var version = Convert.ToString(((dynamic)value).Version); return value; }
                catch (COMException error) when (Disconnected(error))
                {
                    // Closing the last presentation or an Office crash can
                    // invalidate the retained automation object between cases.
                    applications.Remove(host); Release(value);
                    log(host + ": previous connection ended; reconnecting for this case.");
                }
            }
            if (host == "Excel" || host == "PowerPoint")
            {
                value = await TestLabOfficeConnection.ConnectAsync(host, cancel, log);
                suiteId = TestLabSuite.Active()?.id;
            }
            else
            {
                try { value = Marshal.GetActiveObject(host + ".Application"); }
                catch (COMException)
                {
                    value = null;
                    if (host == "Outlook")
                    {
                        // Classic Outlook's singleton automation class attaches
                        // to its normally launched explorer even without ROT.
                        // Give that normal UI launch time to win the singleton
                        // race before Activator can create an -Embedding server
                        // whose COMAddIns collection is still incomplete.
                        using (var launched = System.Diagnostics.Process.Start(
                            new System.Diagnostics.ProcessStartInfo("outlook.exe") { UseShellExecute = true }))
                            value = await WaitForNormallyLaunchedOutlookAsync(launched, cancel, log);
                    }
                    if (value == null)
                    {
                        var type = Type.GetTypeFromProgID(host + ".Application");
                        if (type == null) throw new InvalidOperationException(host + " is not installed or its automation registration is unavailable.");
                        value = Activator.CreateInstance(type);
                    }
                }
            }
            try
            {
                cancel.ThrowIfCancellationRequested();
                dynamic app = value;
                if (host == "PowerPoint") app.Visible = -1;
                else if (host != "Outlook") app.Visible = true;
                applications.Add(host, value);
                suiteId = TestLabSuite.Active()?.id;
                log(host + ": connected and retained for the suite lifetime.");
                return value;
            }
            catch { Release(value); throw; }
        }

        internal async Task Prepare(LabCase c, CancellationToken cancel)
        {
            var session = TestLab.Status();
            if (session == null || session.run_id != null) throw new InvalidOperationException("Preparation requires an idle fixture session.");
            TestLab.VerifyKit(session.fixture_root);
            var destinations = new Dictionary<string, string> { { "xlsx", "Excel" }, { "pptx", "PowerPoint" }, { "msg", "Outlook" }, { "docx", "Word" } };
            foreach (var host in new[] { c.host }.Concat((c.artifacts ?? new string[0]).Where(destinations.ContainsKey).Select(a => destinations[a])).Distinct())
            {
                cancel.ThrowIfCancellationRequested();
                await ConnectAsync(host, cancel);
                await Task.Delay(100, cancel);
            }
            dynamic origin = await ConnectAsync(c.host, cancel);
            var progId = c.host == "Outlook" ? "Scribble.AddIn" : "Scribble." + c.host + "AddIn";
            if (c.host == "Outlook")
            {
                // Activator can initially return Outlook's -Embedding server.
                // Display an Explorer before asking for COMAddIns; otherwise a
                // newly started Outlook can expose an incomplete collection and
                // Item(progId) fails with DISP_E_BADINDEX.
                DisplayOutlookExplorer((object)origin);
            }
            object addinObject = await WaitForAddInAsync((object)origin, progId, c.host, cancel);
            try
            {
                dynamic addin = addinObject;
                if (!Convert.ToBoolean(addin.Connect)) addin.Connect = true;
                if (!Convert.ToBoolean(addin.Connect)) throw new InvalidOperationException("Scribble is disabled in " + c.host + ". Check that app's Disabled Items or organizational policy.");
            }
            finally { Release(addinObject); }
            var extension = c.host == "Excel" ? ".xlsx" : c.host == "PowerPoint" ? ".pptx" : c.host == "Word" ? ".docx" : null;
            foreach (var relative in (c.inputs ?? new string[0]).Where(p => extension != null && p.EndsWith(extension, StringComparison.OrdinalIgnoreCase)))
            {
                cancel.ThrowIfCancellationRequested();
                var path = TestLab.SafeChild(session.fixture_root, relative);
                // Excel refuses two workbooks with the same basename, even in
                // different folders. Always use a unique, byte-identical alias.
                if (c.host == "Excel")
                {
                    var alias = TestLab.SafeChild(session.fixture_root, ".aliases/" + Path.GetFileNameWithoutExtension(path) + "--" + session.session_id + extension);
                    Directory.CreateDirectory(Path.GetDirectoryName(alias));
                    File.Copy(path, alias, false);
                    if (TestLab.FileHash(alias) != TestLab.FileHash(path)) throw new InvalidDataException("Fixture alias verification failed.");
                    path = alias;
                }
                object document = null;
                try
                {
                    if (c.host == "Excel") document = origin.Workbooks.Open(path, 0, true);
                    else if (c.host == "PowerPoint") document = origin.Presentations.Open(path, -1, 0, -1);
                    else document = origin.Documents.Open(path, false, true);
                    if (document == null) throw new InvalidOperationException(c.host + " did not open " + relative + ". Check its visible startup or file dialog.");
                    dynamic native = document;
                    if (c.host == "PowerPoint") native.Windows.Item(1).Activate(); else native.Activate();
                    log("Opened read-only " + relative + " in " + c.host + ".");
                }
                finally { Release(document); }
                await Task.Delay(100, cancel);
            }
            if (c.host == "Outlook")
            {
                if (TestLabMailbox.Enabled)
                {
                    TestLabMailbox.Prepare((object)origin, cancel, log);
                    return;
                }
                TestLabMail.Prepare((object)origin, c, cancel, log);
                // The add-in task pane belongs to the explorer. No mailbox
                // import or PST registration is needed for local MSG fixtures.
                DisplayOutlookExplorer((object)origin);
            }
        }

        private async Task<object> WaitForAddInAsync(object application, string progId, string host, CancellationToken cancel)
        {
            var started = DateTime.UtcNow;
            var announced = false;
            while (DateTime.UtcNow - started < TimeSpan.FromSeconds(30))
            {
                cancel.ThrowIfCancellationRequested();
                try { return ((dynamic)application).COMAddIns.Item(progId); }
                catch (COMException error) when ((uint)error.HResult == 0x8002000B)
                {
                    if (!announced)
                    {
                        log(host + ": waiting for the Scribble COM add-in to appear after Office startup.");
                        announced = true;
                    }
                    await Task.Delay(250, cancel);
                }
            }
            throw new InvalidOperationException("Scribble is not present in " + host + "'s COM add-in collection after startup. Repair the selected Scribble component or check Office Disabled Items.");
        }

        private static async Task<object> WaitForNormallyLaunchedOutlookAsync(
            System.Diagnostics.Process launched, CancellationToken cancel, Action<string> log)
        {
            var deadline = DateTime.UtcNow.AddSeconds(5);
            var announced = false;
            while (DateTime.UtcNow < deadline)
            {
                cancel.ThrowIfCancellationRequested();
                try { return Marshal.GetActiveObject("Outlook.Application"); }
                catch (COMException)
                {
                    if (!announced)
                    {
                        (log ?? delegate { })("Outlook: waiting for the normally launched Explorer before automation attachment.");
                        announced = true;
                    }
                }
                if (launched != null)
                {
                    try { if (launched.HasExited) break; }
                    catch (InvalidOperationException) { break; }
                }
                await Task.Delay(250, cancel);
            }
            return null;
        }

        private static void DisplayOutlookExplorer(object application)
        {
            dynamic outlook = application;
            object explorer = outlook.ActiveExplorer();
            if (explorer == null) explorer = outlook.Session.GetDefaultFolder(6).GetExplorer();
            try { ((dynamic)explorer).Display(); } finally { Release(explorer); }
        }

        internal static void Release(object value) { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
        private static bool Disconnected(COMException error)
        {
            var code = (uint)error.HResult;
            return code == 0x800706BE || code == 0x800706BA || code == 0x80010108 || code == 0x800401FD;
        }
        public void Dispose()
        {
            try
            {
                foreach (var pair in applications)
                {
                    if (suiteId != null && (pair.Key == "Excel" || pair.Key == "PowerPoint"))
                        TestLabOfficeConnection.Cleanup(pair.Key, suiteId, log);
                    else if (suiteId != null && pair.Key == "Outlook" && TestLabMailbox.Enabled)
                    {
                        try { TestLabMailbox.Cleanup(pair.Value, suiteId, log); }
                        catch (Exception error) { log("Outlook: could not detach the suite PST during cleanup; its file was preserved. " + error.Message); }
                    }
                    Release(pair.Value);
                }
                applications.Clear();
            }
            finally { if (current == this) current = previous; messageFilter.Dispose(); }
        }
    }

    // Local native MSG inputs exercise MessageReader and attachment parsers
    // without requiring permission to add a PST to a corporate Outlook profile.
    // Fixture IDs resolve only within the current, verified, synthetic run.
    public static class TestLabMail
    {
        private const string Prefix = "scribble-fixture:";
        private static MailFixture[] Sources(string root, string[] inputs)
        {
            return TestLabSuite.Read<MailFixture[]>(TestLab.SafeChild(root, "operator/mail-index.json"))
                .Where(m => (inputs ?? new string[0]).Contains(m.path)).ToArray();
        }
        private static string NativePath(string root, MailFixture source)
        { return Path.Combine(TestLab.NativeDirectory(TestLab.Hash(System.Text.Encoding.UTF8.GetBytes(root)).Substring(0, 32), "mail"), Path.GetFileNameWithoutExtension(source.path) + ".msg"); }

        internal static void Prepare(object application, LabCase c, CancellationToken cancel, Action<string> log)
        {
            var root = TestLab.Status().fixture_root;
            dynamic app = application;
            foreach (var source in Sources(root, c.inputs))
            {
                cancel.ThrowIfCancellationRequested();
                var path = NativePath(root, source);
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                object value = app.CreateItem(0);
                try
                {
                    dynamic mail = value;
                    mail.Subject = source.subject; mail.Body = source.body; mail.To = "review@example.test";
                    dynamic properties = mail.PropertyAccessor;
                    properties.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0C1A001F", source.sender);
                    properties.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0C1F001F", source.sender);
                    properties.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0C1E001F", "SMTP");
                    var date = DateTime.Parse(source.date, System.Globalization.CultureInfo.InvariantCulture).ToUniversalTime();
                    properties.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x00390040", date);
                    properties.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0E060040", date);
                    foreach (var attachment in source.attachments ?? new string[0])
                    {
                        object attached = mail.Attachments.Add(TestLab.SafeChild(root, attachment));
                        TestLabOfficeEnvironment.Release(attached);
                    }
                    properties.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0E070003", 1);
                    mail.SaveAs(path, 9);
                    mail.Close(1);
                    log("Prepared native local message: " + source.subject);
                }
                finally { TestLabOfficeEnvironment.Release(value); }
            }
        }

        public static object OpenItem(object session, string entryId, string storeId = null)
        {
            dynamic ns = session;
            if (!(entryId ?? "").StartsWith(Prefix, StringComparison.Ordinal))
            {
                TestLabMailbox.ValidateIdentity(entryId, storeId);
                return string.IsNullOrEmpty(storeId) ? ns.GetItemFromID(entryId) : ns.GetItemFromID(entryId, storeId);
            }
            var parts = entryId.Split(':');
            int index;
            var runId = TestLab.ActiveRunId();
            if (parts.Length != 3 || parts[1] != runId || runId == null || !int.TryParse(parts[2], out index))
                throw new InvalidOperationException("The synthetic email belongs to an inactive test run.");
            var run = TestLab.GetRun(runId);
            var sources = Sources(run.fixture_root, run.input_paths);
            if (index < 0 || index >= sources.Length) throw new InvalidDataException("Unknown synthetic email.");
            object item = ns.OpenSharedItem(NativePath(run.fixture_root, sources[index]));
            try { TestLab.CheckMailSource(item); return item; }
            catch { TestLabOfficeEnvironment.Release(item); throw; }
        }

        public static IReadOnlyList<MessageSnapshot> Load(object application, LabCase c)
        {
            if (TestLabMailbox.Enabled) return TestLabMailbox.Load(application, c);
            var runId = TestLab.ActiveRunId();
            if (runId == null) throw new InvalidOperationException("No active synthetic run.");
            var run = TestLab.GetRun(runId);
            if (run.case_id != c.id) throw new InvalidOperationException("Synthetic case mismatch.");
            var sources = Sources(run.fixture_root, run.input_paths);
            var reader = new MessageReader(application);
            var messages = new List<MessageSnapshot>();
            for (int i = 0; i < sources.Length; i++) messages.Add(reader.CaptureById(Prefix + runId + ":" + i, ""));
            return messages;
        }
    }
}

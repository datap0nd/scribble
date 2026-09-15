using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Microsoft.Win32;

namespace Scribble.Testing
{
    // Operator-only bootstrap. These APIs are not COM-visible or model tools.
    // Office automation startup can omit per-user add-ins even when registration
    // is correct. Launch a private document normally, then bind its exact host.
    [ComVisible(false)]
    public static class TestLabOfficeConnection
    {
        private const int StartupSeconds = 45;
        private static readonly Guid Dispatch = new Guid("00020400-0000-0000-C000-000000000046");

        public static async Task<object> ConnectAsync(string host, CancellationToken cancel, Action<string> log)
        {
            cancel.ThrowIfCancellationRequested();
            ValidateHost(host);
            var state = RequireScope(false);
            RequireSta();
            log = log ?? delegate { };
            var directory = DirectoryFor(state);
            var descriptor = Descriptor(directory, host);
            var attemptedConnect = new HashSet<int>();
            if (File.Exists(descriptor))
            {
                var previous = ReadBinding(state, host, false);
                if (ProcessMatches(previous))
                    return await WaitForApplication(previous, state.id, cancel, log, attemptedConnect);
                log(host + ": the previously prepared process ended; retaining its startup file and opening a fresh host.");
            }

            var executable = FindExecutable(host);
            var before = Processes(host);
            if (host == "PowerPoint" && before.Count > 1)
                throw new InvalidOperationException("Multiple PowerPoint processes are open. The Test Lab cannot identify a single safe destination; existing documents were preserved.");
            // PowerPoint normally reuses its existing process. Do not repair or
            // replace an automation-only instance by terminating user work.
            if (host == "PowerPoint" && before.Count == 1)
            {
                var existing = before[0];
                object app = null;
                try
                {
                    app = TryAttach(host, existing.pid, existing.process_start, executable, null, cancel);
                    if (app != null && !HasScribble(app, host, existing.pid, attemptedConnect))
                        throw new InvalidOperationException("The existing PowerPoint instance does not expose Scribble. It was preserved. Open PowerPoint normally with Scribble available, then retry the suite.");
                }
                finally { Release(app); }
            }

            cancel.ThrowIfCancellationRequested();
            var startup = TestLab.SafeChild(directory, "startup-" + host.ToLowerInvariant() + "-" + Guid.NewGuid().ToString("N") + Extension(host));
            var bytes = StartupBytes(host);
            using (var file = new FileStream(startup, FileMode.CreateNew, FileAccess.Write, FileShare.Read)) file.Write(bytes, 0, bytes.Length);
            File.SetAttributes(startup, File.GetAttributes(startup) | FileAttributes.ReadOnly);
            var binding = new HostBinding { schema = 1, suite_id = state.id, host = host, executable = executable,
                startup_path = startup, startup_sha256 = TestLab.Hash(bytes) };
            cancel.ThrowIfCancellationRequested();
            var launchedAt = DateTime.UtcNow;
            using (var launched = Process.Start(new ProcessStartInfo(executable)
            {
                UseShellExecute = true,
                Arguments = (host == "Excel" ? "/x " : "") + "\"" + startup + "\""
            }))
            {
                log(host + ": opening a private startup document with normal Office startup.");
                var deadline = DateTime.UtcNow.AddSeconds(StartupSeconds);
                Exception lastError = null;
                while (DateTime.UtcNow < deadline)
                {
                    cancel.ThrowIfCancellationRequested();
                    RequireSameSuite(state.id, false);
                    var candidates = Processes(host).Where(p => IsLaunchCandidate(host, before.Count,
                        before.Any(old => old.pid == p.pid && old.process_start == p.process_start),
                        p.process_start, launchedAt.Ticks)).ToList();
                    foreach (var candidate in candidates)
                    {
                        object app = null;
                        try
                        {
                            cancel.ThrowIfCancellationRequested();
                            if (!SamePath(candidate.executable, executable)) continue;
                            app = TryAttach(host, candidate.pid, candidate.process_start, executable, startup, cancel);
                            if (app == null) continue;
                            if (binding.pid != candidate.pid || binding.process_start != candidate.process_start)
                            {
                                // Retain a verified native host even when its
                                // add-in is slow or disabled. The next case must
                                // retry this process instead of spawning another
                                // Excel instance after each add-in timeout.
                                binding.pid = candidate.pid; binding.process_start = candidate.process_start;
                                RequireSameSuite(state.id, false);
                                cancel.ThrowIfCancellationRequested();
                                WriteBinding(descriptor, binding);
                            }
                            if (!HasScribble(app, host, candidate.pid, attemptedConnect))
                            { lastError = new InvalidOperationException("Scribble has not appeared in the interactive add-in collection."); continue; }
                            RequireSameSuite(state.id, false);
                            cancel.ThrowIfCancellationRequested();
                            log(host + ": verified interactive process " + binding.pid + " and Scribble; retained private startup document.");
                            var result = app; app = null; return result;
                        }
                        catch (OperationCanceledException) { throw; }
                        catch (COMException error) { lastError = error; }
                        finally { Release(app); }
                    }
                    await Task.Delay(250, cancel);
                }
                throw StartupFailure(host, lastError);
            }
        }

        // Office add-ins run in another process, so the runner's retained RCW
        // cannot resolve sibling writes there. Its protected receipt can.
        public static object ResolvePreparedSibling(string progId)
        {
            var host = progId == "Excel.Application" ? "Excel" : progId == "PowerPoint.Application" ? "PowerPoint" : null;
            ValidateHost(host);
            var state = RequireScope(true);
            RequireSta();
            var binding = ReadBinding(state, host, true);
            var application = TryAttach(host, binding.pid, binding.process_start, binding.executable, binding.startup_path, CancellationToken.None);
            if (application == null)
                throw new InvalidOperationException(host + " lost its prepared document window. The Test Lab will not create a different automation destination.");
            try { RequireSameSuite(state.id, true); return application; }
            catch { Release(application); throw; }
        }

        private static async Task<object> WaitForApplication(HostBinding binding, string suiteId, CancellationToken cancel,
            Action<string> log, HashSet<int> attemptedConnect)
        {
            var deadline = DateTime.UtcNow.AddSeconds(StartupSeconds);
            Exception lastError = null;
            while (DateTime.UtcNow < deadline)
            {
                cancel.ThrowIfCancellationRequested(); RequireSameSuite(suiteId, false);
                object app = null;
                try
                {
                    if (!ProcessMatches(binding)) throw new InvalidOperationException("The prepared " + binding.host + " process ended during attachment.");
                    app = TryAttach(binding.host, binding.pid, binding.process_start, binding.executable, binding.startup_path, cancel);
                    if (app != null && HasScribble(app, binding.host, binding.pid, attemptedConnect))
                    { var result = app; app = null; log(binding.host + ": reused its verified interactive process."); return result; }
                }
                catch (COMException error) { lastError = error; }
                finally { Release(app); }
                await Task.Delay(250, cancel);
            }
            throw StartupFailure(binding.host, lastError);
        }

        private static InvalidOperationException StartupFailure(string host, Exception error)
        { return new InvalidOperationException(host + " did not make Scribble and its private document window available within " + StartupSeconds + " seconds. Check the visible startup, sign-in or recovery dialog. Existing documents and the startup file were preserved; no model request was submitted.", error); }

        private static SuiteState RequireScope(bool activeRun)
        {
            var state = TestLabSuite.Active(); var session = TestLab.Status();
            if (state == null || session == null || session.suite_id != state.fixtureSuiteId ||
                !Inside(session.fixture_root, state.folder) || TestLab.FileHash(TestLab.SafeChild(session.fixture_root, "manifest.json")) != session.manifest_sha256)
                throw new InvalidOperationException("A live, verified Test Lab suite and its fixture session are required for Office attachment.");
            if (!activeRun)
            {
                if (!string.IsNullOrEmpty(session.run_id)) throw new InvalidOperationException("Office startup requires an idle fixture session.");
            }
            else
            {
                if (string.IsNullOrEmpty(session.run_id) || state.runId != session.run_id)
                    throw new InvalidOperationException("Prepared sibling attachment requires the active suite case.");
                var run = TestLab.GetRun(session.run_id);
                if (run.session_id != session.session_id || run.suite_id != session.suite_id || run.fixture_root != session.fixture_root || run.manifest_sha256 != session.manifest_sha256 ||
                    run.case_id != state.caseId || run.host != state.host)
                    throw new InvalidOperationException("The active Office case no longer matches its verified fixture session.");
            }
            return state;
        }

        private static void RequireSameSuite(string id, bool activeRun)
        { if (RequireScope(activeRun).id != id) throw new InvalidOperationException("The Test Lab suite changed during Office attachment."); }
        private static void RequireSta()
        { if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA) throw new InvalidOperationException("Office attachment requires the operator's pumped STA thread."); }
        private static void ValidateHost(string host)
        { if (host != "Excel" && host != "PowerPoint") throw new ArgumentException("Only Excel and PowerPoint use this interactive bootstrap.", nameof(host)); }
        private static string Extension(string host) { return host == "Excel" ? ".xlsx" : ".pptx"; }
        private static string ExeName(string host) { return host == "Excel" ? "EXCEL.EXE" : "POWERPNT.EXE"; }
        private static string DirectoryFor(SuiteState state) { return TestLab.NativeDirectory(state.id, "office"); }
        private static string Descriptor(string directory, string host) { return TestLab.SafeChild(directory, host.ToLowerInvariant() + ".binding.bin"); }
        private static bool SamePath(string a, string b)
        { return !string.IsNullOrEmpty(a) && !string.IsNullOrEmpty(b) && string.Equals(Path.GetFullPath(a), Path.GetFullPath(b), StringComparison.OrdinalIgnoreCase); }
        private static bool Inside(string path, string root)
        { return !string.IsNullOrEmpty(path) && !string.IsNullOrEmpty(root) && Path.GetFullPath(path).StartsWith(Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase); }

        private static byte[] StartupBytes(string host)
        {
            using (var stream = typeof(TestLabOfficeConnection).Assembly.GetManifestResourceStream("Scribble.Testing.Bootstrap." + host + Extension(host)))
            {
                if (stream == null) throw new InvalidDataException("The installed Office startup fixture is missing.");
                using (var bytes = new MemoryStream()) { stream.CopyTo(bytes); return bytes.ToArray(); }
            }
        }

        private static HostBinding ReadBinding(SuiteState state, string host, bool requireAlive)
        {
            var directory = DirectoryFor(state); var path = Descriptor(directory, host);
            if (!File.Exists(path)) throw new InvalidOperationException(host + " has no prepared host receipt. The Test Lab will not create another automation destination.");
            byte[] data;
            using (var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read | FileShare.Delete))
            {
                if (file.Length > 64 * 1024) throw new InvalidDataException("The Office host receipt exceeds its size limit.");
                using (var bytes = new MemoryStream()) { file.CopyTo(bytes); data = ProtectedData.Unprotect(bytes.ToArray(), null, DataProtectionScope.CurrentUser); }
            }
            var binding = new JavaScriptSerializer { MaxJsonLength = 64 * 1024 }.Deserialize<HostBinding>(Encoding.UTF8.GetString(data));
            if (binding == null || binding.schema != 1 || binding.suite_id != state.id || binding.host != host || binding.pid <= 0 || binding.process_start <= 0 ||
                string.IsNullOrEmpty(binding.startup_path) || !Regex.IsMatch(Path.GetFileName(binding.startup_path), "^startup-" + host.ToLowerInvariant() + "-[a-f0-9]{32}\\" + Extension(host) + "$") ||
                !SamePath(binding.startup_path, TestLab.SafeChild(directory, Path.GetFileName(binding.startup_path))) ||
                binding.startup_sha256 != TestLab.Hash(StartupBytes(host)) || !File.Exists(binding.startup_path) || StartupHash(binding.startup_path) != binding.startup_sha256 ||
                string.IsNullOrEmpty(binding.executable) || !string.Equals(Path.GetFileName(binding.executable), ExeName(host), StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("The Office host receipt is outside its verified suite/startup boundary.");
            if (requireAlive && !ProcessMatches(binding)) throw new InvalidOperationException(host + " no longer matches its prepared process identity. No replacement automation host was created.");
            return binding;
        }

        private static string StartupHash(string path)
        {
            // Office can retain a write-sharing handle even for read-only opens.
            // Read the bytes under that lock and still require the exact embedded
            // blank package; allowing a reader does not allow different content.
            using (var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (var hash = SHA256.Create())
                return BitConverter.ToString(hash.ComputeHash(file)).Replace("-", "").ToLowerInvariant();
        }

        private static void WriteBinding(string path, HostBinding binding)
        {
            var temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                File.WriteAllBytes(temporary, ProtectedData.Protect(Encoding.UTF8.GetBytes(TestLab.Serialize(binding)), null, DataProtectionScope.CurrentUser));
                if (File.Exists(path)) File.Replace(temporary, path, null); else File.Move(temporary, path);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }

        private static string FindExecutable(string host)
        {
            foreach (var hive in new[] { RegistryHive.LocalMachine, RegistryHive.CurrentUser })
                foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
                    using (var root = RegistryKey.OpenBaseKey(hive, view))
                    using (var key = root.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\" + ExeName(host)))
                    {
                        var path = Convert.ToString(key?.GetValue(""));
                        if (!string.IsNullOrWhiteSpace(path))
                        {
                            path = path.Trim('"');
                            if (Path.IsPathRooted(path) && File.Exists(path) && string.Equals(Path.GetFileName(path), ExeName(host), StringComparison.OrdinalIgnoreCase)) return Path.GetFullPath(path);
                        }
                    }
            throw new InvalidOperationException(host + " has no installed Office executable registered for normal startup.");
        }

        private static List<HostBinding> Processes(string host)
        {
            var result = new List<HostBinding>();
            foreach (var process in Process.GetProcessesByName(Path.GetFileNameWithoutExtension(ExeName(host))))
                using (process)
                    try { if (!process.HasExited) result.Add(new HostBinding { pid = process.Id, process_start = process.StartTime.ToUniversalTime().Ticks, executable = process.MainModule.FileName }); }
                    catch (InvalidOperationException) { }
                    catch (System.ComponentModel.Win32Exception error)
                    { throw new InvalidOperationException("An existing " + host + " process could not be identified. The Test Lab preserved it and did not launch another host.", error); }
            return result;
        }

        private static bool ProcessMatches(HostBinding binding)
        {
            try
            {
                using (var process = Process.GetProcessById(binding.pid))
                    return !process.HasExited && process.StartTime.ToUniversalTime().Ticks == binding.process_start && SamePath(process.MainModule.FileName, binding.executable);
            }
            catch (ArgumentException) { return false; }
            catch (InvalidOperationException) { return false; }
            catch (System.ComponentModel.Win32Exception) { return false; }
        }

        private static bool HasScribble(object application, string host, int pid, HashSet<int> attemptedConnect)
        {
            object collection = null, addin = null, controller = null;
            try
            {
                collection = ((dynamic)application).COMAddIns;
                try { addin = ((dynamic)collection).Item("Scribble." + host + "AddIn"); }
                catch (COMException error) when ((uint)error.HResult == 0x8002000B || (uint)error.HResult == 0x800A0009) { return false; }
                if (!Convert.ToBoolean(((dynamic)addin).Connect) && attemptedConnect.Add(pid)) ((dynamic)addin).Connect = true;
                if (!Convert.ToBoolean(((dynamic)addin).Connect)) return false;
                controller = ((dynamic)addin).Object; return controller != null;
            }
            finally { Release(controller); Release(addin); Release(collection); }
        }

        private static object TryAttach(string host, int pid, long started, string executable, string startup, CancellationToken cancel)
        {
            cancel.ThrowIfCancellationRequested();
            if (!ProcessAlive(pid, started)) return null;
            object candidate = null;
            try
            {
                // PowerPoint.Application has no Hwnd property. Its singleton
                // or native document-window identity is verified below.
                try { if (host == "Excel") candidate = Marshal.GetActiveObject("Excel.Application"); }
                catch (COMException) { }
                if (candidate != null && ApplicationPid(candidate) == pid && (startup == null || FindDocument(candidate, host, startup, false)) && ProcessAlive(pid, started))
                { var result = candidate; candidate = null; return result; }
            }
            finally { Release(candidate); }
            // Current PowerPoint builds can expose mdiClass without paneClassDC
            // or an automation HWND. Its ROT entry is still usable when the
            // complete process inventory proves a single, unchanged approved
            // host and that Application contains our unique startup document.
            if (host == "PowerPoint" && PowerPointSingleton(pid, started, executable))
            {
                candidate = null;
                try
                {
                    try { candidate = Marshal.GetActiveObject("PowerPoint.Application"); }
                    catch (COMException) { }
                    if (candidate != null && (startup == null || FindDocument(candidate, host, startup, false)) &&
                        PowerPointSingleton(pid, started, executable))
                    { var result = candidate; candidate = null; return result; }
                }
                finally { Release(candidate); }
            }
            foreach (var window in Windows(pid, NativeWindowClasses(host)))
            {
                cancel.ThrowIfCancellationRequested(); object native = null; candidate = null;
                try
                {
                    var iid = Dispatch; Marshal.ThrowExceptionForHR(AccessibleObjectFromWindow(window, 0xfffffff0, ref iid, out native));
                    candidate = ((dynamic)native).Application;
                    uint actual; GetWindowThreadProcessId(window, out actual);
                    if (actual == pid && (host != "Excel" || ApplicationPid(candidate) == pid) &&
                        (startup == null || FindDocument(candidate, host, startup, false)) && ProcessAlive(pid, started))
                    { var result = candidate; candidate = null; return result; }
                }
                catch (COMException) { }
                finally { Release(candidate); Release(native); }
            }
            return null;
        }

        private static bool PowerPointSingleton(int pid, long started, string executable)
        {
            var processes = Processes("PowerPoint");
            return processes.Count == 1 && processes[0].pid == pid && processes[0].process_start == started &&
                SamePath(processes[0].executable, executable);
        }

        private static bool IsLaunchCandidate(string host, int beforeCount, bool existedBefore, long processStart, long launchedAt)
        {
            // A normal PowerPoint file launch may reuse the one approved process
            // or create a fresh process. Keep both possibilities in scope; the
            // exact private document and native PID are still verified by
            // TryAttach. Other old processes and all old Excel processes remain
            // outside the candidate set.
            return (host == "PowerPoint" && beforeCount == 1 && existedBefore) ||
                (!existedBefore && processStart >= launchedAt - TimeSpan.FromSeconds(2).Ticks);
        }

        private static string[] NativeWindowClasses(string host)
        {
            return host == "Excel" ? new[] { "EXCEL7" } : new[] { "paneClassDC", "mdiClass" };
        }

        private static bool ProcessAlive(int pid, long started)
        {
            try { using (var process = Process.GetProcessById(pid)) return !process.HasExited && process.StartTime.ToUniversalTime().Ticks == started; }
            catch (ArgumentException) { return false; }
            catch (InvalidOperationException) { return false; }
        }

        private static bool FindDocument(object application, string host, string path, bool close)
        {
            object documents = null;
            try
            {
                documents = host == "Excel" ? ((dynamic)application).Workbooks : ((dynamic)application).Presentations;
                for (int index = 1; index <= Convert.ToInt32(((dynamic)documents).Count); index++)
                {
                    object document = null;
                    try
                    {
                        document = ((dynamic)documents).Item(index);
                        if (!SamePath(Convert.ToString(((dynamic)document).FullName), path)) continue;
                        if (close)
                        {
                            if (!Convert.ToBoolean(((dynamic)document).Saved)) return false;
                            if (host == "Excel") ((dynamic)document).Close(false); else ((dynamic)document).Close();
                        }
                        return true;
                    }
                    finally { Release(document); }
                }
                return false;
            }
            finally { Release(documents); }
        }

        internal static void Cleanup(string host, string suiteId, Action<string> log)
        {
            // An unconfirmed stop deliberately keeps its active capture. Do not
            // close even the startup window while that request may still run.
            if (TestLab.Status() != null) return;
            object application = null;
            try
            {
                var state = new SuiteState { id = suiteId };
                var binding = ReadBinding(state, host, true);
                application = TryAttach(host, binding.pid, binding.process_start, binding.executable, binding.startup_path, CancellationToken.None);
                if (application == null) return;
                if (FindDocument(application, host, binding.startup_path, true)) log(host + ": closed its unchanged private startup document.");
                else log(host + ": retained its private startup document because it was changed or unavailable.");
            }
            catch (Exception error) { log(host + ": preserved startup resources: " + error.Message); }
            finally { Release(application); }
        }

        private static int ApplicationPid(object application)
        { uint pid; GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(((dynamic)application).Hwnd)), out pid); return checked((int)pid); }
        private static void Release(object value) { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
        private static IntPtr[] Windows(int pid, string[] childClasses)
        {
            var result = new List<IntPtr>();
            EnumCallback child = (window, unused) => { uint actual; GetWindowThreadProcessId(window, out actual); if (actual != pid) return true;
                var name = new StringBuilder(128); GetClassName(window, name, name.Capacity); if (childClasses.Contains(name.ToString())) result.Add(window); return true; };
            EnumCallback top = (window, unused) => { uint actual; GetWindowThreadProcessId(window, out actual); if (actual == pid) EnumChildWindows(window, child, IntPtr.Zero); return true; };
            EnumWindows(top, IntPtr.Zero); GC.KeepAlive(child); GC.KeepAlive(top); return result.ToArray();
        }

        [ComVisible(false)]
        private sealed class HostBinding
        {
            public int schema { get; set; } public string suite_id { get; set; } public string host { get; set; }
            public int pid { get; set; } public long process_start { get; set; } public string executable { get; set; }
            public string startup_path { get; set; } public string startup_sha256 { get; set; }
        }
        private delegate bool EnumCallback(IntPtr hwnd, IntPtr parameter);
        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumCallback callback, IntPtr parameter);
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr hwnd, EnumCallback callback, IntPtr parameter);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr hwnd, StringBuilder name, int count);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
        [DllImport("oleacc.dll")] private static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint objectId, ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out object value);
    }
}

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using Scribble.Outlook;

namespace Scribble.Testing
{
    // The operator imports a manifest-verified synthetic mailbox into its own
    // PST. Model tools still enumerate native Outlook tables and opaque IDs.
    // The protected binding contains source identities, never expected answers.
    public static class TestLabMailbox
    {
        private const string SuiteId = "scribble-stress-v1";
        private const string CorpusId = "SCRIBBLE500-V1";
        private const string Marker = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/ScribbleStressMessageId";
        private static readonly object Gate = new object();
        private static Binding cached;
        private static Dictionary<string, CorpusMessage> sources;
        private static string recordedRun;
        private static readonly HashSet<string> recordedSources = new HashSet<string>(StringComparer.Ordinal);

        public static bool Enabled
        {
            get
            {
                var session = TestLab.Status();
                return session?.suite_id == SuiteId || TestLabSuite.Active()?.fixtureSuiteId == SuiteId;
            }
        }

        internal static string ScopeToken()
        {
            if (!Enabled) return null;
            var session = TestLab.Status();
            var state = RequireSuite();
            if (session?.suite_id != SuiteId || string.IsNullOrEmpty(session.run_id))
                throw new InvalidOperationException("The isolated mailbox requires an active synthetic case.");
            RequireBinding(state);
            var run = TestLab.GetRun(session.run_id);
            var manifest = TestLabSuite.Read<KitManifest>(TestLab.SafeChild(session.fixture_root, "manifest.json"));
            if (run.suite_id != SuiteId || run.fixture_root != session.fixture_root || run.manifest_sha256 != session.manifest_sha256 ||
                TestLab.FileHash(TestLab.SafeChild(session.fixture_root, "manifest.json")) != session.manifest_sha256 ||
                !(SamePath(session.fixture_root, state.catalogRoot) || manifest.parent_manifest_sha256 == state.kitHash))
                throw new InvalidDataException("The active mailbox case is not a verified projection of this suite.");
            return state.id + ":" + session.run_id;
        }

        internal static void AssertScope(string token)
        {
            if (token != ScopeToken()) throw new InvalidOperationException("The mailbox scope changed. Start a new request in the active fixture case.");
        }

        private static SuiteState RequireSuite()
        {
            var state = TestLabSuite.Active();
            if (state == null || state.fixtureSuiteId != SuiteId || string.IsNullOrEmpty(state.catalogRoot))
                throw new InvalidOperationException("The isolated mailbox suite is unavailable; real mailbox access remains disabled.");
            var catalog = Path.GetFullPath(state.catalogRoot);
            var suiteRoot = Path.GetFullPath(state.folder).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            if (!catalog.StartsWith(suiteRoot, StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("The mailbox catalog must belong to the active suite snapshot.");
            return state;
        }

        private static string Descriptor(SuiteState state)
        { return Path.Combine(TestLab.NativeDirectory(state.id, "mailbox"), "binding.bin"); }

        private static string PstPath(SuiteState state)
        { return Path.Combine(TestLab.NativeDirectory(state.id, "mailbox"), "synthetic-500.pst"); }

        private static Dictionary<string, CorpusMessage> ReadSources(SuiteState state, out string manifestHash)
        {
            var manifest = TestLab.VerifyKit(state.catalogRoot);
            if (manifest.suite_id != SuiteId) throw new InvalidDataException("Unexpected mailbox fixture suite.");
            manifestHash = TestLab.FileHash(TestLab.SafeChild(state.catalogRoot, "manifest.json"));
            if (!string.Equals(manifestHash, state.kitHash, StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("The mailbox catalog no longer matches the operator-pinned manifest.");
            var listed = new HashSet<string>(manifest.files.Select(f => f.path), StringComparer.Ordinal);
            if (!listed.Contains("operator/mail-index.json")) throw new InvalidDataException("The mailbox source index is not manifest-verified.");
            var records = TestLabSuite.Read<CorpusMessage[]>(TestLab.SafeChild(state.catalogRoot, "operator/mail-index.json"));
            if (records == null || records.Length != 500 || records.Select(m => m.id).Distinct().Count() != 500)
                throw new InvalidDataException("The isolated mailbox requires exactly 500 unique fixture messages.");
            foreach (var m in records)
            {
                if (!Regex.IsMatch(m.id ?? "", "^MAIL[0-9]{4}$") || m.corpus_id != CorpusId ||
                    m.path != "inputs/outlook/" + m.id + ".eml" || !listed.Contains(m.path) ||
                    (m.folder != "inbox" && m.folder != "sent") ||
                    !(m.body ?? "").Contains("Synthetic message ID: " + m.id) ||
                    !(m.sender_email ?? "").EndsWith("@staff.example.test", StringComparison.Ordinal) ||
                    (m.attachments ?? new string[0]).Any(a => !a.StartsWith("inputs/", StringComparison.Ordinal) || !listed.Contains(a)))
                    throw new InvalidDataException("Invalid synthetic source identity: " + m.id);
            }
            return records.ToDictionary(m => m.id, StringComparer.Ordinal);
        }

        private static Binding RequireBinding(SuiteState state)
        {
            lock (Gate)
            {
                var hash = TestLab.FileHash(TestLab.SafeChild(state.catalogRoot, "manifest.json"));
                if (!string.Equals(hash, state.kitHash, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException("The mailbox catalog no longer matches the operator-pinned manifest.");
                if (cached != null && cached.suite_id == state.id && cached.catalog_root == state.catalogRoot && cached.manifest_sha256 == hash)
                    return cached;
                var path = Descriptor(state);
                if (!File.Exists(path)) throw new InvalidOperationException("The isolated 500-message mailbox has not completed preparation.");
                byte[] plain;
                using (var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read | FileShare.Delete))
                {
                    if (file.Length > 2 * 1024 * 1024) throw new InvalidDataException("The isolated mailbox binding exceeds its size limit.");
                    using (var bytes = new MemoryStream()) { file.CopyTo(bytes); plain = ProtectedData.Unprotect(bytes.ToArray(), null, DataProtectionScope.CurrentUser); }
                }
                var binding = new System.Web.Script.Serialization.JavaScriptSerializer { MaxJsonLength = 2 * 1024 * 1024 }.Deserialize<Binding>(Encoding.UTF8.GetString(plain));
                string verifiedHash;
                var records = ReadSources(state, out verifiedHash);
                if (binding == null || binding.schema != 1 || binding.suite_id != state.id || binding.catalog_root != state.catalogRoot ||
                    binding.manifest_sha256 != verifiedHash || !SamePath(binding.pst_path, PstPath(state)) ||
                    string.IsNullOrEmpty(binding.store_id) || string.IsNullOrEmpty(binding.inbox_id) || string.IsNullOrEmpty(binding.sent_id) ||
                    binding.items == null || binding.items.Length != 500 || binding.items.Select(i => i.entry_id).Distinct().Count() != 500 ||
                    binding.items.Select(i => i.id).Distinct().Count() != 500 || binding.items.Any(i => !records.ContainsKey(i.id) || string.IsNullOrEmpty(i.entry_id)))
                    throw new InvalidDataException("The isolated mailbox binding does not match the active verified suite.");
                sources = records; cached = binding; return binding;
            }
        }

        // Public for the operator's native-environment validation, with no LLM
        // request or budget bypass. The same active suite/manifest checks apply.
        public static void Prepare(object application, CancellationToken cancel, Action<string> log)
        {
            var state = RequireSuite();
            if (TestLab.Status()?.run_id != null) throw new InvalidOperationException("Mailbox preparation requires an idle fixture session.");
            string hash;
            var records = ReadSources(state, out hash);
            if (File.Exists(Descriptor(state)))
            {
                var existing = RequireBinding(state);
                VerifyNativeInventory(application, existing, records, cancel);
                Display(application, existing);
                log("Reused verified isolated Outlook corpus: 500 messages.");
                return;
            }
            var pst = PstPath(state);
            // Never erase or repair an unknown/partial PST. A new suite owns a
            // new path, so interrupted import remains reviewable and isolated.
            if (File.Exists(pst)) throw new InvalidOperationException("A partial synthetic PST already exists. Start a new suite to create a fresh isolated corpus.");
            Directory.CreateDirectory(Path.GetDirectoryName(pst));
            object session = null, store = null, root = null, folders = null, inbox = null, sent = null;
            try
            {
                session = ((dynamic)application).Session;
                ((dynamic)session).AddStoreEx(pst, 2); // olStoreUnicode
                store = FindStore(session, pst);
                root = ((dynamic)store).GetRootFolder();
                ((dynamic)root).Name = "Scribble Synthetic 500 — " + state.id.Substring(0, Math.Min(8, state.id.Length));
                folders = ((dynamic)root).Folders;
                inbox = ((dynamic)folders).Add("Corpus Inbox", 6);
                // Folders.Add accepts olFolderInbox for a mail folder, but
                // explicitly rejects olFolderSentMail. Logical Sent scope is
                // bound to this separate folder's verified native identity.
                sent = ((dynamic)folders).Add("Corpus Sent", 6);
                var binding = new Binding { schema = 1, suite_id = state.id, catalog_root = state.catalogRoot, manifest_sha256 = hash,
                    pst_path = pst, store_id = Convert.ToString(((dynamic)store).StoreID), inbox_id = Convert.ToString(((dynamic)inbox).EntryID),
                    sent_id = Convert.ToString(((dynamic)sent).EntryID) };
                var identities = new List<NativeMessage>();
                foreach (var record in records.Values.OrderBy(m => m.id, StringComparer.Ordinal))
                {
                    cancel.ThrowIfCancellationRequested();
                    identities.Add(Import(record.folder == "sent" ? sent : inbox, state.catalogRoot, record));
                    if (identities.Count % 25 == 0) { log("Imported " + identities.Count + "/500 synthetic messages into the isolated PST."); System.Windows.Forms.Application.DoEvents(); }
                }
                binding.items = identities.ToArray();
                VerifyNativeInventory(application, binding, records, cancel);
                var bytes = ProtectedData.Protect(Encoding.UTF8.GetBytes(TestLab.Serialize(binding)), null, DataProtectionScope.CurrentUser);
                var temporary = Descriptor(state) + ".tmp";
                try { File.WriteAllBytes(temporary, bytes); File.Move(temporary, Descriptor(state)); }
                finally { if (File.Exists(temporary)) File.Delete(temporary); }
                lock (Gate) { cached = binding; sources = records; }
                Display(application, binding);
                log("Verified isolated Outlook corpus: 500 unique native messages, 400 Inbox and 100 Sent.");
            }
            finally { Release(sent); Release(inbox); Release(folders); Release(root); Release(store); Release(session); }
        }

        private static NativeMessage Import(object folder, string root, CorpusMessage source)
        {
            object items = null, item = null, moved = null, parent = null, properties = null, attachments = null;
            try
            {
                items = ((dynamic)folder).Items;
                item = ((dynamic)items).Add("IPM.Note");
                dynamic mail = item;
                mail.Subject = source.subject; mail.Body = source.body;
                mail.To = string.Join("; ", source.to ?? new string[0]);
                mail.CC = string.Join("; ", source.cc ?? new string[0]);
                properties = mail.PropertyAccessor;
                dynamic p = properties;
                Set(p, "0x0C1A001F", source.sender.Split('<')[0].Trim());
                Set(p, "0x0C1F001F", source.sender_email); Set(p, "0x0C1E001F", "SMTP");
                Set(p, "0x5D01001F", source.sender_email);
                Set(p, "0x0042001F", source.sender.Split('<')[0].Trim());
                Set(p, "0x0065001F", source.sender_email); Set(p, "0x0064001F", "SMTP");
                var date = DateTimeOffset.Parse(source.date, CultureInfo.InvariantCulture).UtcDateTime;
                Set(p, "0x00390040", date); Set(p, "0x0E060040", date);
                Set(p, "0x1035001F", source.message_id); Set(p, "0x0070001F", source.thread_topic);
                Set(p, "0x00710102", Convert.FromBase64String(source.conversation_index));
                if (!string.IsNullOrEmpty(source.in_reply_to)) Set(p, "0x1042001F", source.in_reply_to);
                if (source.references != null && source.references.Length > 0) Set(p, "0x1039001F", string.Join(" ", source.references));
                p.SetProperty(Marker, source.id);
                attachments = mail.Attachments;
                foreach (var relative in source.attachments ?? new string[0])
                { object attachment = ((dynamic)attachments).Add(TestLab.SafeChild(root, relative)); Release(attachment); }
                // This is operator-owned corpus authoring inside the new PST;
                // it never queues or sends a message in a real account.
                Set(p, "0x0E070003", source.unread ? 0 : 1);
                // Outlook may give a new MailItem the default Outbox as its
                // parent even when created through another folder's Items.
                // Move the still-unsaved operator fixture directly to the
                // verified PST; Save alone can persist it in the real store.
                moved = mail.Move(folder);
                parent = ((dynamic)moved).Parent;
                if (Convert.ToString(((dynamic)parent).StoreID) != Convert.ToString(((dynamic)folder).StoreID) ||
                    Convert.ToString(((dynamic)parent).EntryID) != Convert.ToString(((dynamic)folder).EntryID))
                    throw new InvalidDataException("Outlook did not place the new synthetic message in its owned PST folder.");
                // UnRead can persist changes. Apply it only after the native
                // parent has been verified inside the operator-owned PST.
                ((dynamic)moved).UnRead = source.unread;
                ((dynamic)moved).Save();
                return new NativeMessage { id = source.id, entry_id = Convert.ToString(((dynamic)moved).EntryID) };
            }
            finally { Release(parent); Release(moved); Release(attachments); Release(properties); Release(item); Release(items); }
        }

        private static void Set(dynamic properties, string tag, object value)
        { properties.SetProperty("http://schemas.microsoft.com/mapi/proptag/" + tag, value); }

        private static object FindStore(object session, string path)
        {
            object stores = null;
            try
            {
                stores = ((dynamic)session).Stores;
                var count = Convert.ToInt32(((dynamic)stores).Count);
                for (var i = 1; i <= count; i++)
                {
                    object store = ((dynamic)stores).Item(i);
                    try
                    {
                        if (SamePath(Convert.ToString(((dynamic)store).FilePath), path)) { var result = store; store = null; return result; }
                    }
                    finally { Release(store); }
                }
                throw new InvalidOperationException("The owned synthetic PST is not attached; real mailbox access remains disabled.");
            }
            finally { Release(stores); }
        }

        private static bool SamePath(string first, string second)
        { return !string.IsNullOrEmpty(first) && string.Equals(Path.GetFullPath(first), Path.GetFullPath(second), StringComparison.OrdinalIgnoreCase); }

        internal static object GetFolder(object application, int kind)
        {
            var state = RequireSuite();
            var binding = RequireBinding(state);
            return BoundFolder(application, binding, kind);
        }

        private static object BoundFolder(object application, Binding binding, int kind)
        {
            if (kind != 6 && kind != 5) throw new InvalidOperationException("Only synthetic Inbox and Sent folders are in scope.");
            object session = null, store = null, folder = null;
            try
            {
                session = ((dynamic)application).Session;
                store = FindStore(session, binding.pst_path);
                if (Convert.ToString(((dynamic)store).StoreID) != binding.store_id) throw new InvalidDataException("The isolated PST store identity changed.");
                folder = ((dynamic)session).GetFolderFromID(kind == 6 ? binding.inbox_id : binding.sent_id, binding.store_id);
                if (Convert.ToString(((dynamic)folder).StoreID) != binding.store_id) throw new InvalidDataException("The synthetic folder is outside the owned PST.");
                var result = folder; folder = null; return result;
            }
            finally { Release(folder); Release(store); Release(session); }
        }

        private static void Display(object application, Binding binding)
        {
            object inbox = null, explorer = null;
            try
            {
                inbox = BoundFolder(application, binding, 6); explorer = ((dynamic)inbox).GetExplorer();
                ((dynamic)explorer).Display();
                // The reading pane must not mark the newest fixture read merely
                // because the operator opened the corpus explorer.
                ((dynamic)explorer).ClearSelection();
            }
            finally { Release(explorer); Release(inbox); }
        }

        private static void VerifyNativeInventory(object application, Binding binding, Dictionary<string, CorpusMessage> records, CancellationToken cancel)
        {
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (var kind in new[] { 6, 5 })
            {
                object folder = null, table = null, session = null;
                try
                {
                    folder = BoundFolder(application, binding, kind);
                    table = ((dynamic)folder).GetTable(); session = ((dynamic)application).Session;
                    while (!Convert.ToBoolean(((dynamic)table).EndOfTable))
                    {
                        cancel.ThrowIfCancellationRequested();
                        object row = null, item = null;
                        try
                        {
                            row = ((dynamic)table).GetNextRow();
                            var entry = Convert.ToString(((dynamic)row)["EntryID"]);
                            var identity = binding.items.FirstOrDefault(m => m.entry_id == entry);
                            if (identity == null || !seen.Add(identity.id) || records[identity.id].folder != (kind == 6 ? "inbox" : "sent"))
                                throw new InvalidDataException("The native PST inventory contains an unexpected or moved item.");
                            item = ((dynamic)session).GetItemFromID(entry, binding.store_id);
                            ValidateNative(item, binding, records[identity.id]);
                            if (seen.Count % 25 == 0) System.Windows.Forms.Application.DoEvents();
                        }
                        finally { Release(item); Release(row); }
                    }
                }
                finally { Release(session); Release(table); Release(folder); }
            }
            if (seen.Count != 500) throw new InvalidDataException("The native PST is missing synthetic messages.");
        }

        public static void ValidateIdentity(string entryId, string storeId)
        {
            if (!Enabled) return;
            var binding = RequireBinding(RequireSuite());
            if (storeId != binding.store_id || !binding.items.Any(i => i.entry_id == entryId))
                throw new InvalidOperationException("This item is outside the verified synthetic PST; real mailbox reads are disabled during the case.");
        }

        internal static bool TryResolveAttachment(string entryId, string storeId, int index,
            out string path, out string fileName)
        {
            path = null;
            fileName = null;
            if (!Enabled) return false;

            ScopeToken();
            var state = RequireSuite();
            var binding = RequireBinding(state);
            if (storeId != binding.store_id)
                throw new InvalidOperationException("This attachment is outside the verified synthetic PST.");
            var identity = binding.items.FirstOrDefault(i => i.entry_id == entryId);
            if (identity == null)
                throw new InvalidOperationException("This attachment is outside the verified synthetic PST.");
            var source = sources[identity.id];
            if (index < 1 || index > (source.attachments ?? new string[0]).Length)
                throw new ArgumentException("Attachment index is outside the captured message.");

            var relative = source.attachments[index - 1];
            var candidate = TestLab.SafeChild(state.catalogRoot, relative);
            var manifest = TestLabSuite.Read<KitManifest>(TestLab.SafeChild(state.catalogRoot, "manifest.json"));
            var declared = (manifest.files ?? new KitFile[0]).FirstOrDefault(f => f.path == relative);
            if (declared == null || !File.Exists(candidate) || new FileInfo(candidate).Length != declared.size ||
                !string.Equals(TestLab.FileHash(candidate), declared.sha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("The verified synthetic attachment changed: " + identity.id);

            path = candidate;
            fileName = Path.GetFileName(relative);
            return true;
        }

        public static void ValidateSource(object item)
        {
            var state = RequireSuite();
            var binding = RequireBinding(state);
            var entryId = Convert.ToString(((dynamic)item).EntryID);
            var identity = binding.items.FirstOrDefault(i => i.entry_id == entryId);
            if (identity == null) throw new InvalidOperationException("This selected item is outside the verified synthetic PST.");
            ValidateNative(item, binding, sources[identity.id]);
            var runId = TestLab.ActiveRunId();
            if (runId != null)
            {
                lock (Gate)
                {
                    if (recordedRun != runId) { recordedRun = runId; recordedSources.Clear(); }
                    if (recordedSources.Add(identity.id)) TestLab.Record(runId, "source", "mail_verified", new { source_id = sources[identity.id].path, corpus_id = CorpusId,
                        native_entry_id = entryId, native_store_id = binding.store_id, body_sha256 = TestLab.Hash(Encoding.UTF8.GetBytes(sources[identity.id].body)) });
                }
            }
        }

        private static void ValidateNative(object item, Binding binding, CorpusMessage source)
        {
            object parent = null, properties = null, attachments = null;
            try
            {
                dynamic mail = item; parent = mail.Parent;
                var folderId = source.folder == "sent" ? binding.sent_id : binding.inbox_id;
                if (Convert.ToString(((dynamic)parent).StoreID) != binding.store_id || Convert.ToString(((dynamic)parent).EntryID) != folderId)
                    throw new InvalidDataException("The native source is outside its isolated fixture folder.");
                properties = mail.PropertyAccessor;
                if (Convert.ToString(((dynamic)properties).GetProperty(Marker)) != source.id || Convert.ToString(mail.Subject) != source.subject ||
                    !string.Equals(Convert.ToString(mail.SenderEmailAddress), source.sender_email, StringComparison.OrdinalIgnoreCase) ||
                    Normalize(Convert.ToString(mail.Body)) != Normalize(source.body) ||
                    Convert.ToDateTime(mail.ReceivedTime).ToUniversalTime() != DateTimeOffset.Parse(source.date, CultureInfo.InvariantCulture).UtcDateTime)
                    throw new InvalidDataException("The native synthetic message differs from its verified source: " + source.id);
                attachments = mail.Attachments;
                if (Convert.ToInt32(((dynamic)attachments).Count) != (source.attachments ?? new string[0]).Length)
                    throw new InvalidDataException("The native synthetic attachment inventory changed: " + source.id);
                for (var i = 0; i < (source.attachments ?? new string[0]).Length; i++)
                {
                    object attachment = null;
                    try
                    {
                        attachment = ((dynamic)attachments).Item(i + 1);
                        if (Convert.ToString(((dynamic)attachment).FileName) != Path.GetFileName(source.attachments[i]))
                            throw new InvalidDataException("The native synthetic attachment identity changed: " + source.id);
                    }
                    finally { Release(attachment); }
                }
            }
            finally { Release(attachments); Release(properties); Release(parent); }
        }

        private static string Normalize(string value) { return Regex.Replace(value ?? "", @"\s+", " ").Trim(); }

        internal static IReadOnlyList<MessageSnapshot> Load(object application, LabCase testCase)
        {
            ScopeToken();
            var binding = RequireBinding(RequireSuite());
            var selected = new HashSet<string>(testCase.inputs ?? new string[0], StringComparer.Ordinal);
            var reader = new MessageReader(application);
            return binding.items.Where(i => selected.Contains(sources[i.id].path)).Select(i => reader.CaptureById(i.entry_id, binding.store_id)).ToArray();
        }

        private static void Release(object value) { TestLabOfficeEnvironment.Release(value); }

        public sealed class Binding
        {
            public int schema { get; set; }
            public string suite_id { get; set; }
            public string catalog_root { get; set; }
            public string manifest_sha256 { get; set; }
            public string pst_path { get; set; }
            public string store_id { get; set; }
            public string inbox_id { get; set; }
            public string sent_id { get; set; }
            public NativeMessage[] items { get; set; }
        }
        public sealed class NativeMessage { public string id { get; set; } public string entry_id { get; set; } }
        public sealed class CorpusMessage
        {
            public string id { get; set; } public string path { get; set; } public string subject { get; set; }
            public string sender { get; set; } public string sender_email { get; set; } public string body { get; set; }
            public string date { get; set; } public string[] attachments { get; set; } public string[] to { get; set; }
            public string[] cc { get; set; } public string folder { get; set; } public bool unread { get; set; }
            public string corpus_id { get; set; } public string message_id { get; set; } public string thread_topic { get; set; }
            public string conversation_index { get; set; } public string in_reply_to { get; set; } public string[] references { get; set; }
        }
    }
}

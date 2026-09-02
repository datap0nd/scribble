using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Web.Script.Serialization;
using Scribble.Configuration;
using Scribble.Security;

namespace Scribble.Chat
{
    internal sealed class StoredTopicToolSession
    {
        public string ChatId { get; set; }

        public string TurnId { get; set; }

        public string TopicId { get; set; }

        public long ExpiresUtcTicks { get; set; }

        public bool SearchExecuted { get; set; }

        public List<string> LoadedHandles { get; set; }

        public int LoadedCharacters { get; set; }

        public Dictionary<string, string> Handles { get; set; }
    }

    public sealed class TopicToolHost
    {
        public const int MaxSearchResults = 10;
        public const int MaxReadDocuments = 3;
        public const int MaxReadCharacters = 120000;
        public const int SessionMinutes = 15;
        public const int MaxSerializedResultCharacters =
            (MaxReadCharacters * 6) + 8192;

        private readonly TopicConfig _topic;
        private readonly string _chatId;
        private readonly string _turnId;
        private readonly bool _persistent;
        private readonly TopicIndex _index = new TopicIndex();
        private readonly JavaScriptSerializer _serializer =
            new JavaScriptSerializer
            {
                MaxJsonLength = int.MaxValue
            };
        private readonly Dictionary<string, string> _handles =
            new Dictionary<string, string>(StringComparer.Ordinal);
        private readonly HashSet<string> _loadedHandles =
            new HashSet<string>(StringComparer.Ordinal);
        private bool _searchExecuted;
        private int _loadedCharacters;

        public TopicToolHost(
            TopicConfig topic,
            string chatId,
            string turnId,
            bool persistent)
        {
            _topic = topic?.Sanitized() ??
                throw new ArgumentNullException(nameof(topic));
            _chatId = BoundScopeId(chatId);
            _turnId = BoundScopeId(turnId);
            _persistent = persistent;
            if (_persistent)
            {
                CleanupExpiredPersistentSessions();
                RestoreSession();
            }
        }

        public MailboxToolResult Execute(
            ChatToolCall call,
            CancellationToken cancellationToken)
        {
            if (call?.function == null ||
                string.IsNullOrWhiteSpace(call.id) ||
                !TopicToolCatalog.IsTopicTool(
                    call.function.name))
            {
                return Error(
                    call?.id,
                    "TOPIC_TOOL_NOT_ALLOWED",
                    "The requested Topic tool is not allowed.");
            }

            try
            {
                var arguments = ParseArguments(
                    call.function.arguments);
                if (string.Equals(
                        call.function.name,
                        TopicToolCatalog.SearchTopic,
                        StringComparison.Ordinal))
                {
                    return Search(
                        call.id,
                        arguments,
                        cancellationToken);
                }

                return Read(
                    call.id,
                    arguments,
                    cancellationToken);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception exception)
            {
                return Error(
                    call.id,
                    "TOPIC_TOOL_FAILED",
                    "The active Topic could not be read (" +
                    exception.GetType().Name + ").");
            }
        }

        public void CompleteSession()
        {
            if (!_persistent)
            {
                return;
            }

            var path = SessionPath(_turnId);
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
            }
        }

        public static void ClearPersistentChat(string chatId)
        {
            var bounded = BoundScopeId(chatId);
            var serializer = new JavaScriptSerializer();
            var directory = SessionsDirectory();
            if (!Directory.Exists(directory))
            {
                return;
            }

            foreach (var path in Directory.EnumerateFiles(
                directory,
                "*.json",
                SearchOption.TopDirectoryOnly))
            {
                try
                {
                    var session = serializer
                        .Deserialize<StoredTopicToolSession>(
                            File.ReadAllText(path));
                    if (session != null && string.Equals(
                            session.ChatId,
                            bounded,
                            StringComparison.Ordinal))
                    {
                        File.Delete(path);
                    }
                }
                catch
                {
                }
            }
        }

        private MailboxToolResult Search(
            string callId,
            IDictionary<string, object> arguments,
            CancellationToken cancellationToken)
        {
            if (_searchExecuted)
            {
                return Error(
                    callId,
                    "TOPIC_SEARCH_LIMIT_REACHED",
                    "Only one Topic search is allowed per request.");
            }

            _searchExecuted = true;
            var query = GetString(arguments, "query", string.Empty);
            var maximum = GetInteger(
                arguments,
                "max_results",
                MaxSearchResults,
                1,
                MaxSearchResults);
            var hits = _index.Search(
                _topic,
                query,
                maximum,
                cancellationToken);
            var results = new List<object>();
            foreach (var hit in hits)
            {
                var handle = "topic_" +
                    Guid.NewGuid().ToString("N");
                _handles[handle] = hit.Entry.RelativePath;
                results.Add(new Dictionary<string, object>
                {
                    { "handle", handle },
                    { "relative_path", hit.Entry.RelativePath },
                    {
                        "modified_utc",
                        new DateTime(
                            hit.Entry.ModifiedUtcTicks,
                            DateTimeKind.Utc).ToString(
                                "O",
                                CultureInfo.InvariantCulture)
                    },
                    { "snippet", hit.Snippet }
                });
            }

            SaveSession();
            return Success(
                callId,
                new Dictionary<string, object>
                {
                    { "untrusted_topic_data", true },
                    { "topic", _topic.Name },
                    { "query", query },
                    { "result_count", results.Count },
                    { "results", results }
                },
                "Topic search loaded " + results.Count +
                (results.Count == 1 ? " result." : " results."));
        }

        private MailboxToolResult Read(
            string callId,
            IDictionary<string, object> arguments,
            CancellationToken cancellationToken)
        {
            var requested = GetStringList(arguments, "handles")
                .Distinct(StringComparer.Ordinal)
                .Take(MaxReadDocuments)
                .ToArray();
            if (requested.Length == 0)
            {
                return Error(
                    callId,
                    "TOPIC_HANDLES_REQUIRED",
                    "At least one Topic handle is required.");
            }

            var remaining = Math.Max(
                0,
                MaxReadCharacters - _loadedCharacters);
            var documents = new List<object>();
            var loaded = 0;
            foreach (var handle in requested)
            {
                cancellationToken.ThrowIfCancellationRequested();
                string relativePath;
                if (!_handles.TryGetValue(handle, out relativePath))
                {
                    documents.Add(new Dictionary<string, object>
                    {
                        { "handle", handle },
                        { "error_code", "TOPIC_HANDLE_UNKNOWN" }
                    });
                    continue;
                }

                if (_loadedHandles.Contains(handle))
                {
                    documents.Add(new Dictionary<string, object>
                    {
                        { "handle", handle },
                        { "already_loaded", true }
                    });
                    continue;
                }

                if (_loadedHandles.Count >= MaxReadDocuments ||
                    remaining <= 0)
                {
                    documents.Add(new Dictionary<string, object>
                    {
                        { "handle", handle },
                        { "error_code", "TOPIC_CONTEXT_LIMIT_REACHED" }
                    });
                    continue;
                }

                var entry = _index.Find(_topic, relativePath);
                var validationError =
                    "The indexed file is no longer available.";
                if (entry == null || !_index.Revalidate(
                        _topic,
                        entry,
                        out validationError))
                {
                    documents.Add(new Dictionary<string, object>
                    {
                        { "handle", handle },
                        { "error_code", "TOPIC_FILE_STALE" },
                        { "message", validationError }
                    });
                    continue;
                }

                var content = TextBoundary.PlainText(
                    entry.Content,
                    Math.Min(
                        TopicIndex.MaxCharactersPerFile,
                        remaining));
                documents.Add(new Dictionary<string, object>
                {
                    { "handle", handle },
                    { "relative_path", entry.RelativePath },
                    { "content", content }
                });
                remaining -= content.Length;
                _loadedCharacters += content.Length;
                _loadedHandles.Add(handle);
                loaded++;
            }

            SaveSession();
            return Success(
                callId,
                new Dictionary<string, object>
                {
                    { "untrusted_topic_data", true },
                    { "topic", _topic.Name },
                    { "documents", documents }
                },
                "Loaded " + loaded +
                (loaded == 1
                    ? " Topic document."
                    : " Topic documents."));
        }

        private MailboxToolResult Success(
            string callId,
            object payload,
            string status)
        {
            var json = _serializer.Serialize(payload);
            // The semantic document budget is 120,000 characters. JSON may
            // escape each control character as six wire characters, so keep
            // a separate finite transport bound instead of rejecting a valid
            // bounded read solely because of encoding overhead.
            if (json.Length > MaxSerializedResultCharacters)
            {
                return Error(
                    callId,
                    "TOPIC_RESULT_TOO_LARGE",
                    "The bounded Topic result was still too large.");
            }

            return new MailboxToolResult(
                callId,
                json,
                status,
                null,
                MaxSerializedResultCharacters);
        }

        private MailboxToolResult Error(
            string callId,
            string code,
            string message)
        {
            return new MailboxToolResult(
                callId ?? string.Empty,
                _serializer.Serialize(
                    new Dictionary<string, object>
                    {
                        { "ok", false },
                        { "error_code", code },
                        { "message", TextBoundary.PlainText(message, 800) }
                    }),
                "[" + code + "] " + message);
        }

        private IDictionary<string, object> ParseArguments(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                return new Dictionary<string, object>();
            }

            return _serializer.DeserializeObject(json) as
                IDictionary<string, object> ??
                new Dictionary<string, object>();
        }

        private static string GetString(
            IDictionary<string, object> arguments,
            string key,
            string fallback)
        {
            object value;
            return arguments.TryGetValue(key, out value)
                ? TextBoundary.SingleLine(Convert.ToString(value), 240)
                : fallback;
        }

        private static int GetInteger(
            IDictionary<string, object> arguments,
            string key,
            int fallback,
            int minimum,
            int maximum)
        {
            object value;
            int parsed;
            if (!arguments.TryGetValue(key, out value) ||
                !int.TryParse(
                    Convert.ToString(value),
                    NumberStyles.Integer,
                    CultureInfo.InvariantCulture,
                    out parsed))
            {
                return fallback;
            }

            return Math.Max(minimum, Math.Min(maximum, parsed));
        }

        private static IEnumerable<string> GetStringList(
            IDictionary<string, object> arguments,
            string key)
        {
            object value;
            if (!arguments.TryGetValue(key, out value))
            {
                return new string[0];
            }

            var array = value as object[];
            if (array != null)
            {
                return array.Select(item =>
                    TextBoundary.SingleLine(Convert.ToString(item), 100));
            }

            var list = value as ArrayList;
            if (list != null)
            {
                return list.Cast<object>().Select(item =>
                    TextBoundary.SingleLine(Convert.ToString(item), 100));
            }

            return new[]
            {
                TextBoundary.SingleLine(Convert.ToString(value), 100)
            };
        }

        private void RestoreSession()
        {
            try
            {
                var path = SessionPath(_turnId);
                if (!File.Exists(path))
                {
                    return;
                }

                var stored = _serializer
                    .Deserialize<StoredTopicToolSession>(
                        File.ReadAllText(path));
                if (stored == null ||
                    stored.ExpiresUtcTicks < DateTime.UtcNow.Ticks ||
                    !string.Equals(stored.ChatId, _chatId,
                        StringComparison.Ordinal) ||
                    !string.Equals(stored.TurnId, _turnId,
                        StringComparison.Ordinal) ||
                    !string.Equals(stored.TopicId, _topic.Id,
                        StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                _searchExecuted = stored.SearchExecuted;
                _loadedCharacters = Math.Max(
                    0,
                    Math.Min(
                        MaxReadCharacters,
                        stored.LoadedCharacters));
                foreach (var pair in stored.Handles ??
                    new Dictionary<string, string>())
                {
                    _handles[pair.Key] = pair.Value;
                }

                foreach (var handle in stored.LoadedHandles ??
                    new List<string>())
                {
                    _loadedHandles.Add(handle);
                }
            }
            catch
            {
            }
        }

        private void SaveSession()
        {
            if (!_persistent)
            {
                return;
            }

            var directory = SessionsDirectory();
            Directory.CreateDirectory(directory);
            var path = SessionPath(_turnId);
            var temporary = path + "." +
                Guid.NewGuid().ToString("N") + ".tmp";
            var stored = new StoredTopicToolSession
            {
                ChatId = _chatId,
                TurnId = _turnId,
                TopicId = _topic.Id,
                ExpiresUtcTicks = DateTime.UtcNow
                    .AddMinutes(SessionMinutes).Ticks,
                SearchExecuted = _searchExecuted,
                LoadedCharacters = _loadedCharacters,
                LoadedHandles = _loadedHandles.ToList(),
                Handles = new Dictionary<string, string>(_handles)
            };
            File.WriteAllText(temporary, _serializer.Serialize(stored));
            try
            {
                if (File.Exists(path))
                {
                    File.Replace(temporary, path, null);
                }
                else
                {
                    File.Move(temporary, path);
                }
            }
            finally
            {
                if (File.Exists(temporary))
                {
                    File.Delete(temporary);
                }
            }
        }

        public static void CleanupExpiredPersistentSessions()
        {
            var directory = SessionsDirectory();
            if (!Directory.Exists(directory))
            {
                return;
            }

            var serializer = new JavaScriptSerializer
            {
                MaxJsonLength = int.MaxValue
            };

            foreach (var path in Directory.EnumerateFiles(
                directory,
                "*",
                SearchOption.TopDirectoryOnly))
            {
                try
                {
                    if (path.EndsWith(
                        ".tmp",
                        StringComparison.OrdinalIgnoreCase))
                    {
                        File.Delete(path);
                        continue;
                    }

                    if (!path.EndsWith(
                        ".json",
                        StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var stored = serializer
                        .Deserialize<StoredTopicToolSession>(
                            File.ReadAllText(path));
                    if (stored == null ||
                        stored.ExpiresUtcTicks < DateTime.UtcNow.Ticks)
                    {
                        File.Delete(path);
                    }
                }
                catch
                {
                    try
                    {
                        File.Delete(path);
                    }
                    catch
                    {
                    }
                }
            }
        }

        private static string BoundScopeId(string value)
        {
            Guid parsed;
            return Guid.TryParse(value, out parsed)
                ? parsed.ToString("N")
                : Guid.NewGuid().ToString("N");
        }

        private static string SessionsDirectory()
        {
            return Path.Combine(
                Environment.GetFolderPath(
                    Environment.SpecialFolder.LocalApplicationData),
                "Scribble",
                "topic-sessions");
        }

        private static string SessionPath(string turnId)
        {
            return Path.Combine(SessionsDirectory(), turnId + ".json");
        }
    }
}

using System;
using System.Collections.Generic;
using System.Text;
using Scribble.Security;

namespace Scribble.Utilities
{
    // Per-pane field diagnostics: a bounded in-memory record of the
    // last few chat requests - the local intent-gate verdict, which
    // tools were exposed to the model, which tool calls ran, and how
    // the request ended. Copied to the clipboard on demand so a
    // misbehaving request can be reported precisely. Never contains
    // API keys, tokens, or settings; message text stays out except
    // for bounded status lines the pane already displayed.
    public sealed class DiagnosticsRecorder
    {
        public const int MaxRequests = 5;
        private const int MaxEventsPerRequest = 128;
        private const int MaxLineCharacters = 300;

        private sealed class RequestRecord
        {
            internal string StartedAt = string.Empty;
            internal string Host = string.Empty;
            internal string Model = string.Empty;
            internal string DiagnosticId = string.Empty;
            internal bool DraftAllowed;
            internal long TotalEvents;
            internal string ExposedTools = string.Empty;
            internal readonly List<string> Events =
                new List<string>();
            internal string Outcome = "(in progress)";
        }

        private readonly object _gate = new object();
        private readonly List<RequestRecord> _requests =
            new List<RequestRecord>();

        public void BeginRequest(
            string host,
            string model,
            bool draftAllowed)
        {
            lock (_gate)
            {
                _requests.Add(new RequestRecord
                {
                    StartedAt = DateTime.Now.ToString(
                        "yyyy-MM-dd HH:mm:ss"),
                    Host = Line(host),
                    Model = Line(model),
                    DraftAllowed = draftAllowed
                });
                if (_requests.Count > MaxRequests)
                {
                    _requests.RemoveAt(0);
                }
            }
        }

        public void SetExposedTools(
            IEnumerable<string> toolNames)
        {
            var names = new List<string>();
            foreach (var name in
                toolNames ?? new string[0])
            {
                names.Add(Line(name));
            }

            lock (_gate)
            {
                var current = Current();
                if (current != null)
                {
                    current.ExposedTools = Line(
                        string.Join(", ", names));
                }
            }
        }

        public void BindTask(string diagnosticId, string effectiveModel)
        {
            lock (_gate)
            {
                var current = Current();
                if (current == null) return;
                current.DiagnosticId = Line(diagnosticId);
                current.Model = Line(effectiveModel);
            }
        }

        public void RecordEvent(string text)
        {
            lock (_gate)
            {
                var current = Current();
                if (current != null)
                {
                    current.TotalEvents++;
                    current.Events.Add(Line(text));
                    if (current.Events.Count > MaxEventsPerRequest) current.Events.RemoveAt(0);
                }
            }
        }

        public void CompleteRequest(string outcome)
        {
            lock (_gate)
            {
                var current = Current();
                if (current != null)
                {
                    current.Outcome = Line(outcome);
                }
            }
        }

        public string BuildReport(string paneDescription)
        {
            var report = new StringBuilder();
            report.AppendLine("Scribble diagnostics");
            report.AppendLine(
                "Pane: " + Line(paneDescription));
            report.AppendLine(
                "Version: " + SelfUpdater.InstalledVersion());
            report.AppendLine(
                "Copied: " + DateTime.Now.ToString(
                    "yyyy-MM-dd HH:mm:ss"));
            lock (_gate)
            {
                if (_requests.Count == 0)
                {
                    report.AppendLine(
                        "No requests recorded in this pane yet.");
                }

                for (var index = 0;
                     index < _requests.Count;
                     index++)
                {
                    var request = _requests[index];
                    report.AppendLine();
                    report.AppendLine(
                        "--- Request " + (index + 1) + " of " +
                        _requests.Count + " (" +
                        request.StartedAt + ") ---");
                    report.AppendLine(
                        "Host: " + request.Host +
                        "  Model: " + request.Model);
                    if (request.DiagnosticId.Length > 0)
                        report.AppendLine("Local diagnostic ID: " + request.DiagnosticId);
                    report.AppendLine(
                        "Draft intent gate: " +
                        (request.DraftAllowed
                            ? "UNLOCKED by the prompt"
                            : "locked (read-only tools)"));
                    if (request.ExposedTools.Length > 0)
                    {
                        report.AppendLine(
                            "Tools exposed: " +
                            request.ExposedTools);
                    }

                    report.AppendLine("Events: " + request.TotalEvents + "; retained tail: " + request.Events.Count);
                    foreach (var entry in request.Events)
                    {
                        report.AppendLine("  " + entry);
                    }

                    report.AppendLine(
                        "Outcome: " + request.Outcome);
                }
            }

            return report.ToString();
        }

        private RequestRecord Current()
        {
            return _requests.Count > 0
                ? _requests[_requests.Count - 1]
                : null;
        }

        private static string Line(string value)
        {
            return TextBoundary.SingleLine(
                value,
                MaxLineCharacters);
        }
    }
}

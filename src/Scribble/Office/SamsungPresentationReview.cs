using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using Scribble.Chat;

namespace Scribble.Office
{
    public static class SamsungPresentationReview
    {
        public const string AuthoringInstructions =
            "For substantial slide decks, establish the brief before drafting: audience, intended decision, depth, and source completeness. " +
            "Use ask_user for two or three relevant missing details; do not repeat answers already supplied. If the user asks you to proceed, make explicit assumptions. " +
            "Read sources fully, then outline the storyline and choose a Samsung layout for each slide. " +
            "A slide needs a concise title and a separate single-line action title in subtitle. " +
            "Use sources for exact citations and evidence for one contiguous verbatim source passage supporting all numbers on that slide. " +
            "Split slides whose evidence spans unrelated sources. Never invent numbers or quotes. " +
            "Use takeaway for the conclusion, and highlight_rows for the data rows/categories that support it. " +
            "The host independently checks evidence, renders editable slides, reviews each rendered image, and repairs owned draft shapes. " +
            "Never claim completion when a review reports a blocker. Themes and positions are host-controlled. " +
            "Content may be quantitative tables/charts, diagrams, action lists, or concise summaries as appropriate; do not invent a table just to fill space.";

        public static void ValidateEvidence(string slideJson, string actualSource)
        {
            var json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            var data = json.Deserialize<Dictionary<string, object>>(slideJson);
            object raw;
            var evidence = data.TryGetValue("evidence", out raw) ? Convert.ToString(raw) : "";
            var layout = data.TryGetValue("layout", out raw) ? Convert.ToString(raw) : "";
            var special = new[] { "cover", "divider", "closing", "agenda" }.Contains(layout);
            if (!special && string.IsNullOrWhiteSpace(evidence)) throw new InvalidOperationException("SLIDE_EVIDENCE_REQUIRED: Supply the verbatim supporting source passage.");
            if (!string.IsNullOrWhiteSpace(evidence) && !(actualSource ?? "").Contains(evidence))
                throw new InvalidOperationException("SLIDE_EVIDENCE_UNVERIFIED: The excerpt does not occur in the original input or read receipts. Read the source and copy an exact passage.");
            if (special) return;
            // Metadata is not a numeric claim. Layout indices and outline levels
            // are not facts either. Inspect displayed content recursively.
            var content = string.Join(" ", data.Where(p => !new[] { "sources", "evidence", "layout", "highlight_rows" }.Contains(p.Key)).Select(p => json.Serialize(p.Value)));
            var allowed = new HashSet<string>(Numbers(evidence));
            var missing = Numbers(content).Where(n => !allowed.Contains(n)).Distinct().ToArray();
            if (missing.Length > 0) throw new InvalidOperationException("SLIDE_NUMBERS_UNVERIFIED: Values absent from cited evidence: " + string.Join(", ", missing));
            if (!data.TryGetValue("subtitle", out raw) || string.IsNullOrWhiteSpace(Convert.ToString(raw)))
                throw new InvalidOperationException("SLIDE_ACTION_TITLE_REQUIRED");
            if (!data.TryGetValue("sources", out raw) || string.IsNullOrWhiteSpace(Convert.ToString(raw)))
                throw new InvalidOperationException("SLIDE_CITATION_REQUIRED");
        }
        private static IEnumerable<string> Numbers(string text)
        { return Regex.Matches(text ?? "", @"(?<![A-Za-z])[-+]?\d+(?:[,.]\d+)*%?").Cast<Match>().Select(m => m.Value.Replace(",", "").TrimStart('+').TrimEnd('%')); }

        internal static string SourceCorpus(TaskContextManager task, string prompt)
        {
            var sources = new List<string> { prompt ?? "" };
            if (task == null) return string.Join("\n", sources);
            sources.Add(task.State.Objective);
            if (task.State.HostData.ContainsKey("recovery_input"))
            {
                try
                {
                    var input = TaskRecoveryInput.Read(task.State);
                    sources.AddRange(input.Documents.Select(d => d.Content));
                    sources.AddRange(input.Working.Select(m => m.Body));
                    if (input.Selected != null) sources.Add(input.Selected.Body);
                    if (input.Selection != null) sources.Add(input.Selection.Preview);
                }
                catch (ArgumentException) { /* A browser UI recovery record is a different DTO. */ }
            }
            var json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            foreach (var id in task.State.EvidenceIds)
            {
                var raw = task.Store.ReadEvidence(task.State.Id, id);
                if (!raw.StartsWith("{")) continue;
                var exchange = json.Deserialize<Dictionary<string, object>>(raw);
                object responseValue, resultsValue;
                if (!exchange.TryGetValue("response", out responseValue) || !exchange.TryGetValue("results", out resultsValue)) continue;
                var response = json.Deserialize<ChatCompletionResponseMessage>(json.Serialize(responseValue));
                var allowed = new HashSet<string>((response.tool_calls ?? new List<ChatToolCall>()).Where(c =>
                    (c.function.name.StartsWith("read_") && c.function.name != TaskContextManager.ReadEvidenceTool) ||
                    c.function.name == "fetch_web_page" || c.function.name == "search_mailbox").Select(c => c.id));
                foreach (Dictionary<string, object> result in (IEnumerable)resultsValue)
                {
                    object callId, content;
                    if (result.TryGetValue("ToolCallId", out callId) && allowed.Contains(Convert.ToString(callId)) && result.TryGetValue("Content", out content))
                    {
                        var text = Convert.ToString(content); sources.Add(text);
                        // Include decoded strings from JSON receipts so quoted source
                        // text is compared to the actual value, not its JSON escaping.
                        try { AddStrings(json.DeserializeObject(text), sources); } catch (ArgumentException) { }
                    }
                }
            }
            return string.Join("\n", sources);
        }
        private static void AddStrings(object value, List<string> result)
        {
            if (value is string) { result.Add((string)value); return; }
            var map = value as IDictionary<string, object>;
            if (map != null) { foreach (var item in map.Values) AddStrings(item, result); return; }
            var array = value as IEnumerable;
            if (array != null) foreach (var item in array) AddStrings(item, result);
        }
        public static object InspectPlan(string slidesJson)
        {
            var input = new JavaScriptSerializer { MaxJsonLength = int.MaxValue }.DeserializeObject(slidesJson);
            var slides = PresentationDraftWriter.ParseSlides(input);
            return PresentationDraftWriter.ComposeSamsung(slides).Select(p => new
            {
                layout = p.Source.Layout, background = p.Background,
                elements = p.Elements.Select(e => new { x = e.Box.X, y = e.Box.Y, width = e.Box.Width, height = e.Box.Height,
                    text = e.Text, font = e.Font, size = e.Size, minimum = e.Minimum, fill = e.Fill, color = e.Color,
                    hollow = e.Hollow, tableRows = e.Table?.Rows.Count ?? 0, table = e.Table == null ? null : new { headers = e.Table.Headers, rows = e.Table.Rows }, chart = e.Chart != null })
            }).ToArray();
        }
    }
}

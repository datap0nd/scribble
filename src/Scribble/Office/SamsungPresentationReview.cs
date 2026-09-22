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
        public const string AuthoringInstructions = SamsungAuthoringPolicy.Instructions;

        public static bool PrepareSampleEvidence(IDictionary<string, object> slide, string userInstruction)
        {
            // Only the trusted user's instruction can authorize synthetic content.
            // Email text and model-provided evidence cannot switch this mode on.
            if (!Regex.IsMatch(userInstruction ?? "", @"\b(sample|synthetic|illustrative|example)\s+(data|values|numbers)\b", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(userInstruction ?? "", @"\b(no|not|without|never)\b.{0,30}\b(sample|synthetic|illustrative|example)\b", RegexOptions.IgnoreCase)) return false;
            // A source-backed factual slide in a mixed deck cannot inherit sample mode.
            var kind = SamsungAuthoringPolicy.Text(slide, "content_kind");
            if (kind != "sample" && (kind.Length > 0 || slide.ContainsKey("source_spans") || slide.ContainsKey("evidence"))) return false;
            slide["content_kind"] = "sample";
            slide["evidence"] = userInstruction;
            slide["sources"] = "Sample data — supplied by the user; not actual business results";
            slide.Remove("source_spans");
            return true;
        }

        // Models often write the citation line into footnote ("Source: WB01
        // Ledger") and leave sources empty. Both render in the same visible
        // footer and reach the notes, so a footnote that is plainly a source
        // line is the citation; nothing is invented.
        public static void AdoptFootnoteCitation(IDictionary<string, object> slide)
        {
            object sources, footnote;
            if (slide.TryGetValue("sources", out sources) && !string.IsNullOrWhiteSpace(Convert.ToString(sources))) return;
            if (!slide.TryGetValue("footnote", out footnote)) return;
            var text = Convert.ToString(footnote) ?? "";
            if (!Regex.IsMatch(text, @"^\s*sources?\s*[:\u2014\u2013-]", RegexOptions.IgnoreCase)) return;
            slide["sources"] = text.Trim();
            slide.Remove("footnote");
        }

        public static void ValidateEvidence(string slideJson, string actualSource)
        {
            var json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            var data = json.Deserialize<Dictionary<string, object>>(slideJson);
            object raw;
            var evidence = data.TryGetValue("evidence", out raw) ? Convert.ToString(raw) : "";
            var layout = data.TryGetValue("layout", out raw) ? Convert.ToString(raw) : "";
            var special = new[] { "cover", "divider", "closing", "agenda" }.Contains(layout);
            if (!special && string.IsNullOrWhiteSpace(evidence)) throw new InvalidOperationException("SLIDE_EVIDENCE_REQUIRED: Cite source_spans returned by read_task_sources, or supply a verbatim supporting passage.");
            if (!string.IsNullOrWhiteSpace(evidence) && !NormalizeSource(actualSource).Contains(NormalizeSource(evidence)))
                throw new InvalidOperationException("SLIDE_EVIDENCE_UNVERIFIED: The excerpt does not occur in the original input or read receipts. Read the source and copy an exact passage.");
            // Metadata is not a numeric claim. Layout indices and outline levels
            // are not facts either. Inspect displayed content recursively.
            var content = string.Join(" ", data.Where(p => !new[] { "id", "sources", "evidence", "source_spans", "layout", "highlight_rows", "image_names", "purpose", "content_kind", "claims", "calculations", "annotations" }.Contains(p.Key)).SelectMany(p => DisplayedStrings(p.Value, p.Key)));
            if (content.Length > 36000) throw new InvalidOperationException("Slide content must be split into smaller review batches.");
            var allowed = new HashSet<string>(Numbers(special && string.IsNullOrWhiteSpace(evidence) ? actualSource : evidence));
            foreach (Match range in Regex.Matches(evidence ?? "", @"\b(?:weeks?|days?|months?|years?)\s+(\d+)\s*[-–]\s*(\d+)\b", RegexOptions.IgnoreCase))
            {
                int from, to;
                if (int.TryParse(range.Groups[1].Value, out from) && int.TryParse(range.Groups[2].Value, out to) && to >= from && to - (long)from <= 100)
                    for (var value = (long)from; value <= to; value++) allowed.Add(value.ToString(System.Globalization.CultureInfo.InvariantCulture));
            }
            foreach (var value in SamsungEvidence.ValidateCalculations(data, evidence))
            {
                allowed.Add(value);
                // "Revenue fell 2.95%" states the size of a host-computed -2.95.
                if (value.StartsWith("-", StringComparison.Ordinal)) allowed.Add(value.Substring(1));
            }
            SamsungEvidence.ValidateClaims(data, evidence);
            ValidateTableComparatives(data);
            // A period label (2026-05, June 2026) names a column rather than a
            // quantity; it must occur in the sources this task has read.
            var quantities = SamsungEvidence.RemoveVerifiedPeriodLabels(content, actualSource, evidence);
            var missing = Numbers(quantities).Where(n => !allowed.Contains(n)).Distinct().ToArray();
            if (missing.Length > 0) throw new InvalidOperationException("SLIDE_NUMBERS_UNVERIFIED: Values absent from cited evidence: " + string.Join(", ", missing));
            if (special) return;
            var explanatory = SamsungAuthoringPolicy.Text(data, "purpose") == "explanatory";
            if (!explanatory && (!data.TryGetValue("subtitle", out raw) || string.IsNullOrWhiteSpace(Convert.ToString(raw))))
                throw new InvalidOperationException("SLIDE_ACTION_TITLE_REQUIRED: An analytical slide needs a nonempty subtitle stating its evidence-backed finding. Add subtitle, or set purpose to explanatory for a definitions or setup slide.");
            if (!data.TryGetValue("sources", out raw) || string.IsNullOrWhiteSpace(Convert.ToString(raw)))
                throw new InvalidOperationException("SLIDE_CITATION_REQUIRED: Every factual slide needs a nonempty sources string, the visible citation line such as 'Source: WB01 Ledger; Scribble Draft audit'. A footnote is a separate qualifying note and does not replace sources. Add sources to this slide and to every other factual slide in the batch.");
        }
        private static string NormalizeSource(string value) { return Regex.Replace(value ?? "", @"\s+", " ").Trim(); }

        private static void ValidateTableComparatives(IDictionary<string, object> slide)
        {
            object raw;
            if (!slide.TryGetValue("table", out raw) || raw == null) return;
            var table = SamsungAuthoringPolicy.ReadMap(raw);
            var headers = SamsungAuthoringPolicy.Array(table, "headers").Select(item => Convert.ToString(item)).ToArray();
            if (headers.Length < 2) return;
            var rows = SamsungAuthoringPolicy.Array(table, "rows")
                .Select(item => (item as IEnumerable)?.Cast<object>().Select(cell => Convert.ToString(cell)).ToArray())
                .Where(row => row != null && row.Length == headers.Length && !string.IsNullOrWhiteSpace(row[0]) &&
                    !Regex.IsMatch(row[0], @"^\s*(?:all\s+groups|grand\s+total|total)\s*$", RegexOptions.IgnoreCase))
                .ToArray();
            if (rows.Length < 2) return;
            var assertions = string.Join("; ", new[] { "title", "subtitle", "takeaway", "message" }
                .Select(key => SamsungAuthoringPolicy.Text(slide, key)));
            foreach (var clause in Regex.Split(assertions, @"[;\r\n]"))
            foreach (var column in Enumerable.Range(1, headers.Length - 1))
            {
                var metric = Regex.Match(headers[column] ?? "", @"\b(?:revenue|cost|profit|margin|volume|units?)\b", RegexOptions.IgnoreCase);
                if (!metric.Success) continue;
                var values = new List<Tuple<string, decimal>>();
                foreach (var row in rows)
                {
                    decimal value;
                    if (!decimal.TryParse((row[column] ?? "").Trim().TrimEnd('%'), System.Globalization.NumberStyles.Number,
                        System.Globalization.CultureInfo.InvariantCulture, out value)) { values.Clear(); break; }
                    values.Add(Tuple.Create(row[0], value));
                }
                if (values.Count != rows.Length) continue;
                foreach (var item in values)
                {
                    var ranking = Regex.Match(clause,
                        @"\b" + Regex.Escape(item.Item1) + @"\b.{0,80}?\b(?<rank>highest|largest|lowest|smallest|leads|led)\b\s+(?:(?:on|in|for)\s+)?(?:\w+\s+){0,3}?" +
                        Regex.Escape(metric.Value) + @"\b", RegexOptions.IgnoreCase);
                    if (!ranking.Success) continue;
                    var highest = Regex.IsMatch(ranking.Groups["rank"].Value, @"^(?:highest|largest|leads|led)$", RegexOptions.IgnoreCase);
                    var extreme = highest ? values.Max(value => value.Item2) : values.Min(value => value.Item2);
                    if (item.Item2 != extreme)
                        throw new InvalidOperationException("SLIDE_TABLE_RANKING_FALSE: '" + item.Item1 + "' is not the " +
                            ranking.Groups["rank"].Value.ToLowerInvariant() + " " + metric.Value.ToLowerInvariant() +
                            " row in the supplied table. Correct the takeaway or the table.");
                }
            }
        }

        private static IEnumerable<string> DisplayedStrings(object value, string field)
        {
            if (value == null) yield break;
            if (value is string) { yield return (string)value; yield break; }
            var map = value as IDictionary<string, object>;
            if (map != null)
            {
                foreach (var pair in map) foreach (var text in DisplayedStrings(pair.Value, pair.Key)) yield return text;
                yield break;
            }
            var array = value as IEnumerable;
            if (array != null)
            {
                var index = 0;
                foreach (var item in array)
                {
                    index++;
                    foreach (var text in DisplayedStrings(item, ""))
                        yield return field == "bullets" || field == "points" ? Regex.Replace(text, @"^\s*" + index + @"[.)]\s+", "") : text;
                }
                yield break;
            }
            yield return Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture);
        }
        private static IEnumerable<string> Numbers(string text)
        {
            return Regex.Matches((text ?? "").Replace('\u2212', '-'), @"(?<![A-Za-z0-9])[-+]?(?:\d+(?:[,.]\d+)*|\.\d+)(?:[eE][-+]?\d+)?%?").Cast<Match>().Select(m =>
            {
                var raw = m.Value.Replace(",", "").TrimStart('+').TrimEnd('%');
                double value;
                return double.TryParse(raw, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out value)
                    ? value.ToString("R", System.Globalization.CultureInfo.InvariantCulture) : raw;
            });
        }

        public static void ValidatePlan(string[] plan, string[] batch, string[] completed)
        {
            if (plan == null || plan.Length == 0 || plan.Any(string.IsNullOrWhiteSpace) || plan.Any(p => p.Length > 80) || plan.Distinct().Count() != plan.Length)
                throw new InvalidOperationException("SLIDE_PLAN_REQUIRED: Provide ordered unique IDs for the complete storyline.");
            var outstanding = plan.Where(id => !completed.Contains(id)).ToArray();
            var expected = " Outstanding planned IDs, in order: " + string.Join(", ", outstanding) +
                ". A batch is the first one or more of these, in this order, each with complete slide content.";
            if (batch.Length == 0 || batch.Any(id => !plan.Contains(id) || completed.Contains(id)) || batch.Distinct().Count() != batch.Length)
                throw new InvalidOperationException("SLIDE_PLAN_MISMATCH: Each batch must contain unique outstanding IDs from the original plan." + expected);
            var pending = outstanding.Take(batch.Length).ToArray();
            if (!pending.SequenceEqual(batch)) throw new InvalidOperationException("SLIDE_PLAN_ORDER: Complete the next planned slides in storyline order. This batch began with '" + batch[0] + "'." + expected);
        }

        internal static string SourceCorpus(TaskContextManager task, string prompt)
        {
            var sources = new List<string> { prompt ?? "" };
            if (task == null) return string.Join("\n", sources);
            sources.Add(task.State.Objective);
            sources.AddRange(task.State.OriginalDecisions);
            task.Sources.CaptureInput();
            sources.Add(task.Sources.Resolve(task.Sources.Spans().Select(s => s.Id)));
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
            string browserSource;
            if (task.State.HostData.TryGetValue("browser_source_text", out browserSource)) sources.Add(browserSource);
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
                    c.function.name == "fetch_web_page" || c.function.name == "search_mailbox" ||
                    c.function.name == BrowserToolCatalog.ReadPage || c.function.name == BrowserToolCatalog.SnapshotPage ||
                    c.function.name == BrowserToolCatalog.RecordEvidence).Select(c => c.id));
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

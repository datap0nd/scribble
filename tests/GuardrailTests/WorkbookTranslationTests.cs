using System;
using System.Collections.Generic;
using System.Linq;
using Scribble.Chat;
using Scribble.Office;

namespace GuardrailTests
{
    // Whole-workbook translation runs in both directions through the same
    // snapshot-bound pipeline; the destination language selects discovery.
    internal static class WorkbookTranslationTests
    {
        private static void Check(bool value, string message) { if (!value) throw new Exception(message); }
        private static void Reject(Action action)
        { try { action(); } catch (InvalidOperationException) { return; } throw new Exception("Expected rejection."); }

        internal static void DirectionRouting()
        {
            Func<string, string> target = ExcelSelectionOutputPolicy.WholeWorkbookTranslationTarget;
            Check(target("Translate this whole workbook from English to Korean") == ExcelSelectionOutputPolicy.TargetKorean,
                "An English-to-Korean workbook request was not routed to the Korean target.");
            Check(target("please translate every sheet into Korean") == ExcelSelectionOutputPolicy.TargetKorean,
                "A Korean destination without naming English was not recognized.");
            Check(target("이 통합 문서 전체를 한국어로 번역해 줘") == ExcelSelectionOutputPolicy.TargetKorean,
                "A Korean-language request for a Korean destination was not recognized.");
            Check(target("Translate every Korean text cell in every worksheet of the active workbook into English.") == ExcelSelectionOutputPolicy.TargetEnglish,
                "The established Korean-to-English request changed meaning.");
            Check(target("translate Korean English workbook") == ExcelSelectionOutputPolicy.TargetEnglish,
                "An unmarked language-pair request lost its original Korean-to-English meaning.");
            Check(target("Translate the workbook from Korean to English") == ExcelSelectionOutputPolicy.TargetEnglish,
                "A from/to pair in the English direction was misrouted.");
            Check(target("Translate column B to Korean") == null, "A request without whole-workbook scope was claimed.");
            Check(target("Summarize the Korean workbook") == null, "A non-translation request was claimed.");
            Check(ExcelSelectionOutputPolicy.IsWholeWorkbookEnglishToKoreanRequest("translate all sheets to Korean") &&
                !ExcelSelectionOutputPolicy.IsWholeWorkbookKoreanToEnglishRequest("translate all sheets to Korean"),
                "The two direction predicates overlap.");
        }

        internal static void EnglishCellEligibility()
        {
            foreach (var text in new[] { "Due date", "TOTAL", "In progress", "Q3 revenue target", "Review required (urgent)" })
                Check(ExcelSelectionOutputPolicy.IsTranslatableEnglishText(text), "English text was not eligible: " + text);
            foreach (var text in new[] { "", "  ", "12345", "2026-06", "SKU-1042", "A1B2", "https://example.com/report", "ops@example.com", "마감일", "Due 마감", "%", "A" })
                Check(!ExcelSelectionOutputPolicy.IsTranslatableEnglishText(text), "A non-translatable value was eligible: " + text);
        }

        internal static void KoreanTargetSession()
        {
            var sources = new[] { "Due date", "Complete", "In progress", "Notes", "Owner", "Samsung" };
            var session = new KoreanWorkbookOutputSession("h", sources.Length, ExcelSelectionOutputPolicy.TargetKorean, sources);
            // An echoed window is refused once with actionable guidance.
            try { session.Stage("h", 0, sources, true); throw new Exception("An untranslated window was accepted."); }
            catch (InvalidOperationException ex) { Check(ex.Message.Contains("nothing was translated"), "The untranslated-window refusal is not actionable."); }
            Check(session.StagedCount == 0, "A refused window was staged.");
            var korean = new[] { "마감일", "완료", "진행 중", "비고", "담당자", "Samsung" };
            Check(session.Stage("h", 0, korean, true) && session.StagedCount == 6, "Korean output with a retained brand name was rejected.");

            // Windows that really are only names may be resubmitted unchanged.
            var names = new[] { "Samsung", "Galaxy", "Bixby", "SmartThings", "Knox" };
            var nameSession = new KoreanWorkbookOutputSession("n", names.Length, ExcelSelectionOutputPolicy.TargetKorean, names);
            Reject(() => nameSession.Stage("n", 0, names, true));
            Check(nameSession.Stage("n", 0, names, true), "A deliberate unchanged resubmission deadlocked the translation.");

            var empty = new KoreanWorkbookOutputSession("e", 1, ExcelSelectionOutputPolicy.TargetKorean, new[] { "Notes" });
            Reject(() => empty.Stage("e", 0, new[] { "" }, true));

            // The original direction still refuses Hangul in its English output.
            var english = new KoreanWorkbookOutputSession("k", 1);
            Reject(() => english.Stage("k", 0, new[] { "마감일" }, true));
            Check(english.Stage("k", 0, new[] { "Due date" }, true), "Korean-to-English output was rejected.");
        }

        internal static void RequestSurface()
        {
            var context = new List<ExternalContextDocument> { new ExternalContextDocument("Detected English workbook cells", "Workbook handle: korean_workbook_h1") };
            var request = DocumentChatRequestFactory.Create("test-model", "excel", "Workbook: Book1", new List<ChatTurn>(),
                "Translate the whole workbook from English to Korean", true, context, null, null, false, true, ExcelSelectionOutputPolicy.TargetKorean);
            var names = request.tools.Select(item => item.function.name).ToArray();
            var system = Convert.ToString(((ChatCompletionInputMessage)request.messages[0]).content);
            var tool = request.tools.Single(item => item.function.name == WorkbookToolCatalog.WriteKoreanTranslations);
            Check(names.Contains("write_korean_translations") && !names.Contains("write_selection_output"), "The Korean-target request lost its snapshot-bound write tool.");
            Check(system.Contains("literal English text cell") && system.Contains("one Korean value per source entry") &&
                system.Contains("never saved") && !system.Contains("into English"), "The Korean-target instruction still describes the English direction.");
            Check(tool.function.description.Contains("detected English text cells") && tool.function.description.Contains("one Korean string"),
                "The tool description does not name the Korean target.");

            var legacy = DocumentChatRequestFactory.Create("test-model", "excel", "Workbook: Book1", new List<ChatTurn>(),
                "translate Korean text throughout the workbook", true, context, null, null, false, true);
            var legacySystem = Convert.ToString(((ChatCompletionInputMessage)legacy.messages[0]).content);
            Check(legacySystem.Contains("replacing exactly the detected") && legacySystem.Contains("into English"), "The Korean-to-English instruction changed.");

            var restored = new SavedKoreanWorkbook { WorkbookName = "Book1", TargetLanguage = ExcelSelectionOutputPolicy.TargetKorean }.Restore();
            Check(restored.TranslatesToKorean && restored.SourceLanguage == "English", "A resumed task lost its translation direction.");
            Check(!new SavedKoreanWorkbook { WorkbookName = "Book1" }.Restore().TranslatesToKorean, "An older checkpoint changed direction.");
        }
    }
}

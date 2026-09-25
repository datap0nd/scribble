using System;
using System.Collections.Generic;
using System.IO;
using Scribble.Chat;
using Scribble.Office;

namespace GuardrailTests
{
    internal static class PresentationAcceptanceTests
    {
        internal static void ScopedReceiptCannotCertifyCharts()
        {
            var receipt = new Dictionary<string, object>
            {
                { "assembly_sha256", "candidate" },
                { "policy", SamsungAuthoringPolicy.Version },
                { "execution_kind", "native" },
                { "scope", PresentationRevisionAcceptance.ChartlessScope },
                { "revision_passed", true }, { "preservation_passed", true },
                { "rollback_passed", true }, { "chartless_operations_passed", true },
                { "all_operations_passed", false }
            };
            Check(PresentationRevisionAcceptance.Supports(receipt, "candidate", false));
            Check(!PresentationRevisionAcceptance.Supports(receipt, "candidate", true));
            Check(!PresentationRevisionAcceptance.Supports(receipt, "other", false));
            receipt["all_operations_passed"] = true;
            Check(!PresentationRevisionAcceptance.Supports(receipt, "candidate", false));
            receipt.Remove("scope");
            Check(PresentationRevisionAcceptance.Supports(receipt, "candidate", true));
            receipt["powerpoint_exited"] = true;
            Check(!PresentationRevisionAcceptance.Supports(receipt, "candidate", true));
            receipt["powerpoint_exited"] = false;
            receipt["rollback_passed"] = false;
            Check(!PresentationRevisionAcceptance.Supports(receipt, "candidate", false));
            Check(PresentationRevisionAcceptance.ContainsChartOperation(new object[] {
                new Dictionary<string, object> { { "kind", "replace_slide" },
                    { "slide", new Dictionary<string, object> { { "chart", new object() } } } }
            }));
            Check(PresentationRevisionAcceptance.ContainsChartOperation(
                new Dictionary<string, object> { { "kind", "chart_point" } }));
            Check(PresentationRevisionAcceptance.ContainsChartOperation(
                new Dictionary<string, object> { { "secondary_chart", new object() } }));
            Check(!PresentationRevisionAcceptance.ContainsChartOperation(
                new Dictionary<string, object> { { "kind", "replace_text" }, { "text", "chart" } }));
        }

        internal static void TransportBudgetSurvivesRestart()
        {
            var root = Path.Combine(Path.GetTempPath(), "scribble-delivery-budget-" + Guid.NewGuid().ToString("N"));
            var store = new TaskCheckpointStore(root);
            var state = new DurableTaskState { Host = "powerpoint", Objective = "Offline budget test" };
            state.HostData["delivery_request_limit"] = "18";
            var first = new TaskDiagnostics(store, state);
            for (var index = 0; index < 17; index++) first.ReserveModelRequest();
            var resumedState = store.Load(state.Id);
            var resumed = new TaskDiagnostics(store, resumedState);
            resumed.ReserveModelRequest();
            var rejected = false;
            try { resumed.ReserveModelRequest(); }
            catch (InvalidOperationException error) { rejected = error.Message.StartsWith("DELIVERY_REQUEST_LIMIT:", StringComparison.Ordinal); }
            Check(rejected && store.Load(state.Id).HostData["delivery_request_count"] == "18");
            resumedState.HostData["delivery_request_count"] = "invalid";
            rejected = false;
            try { resumed.ReserveModelRequest(); }
            catch (InvalidOperationException error) { rejected = error.Message == "DELIVERY_REQUEST_BUDGET_INVALID"; }
            Check(rejected);
            // Keep the tiny encrypted checkpoint as local failure-replay evidence.
        }
        private static void Check(bool value)
        { if (!value) throw new InvalidOperationException("Native acceptance scope escaped its evidence."); }
    }
}

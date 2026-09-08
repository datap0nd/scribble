using System.Collections.Generic;

namespace Scribble.Chat
{
    internal static class SamsungWorkflowSchema
    {
        internal static object List(object item, int max = 24)
        { return new Dictionary<string, object> { { "type", "array" }, { "items", item }, { "maxItems", max } }; }
        internal static object Strings(int max = 1000) { return List(ToolSchema.String("Exact reference or required content."), max); }
        internal static object Briefs()
        { return List(ToolSchema.Build(new Dictionary<string, object> {
            { "id", ToolSchema.String("Stable ID in plan order.") }, { "purpose", ToolSchema.String("Purpose and business question this slide answers.") },
            { "message", ToolSchema.String("Main message, grounded in the source.") }, { "layout", ToolSchema.String("Samsung recipe name.") },
            { "source_spans", Strings() }, { "required_content", Strings() }
        }, "id", "purpose", "message", "layout", "source_spans", "required_content"), 1000); }
        internal static object Claims()
        { return List(ToolSchema.Build(new Dictionary<string, object> {
            { "text", ToolSchema.String("Displayed claim.") }, { "label", ToolSchema.String("Associated metric/category.") },
            { "unit", ToolSchema.String("Unit or not applicable.") }, { "period", ToolSchema.String("Reporting period or not applicable.") },
            { "evidence", ToolSchema.String("Verbatim supporting passage from this slide's sources.") }
        }, "text", "label", "unit", "period", "evidence")); }
        internal static object Calculations()
        {
            var operand = ToolSchema.Build(new Dictionary<string, object> {
                { "value", new { type = "number" } }, { "label", ToolSchema.String("Source label.") },
                { "unit", ToolSchema.String("Source unit.") }, { "period", ToolSchema.String("Source period.") },
                { "evidence", ToolSchema.String("Verbatim passage containing value, label, unit and period.") }
            }, "value", "label", "unit", "period", "evidence");
            return List(ToolSchema.Build(new Dictionary<string, object> {
                { "label", ToolSchema.String("Displayed derived metric.") },
                { "operation", new { type = "string", @enum = new[] { "sum", "difference", "ratio", "percent", "growth_percent" } } },
                { "operands", List(operand) }, { "result", new { type = "number" } },
                { "unit", ToolSchema.String("Result unit; percentage operations use %.") },
                { "decimals", ToolSchema.Integer("Round half away from zero.", 0, 6) }
            }, "label", "operation", "operands", "result", "unit", "decimals"));
        }
        internal static object Annotations()
        { return List(ToolSchema.Build(new Dictionary<string, object> {
            { "target", new { type = "string", @enum = new[] { "table", "secondary_table", "chart", "secondary_chart" } } },
            { "row", ToolSchema.Integer("1-based data row/category.", 1, 200) },
            { "column", ToolSchema.Integer("Optional 1-based table column.", 1, 8) },
            { "series", ToolSchema.Integer("Optional 1-based chart series.", 1, 5) }
        }, "target", "row")); }
    }
}

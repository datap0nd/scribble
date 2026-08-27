using System;

namespace Scribble.Configuration
{
    public static class ModelSelectionPolicy
    {
        public static bool IsGenerativeModel(string model)
        {
            var value = (model ?? string.Empty).Trim();
            return value.Length > 0 &&
                   !ModelCatalog.IsDisallowedModel(value) &&
                   value.IndexOf(
                       "embedding",
                       StringComparison.OrdinalIgnoreCase) < 0;
        }

        public static string DescriptionFor(string model)
        {
            return ModelCatalog.DescribeForSelection(model);
        }
    }
}

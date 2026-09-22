using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace Scribble.Office
{
    public static class AnalysisContract
    {
        public const int Version = 1;
        public const string SourceObserved = "observed_source";
        public const string DerivedCalculation = "derived_calculation";
        public const string UnresolvedAssertion = "unresolved_assertion";
        public const string DecimalValue = "decimal";
        public const string IntegerValue = "integer";
        public const string DateValue = "date";
        public const string TextValue = "text";
        public const string BooleanValue = "boolean";
        public const string MissingValue = "missing";
        public const string ErrorValue = "error";
        public const string Verified = "verified";
        public const string Unresolved = "unresolved";

        public static SourceSnapshot CreateSnapshot(
            string sourceInstanceId,
            string sourceType,
            string captureRevision,
            string coverage,
            string calculationState,
            IEnumerable<SourceLocator> locators,
            IEnumerable<TableDataset> tables)
        {
            Require(sourceInstanceId, "source instance ID");
            Require(sourceType, "source type");
            Require(captureRevision, "capture revision");
            var locatorList = (locators ?? new SourceLocator[0]).ToList();
            var tableList = (tables ?? new TableDataset[0]).ToList();
            ValidateLocators(locatorList);
            ValidateTables(tableList);
            var contentHash = Hash(CanonicalTables(tableList));
            return new SourceSnapshot
            {
                ContractVersion = Version,
                SourceInstanceId = sourceInstanceId.Trim(),
                SourceType = sourceType.Trim(),
                CaptureRevision = captureRevision.Trim(),
                CapturedUtc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
                ContentHash = contentHash,
                SnapshotId = HostId("snapshot", sourceInstanceId.Trim() + "\n" +
                    captureRevision.Trim() + "\n" + contentHash),
                Coverage = (coverage ?? string.Empty).Trim(),
                CalculationState = (calculationState ?? string.Empty).Trim(),
                Locators = locatorList,
                Tables = tableList
            };
        }

        public static bool IsCurrent(
            SourceSnapshot snapshot,
            string sourceInstanceId,
            string captureRevision,
            string contentHash)
        {
            return snapshot != null && snapshot.ContractVersion == Version &&
                string.Equals(snapshot.SourceInstanceId, sourceInstanceId,
                    StringComparison.Ordinal) &&
                string.Equals(snapshot.CaptureRevision, captureRevision,
                    StringComparison.Ordinal) &&
                string.Equals(snapshot.ContentHash, contentHash,
                    StringComparison.OrdinalIgnoreCase) &&
                string.Equals(snapshot.SnapshotId,
                    HostId("snapshot", sourceInstanceId + "\n" +
                        captureRevision + "\n" + contentHash),
                    StringComparison.Ordinal);
        }

        public static VerifiedFact CreateObservedFact(
            string snapshotId,
            string metric,
            string valueType,
            string value,
            string displayText,
            string unit,
            string currency,
            string period,
            IDictionary<string, string> dimensions,
            IEnumerable<SourceLocator> locators,
            string status)
        {
            Require(snapshotId, "snapshot ID");
            Require(metric, "metric");
            Require(valueType, "value type");
            ValidateValue(valueType, value, status);
            var locatorList = (locators ?? new SourceLocator[0]).ToList();
            ValidateLocators(locatorList);
            var fact = new VerifiedFact
            {
                SnapshotId = snapshotId,
                Kind = SourceObserved,
                ValueType = valueType,
                Value = value ?? string.Empty,
                DisplayText = displayText ?? string.Empty,
                Metric = metric.Trim(),
                Unit = (unit ?? string.Empty).Trim(),
                Currency = (currency ?? string.Empty).Trim(),
                Period = (period ?? string.Empty).Trim(),
                Dimensions = dimensions == null
                    ? new Dictionary<string, string>()
                    : new Dictionary<string, string>(dimensions,
                        StringComparer.Ordinal),
                Locators = locatorList,
                Status = string.IsNullOrWhiteSpace(status)
                    ? Unresolved
                    : status.Trim()
            };
            fact.FactId = ExpectedFactId(fact);
            return fact;
        }

        public static AnalysisArtifact CreateArtifact(
            IEnumerable<SourceSnapshot> snapshots,
            IEnumerable<VerifiedFact> facts,
            IEnumerable<AnalysisCalculation> calculations,
            IEnumerable<string> assumptions,
            IEnumerable<string> unresolvedConflicts)
        {
            var artifact = new AnalysisArtifact
            {
                ContractVersion = Version,
                Snapshots = (snapshots ?? new SourceSnapshot[0]).ToList(),
                Facts = (facts ?? new VerifiedFact[0]).ToList(),
                Calculations = (calculations ?? new AnalysisCalculation[0]).ToList(),
                Assumptions = (assumptions ?? new string[0]).ToList(),
                UnresolvedConflicts = (unresolvedConflicts ?? new string[0]).ToList()
            };
            Validate(artifact);
            artifact.AnalysisId = HostId("analysis", CanonicalArtifact(artifact));
            return artifact;
        }

        public static string Serialize(AnalysisArtifact artifact)
        {
            Validate(artifact);
            var expected = HostId("analysis", CanonicalArtifact(artifact));
            if (!string.Equals(artifact.AnalysisId, expected,
                StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "ANALYSIS_ID_MISMATCH: The analysis content changed after its host-issued ID was created.");
            return new JavaScriptSerializer
            {
                MaxJsonLength = int.MaxValue
            }.Serialize(artifact);
        }

        public static AnalysisArtifact Deserialize(string json)
        {
            if (string.IsNullOrWhiteSpace(json) || json.Length > 8 * 1024 * 1024)
                throw new InvalidOperationException(
                    "ANALYSIS_PAYLOAD_INVALID: The analysis payload is missing or oversized.");
            var artifact = new JavaScriptSerializer
            {
                MaxJsonLength = int.MaxValue
            }.Deserialize<AnalysisArtifact>(json);
            Validate(artifact);
            var expected = HostId("analysis", CanonicalArtifact(artifact));
            if (!string.Equals(artifact.AnalysisId, expected,
                StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "ANALYSIS_ID_MISMATCH: The persisted analysis content does not match its host-issued ID.");
            return artifact;
        }

        public static void Validate(AnalysisArtifact artifact)
        {
            if (artifact == null || artifact.ContractVersion != Version)
                throw new InvalidOperationException(
                    "ANALYSIS_VERSION_UNSUPPORTED: The analysis contract version is not supported.");
            artifact.Snapshots = artifact.Snapshots ?? new List<SourceSnapshot>();
            artifact.Facts = artifact.Facts ?? new List<VerifiedFact>();
            artifact.Calculations = artifact.Calculations ??
                new List<AnalysisCalculation>();
            artifact.Assumptions = artifact.Assumptions ?? new List<string>();
            artifact.UnresolvedConflicts = artifact.UnresolvedConflicts ??
                new List<string>();
            var snapshotIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (var snapshot in artifact.Snapshots)
            {
                ValidateSnapshot(snapshot);
                if (!snapshotIds.Add(snapshot.SnapshotId))
                    throw new InvalidOperationException(
                        "ANALYSIS_SNAPSHOT_DUPLICATE: Snapshot IDs must be unique.");
            }
            var factIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (var fact in artifact.Facts)
            {
                ValidateFact(fact);
                if (!snapshotIds.Contains(fact.SnapshotId) ||
                    !factIds.Add(fact.FactId))
                    throw new InvalidOperationException(
                        "ANALYSIS_FACT_BINDING_INVALID: Facts need a unique ID and a retained source snapshot.");
            }
            var calculationIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (var calculation in artifact.Calculations)
            {
                if (calculation == null ||
                    string.IsNullOrWhiteSpace(calculation.CalculationId) ||
                    !calculationIds.Add(calculation.CalculationId) ||
                    calculation.InputFactIds == null ||
                    calculation.InputFactIds.Any(id => !factIds.Contains(id)) ||
                    !factIds.Contains(calculation.OutputFactId))
                    throw new InvalidOperationException(
                        "ANALYSIS_CALCULATION_BINDING_INVALID: A calculation references a missing or duplicate fact.");
            }
        }

        internal static string ExpectedFactId(VerifiedFact fact)
        {
            return HostId("fact", CanonicalFact(fact));
        }

        internal static string HostId(string prefix, string canonical)
        {
            return prefix + "_" + Hash(canonical).Substring(0, 24);
        }

        internal static string Format(decimal value)
        {
            return value.ToString("0.#############################",
                CultureInfo.InvariantCulture);
        }

        internal static decimal Decimal(VerifiedFact fact)
        {
            decimal value;
            if (fact == null || fact.Status != Verified ||
                (fact.ValueType != DecimalValue &&
                 fact.ValueType != IntegerValue) ||
                !decimal.TryParse(fact.Value, NumberStyles.Float,
                    CultureInfo.InvariantCulture, out value))
                throw new InvalidOperationException(
                    "ANALYSIS_INPUT_NOT_VERIFIED_NUMERIC: Calculations require verified decimal or integer facts.");
            return value;
        }

        private static void ValidateSnapshot(SourceSnapshot snapshot)
        {
            if (snapshot == null || snapshot.ContractVersion != Version)
                throw new InvalidOperationException(
                    "ANALYSIS_SNAPSHOT_INVALID: The source snapshot is missing or has an unsupported version.");
            Require(snapshot.SourceInstanceId, "source instance ID");
            Require(snapshot.CaptureRevision, "capture revision");
            ValidateLocators(snapshot.Locators);
            ValidateTables(snapshot.Tables);
            var expectedHash = Hash(CanonicalTables(snapshot.Tables));
            if (!string.Equals(snapshot.ContentHash, expectedHash,
                    StringComparison.OrdinalIgnoreCase) ||
                !IsCurrent(snapshot, snapshot.SourceInstanceId,
                    snapshot.CaptureRevision, expectedHash))
                throw new InvalidOperationException(
                    "ANALYSIS_SNAPSHOT_HASH_MISMATCH: The snapshot content or binding changed after capture.");
        }

        private static void ValidateFact(VerifiedFact fact)
        {
            if (fact == null) throw new InvalidOperationException(
                "ANALYSIS_FACT_INVALID: A fact is missing.");
            ValidateValue(fact.ValueType, fact.Value, fact.Status);
            ValidateLocators(fact.Locators);
            if (!string.Equals(fact.FactId, ExpectedFactId(fact),
                StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "ANALYSIS_FACT_ID_MISMATCH: A fact changed after its host-issued ID was created.");
        }

        private static void ValidateValue(
            string valueType,
            string value,
            string status)
        {
            var known = new[] { DecimalValue, IntegerValue, DateValue,
                TextValue, BooleanValue, MissingValue, ErrorValue };
            if (!known.Contains(valueType, StringComparer.Ordinal))
                throw new InvalidOperationException(
                    "ANALYSIS_VALUE_TYPE_INVALID: The fact value type is not supported.");
            if ((valueType == MissingValue || valueType == ErrorValue) &&
                !string.IsNullOrEmpty(value))
                throw new InvalidOperationException(
                    "ANALYSIS_EMPTY_VALUE_REQUIRED: Missing and error facts cannot carry a numeric replacement.");
            decimal parsed;
            if ((valueType == DecimalValue || valueType == IntegerValue) &&
                !decimal.TryParse(value, NumberStyles.Float,
                    CultureInfo.InvariantCulture, out parsed))
                throw new InvalidOperationException(
                    "ANALYSIS_NUMERIC_VALUE_INVALID: Numeric facts require invariant decimal text.");
            if (status == Verified && valueType == ErrorValue)
                throw new InvalidOperationException(
                    "ANALYSIS_ERROR_NOT_VERIFIED: An error cell cannot be promoted to a verified value.");
        }

        private static void ValidateLocators(IEnumerable<SourceLocator> locators)
        {
            foreach (var locator in locators ?? new SourceLocator[0])
            {
                if (locator == null || string.IsNullOrWhiteSpace(locator.Kind) ||
                    string.IsNullOrWhiteSpace(locator.SourceInstanceId))
                    throw new InvalidOperationException(
                        "ANALYSIS_LOCATOR_INVALID: Source locators require a kind and source instance.");
            }
        }

        private static void ValidateTables(IEnumerable<TableDataset> tables)
        {
            var tableIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (var table in tables ?? new TableDataset[0])
            {
                if (table == null || string.IsNullOrWhiteSpace(table.TableId) ||
                    !tableIds.Add(table.TableId) || table.Rows < 0 ||
                    table.Columns < 0 || table.Cells == null)
                    throw new InvalidOperationException(
                        "ANALYSIS_TABLE_INVALID: Tables need a unique ID, dimensions, and cells.");
                var positions = new HashSet<string>(StringComparer.Ordinal);
                foreach (var cell in table.Cells)
                {
                    if (cell == null || cell.Row < 0 || cell.Column < 0 ||
                        cell.Row >= table.Rows || cell.Column >= table.Columns ||
                        !positions.Add(cell.Row.ToString(CultureInfo.InvariantCulture) +
                            ":" + cell.Column.ToString(CultureInfo.InvariantCulture)))
                        throw new InvalidOperationException(
                            "ANALYSIS_CELL_INVALID: Table cells need a unique in-bounds position.");
                    ValidateValue(cell.ValueType, cell.Value, cell.Status);
                }
            }
        }

        private static string CanonicalArtifact(AnalysisArtifact artifact)
        {
            return string.Join("\n", artifact.Snapshots.OrderBy(item =>
                    item.SnapshotId, StringComparer.Ordinal).Select(item => item.SnapshotId)) +
                "\n--facts--\n" + string.Join("\n", artifact.Facts.OrderBy(item =>
                    item.FactId, StringComparer.Ordinal).Select(CanonicalFact)) +
                "\n--calculations--\n" + string.Join("\n", artifact.Calculations
                    .OrderBy(item => item.CalculationId, StringComparer.Ordinal)
                    .Select(item => item.CalculationId + "|" + item.Operation + "|" +
                        string.Join(",", item.InputFactIds ?? new List<string>()) + "|" +
                        item.OutputFactId + "|" + item.NullPolicy + "|" +
                        item.RoundingPolicy)) +
                "\n--assumptions--\n" + string.Join("\n", artifact.Assumptions) +
                "\n--conflicts--\n" + string.Join("\n", artifact.UnresolvedConflicts);
        }

        private static string CanonicalFact(VerifiedFact fact)
        {
            var dimensions = fact.Dimensions ?? new Dictionary<string, string>();
            var locators = fact.Locators ?? new List<SourceLocator>();
            return string.Join("|", new[] { fact.SnapshotId, fact.Kind,
                fact.ValueType, fact.Value, fact.Metric, fact.Unit,
                fact.Currency, fact.Period, fact.Status,
                string.Join(",", dimensions.OrderBy(item => item.Key,
                    StringComparer.Ordinal).Select(item => item.Key + "=" + item.Value)),
                string.Join(",", locators.Select(CanonicalLocator)) });
        }

        private static string CanonicalTables(IEnumerable<TableDataset> tables)
        {
            return string.Join("\n", (tables ?? new TableDataset[0])
                .OrderBy(table => table.TableId, StringComparer.Ordinal)
                .Select(table => table.TableId + "|" + table.Name + "|" +
                    table.Rows.ToString(CultureInfo.InvariantCulture) + "|" +
                    table.Columns.ToString(CultureInfo.InvariantCulture) + "|" +
                    string.Join(";", table.Cells.OrderBy(cell => cell.Row)
                        .ThenBy(cell => cell.Column).Select(cell =>
                            cell.Row.ToString(CultureInfo.InvariantCulture) + "," +
                            cell.Column.ToString(CultureInfo.InvariantCulture) + "," +
                            cell.Reference + "," + cell.ValueType + "," +
                            cell.RawValue + "," + cell.Value + "," +
                            cell.DisplayText + "," +
                            cell.Formula + "," + cell.NumberFormat + "," +
                            cell.Status))));
        }

        private static string CanonicalLocator(SourceLocator locator)
        {
            return string.Join("/", new[] { locator.Kind,
                locator.SourceInstanceId, locator.WorksheetIdentity,
                locator.Range, locator.Cell, locator.PresentationIdentity,
                locator.SlideId, locator.ShapeId });
        }

        private static string Hash(string value)
        {
            using (var hash = SHA256.Create())
            {
                return string.Concat(hash.ComputeHash(Encoding.UTF8.GetBytes(
                    value ?? string.Empty)).Select(item => item.ToString("x2",
                        CultureInfo.InvariantCulture)));
            }
        }

        private static void Require(string value, string name)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new InvalidOperationException(
                    "ANALYSIS_REQUIRED: Missing " + name + ".");
        }
    }

    public static class AnalysisCalculator
    {
        public const string Sum = "sum";
        public const string Difference = "difference";
        public const string Ratio = "ratio";
        public const string Growth = "growth";
        public const string Margin = "margin";
        public const string Count = "count";

        public static AnalysisCalculation Calculate(
            string operation,
            IEnumerable<VerifiedFact> inputs,
            string outputMetric,
            string outputPeriod,
            out VerifiedFact output)
        {
            var facts = (inputs ?? new VerifiedFact[0]).ToList();
            if (facts.Count == 0)
                throw new InvalidOperationException(
                    "ANALYSIS_CALCULATION_EMPTY: A calculation needs input facts.");
            var values = facts.Select(AnalysisContract.Decimal).ToArray();
            decimal result;
            string unit;
            string currency;
            if (operation == Sum)
            {
                SameUnit(facts);
                result = values.Aggregate(0m, (total, value) =>
                    checked(total + value));
                unit = facts[0].Unit;
                currency = facts[0].Currency;
            }
            else if (operation == Difference)
            {
                RequireCount(facts, 2, operation);
                SameUnit(facts);
                result = checked(values[0] - values[1]);
                unit = facts[0].Unit;
                currency = facts[0].Currency;
            }
            else if (operation == Ratio)
            {
                RequireCount(facts, 2, operation);
                SameUnit(facts);
                result = Divide(values[0], values[1]);
                unit = "ratio";
                currency = string.Empty;
            }
            else if (operation == Growth)
            {
                RequireCount(facts, 2, operation);
                SameUnit(facts);
                result = Divide(checked(values[0] - values[1]), values[1]);
                unit = "ratio";
                currency = string.Empty;
            }
            else if (operation == Margin)
            {
                RequireCount(facts, 2, operation);
                SameUnit(facts);
                result = Divide(checked(values[0] - values[1]), values[0]);
                unit = "ratio";
                currency = string.Empty;
            }
            else if (operation == Count)
            {
                result = facts.Count;
                unit = "count";
                currency = string.Empty;
            }
            else
            {
                throw new InvalidOperationException(
                    "ANALYSIS_OPERATION_UNSUPPORTED: " + operation);
            }

            var snapshotId = facts[0].SnapshotId;
            if (facts.Any(fact => fact.SnapshotId != snapshotId))
                throw new InvalidOperationException(
                    "ANALYSIS_SNAPSHOT_CONFLICT: Inputs from different snapshots need an explicit cross-snapshot binding.");
            var calculationId = AnalysisContract.HostId("calculation",
                operation + "|" + string.Join(",", facts.Select(fact =>
                    fact.FactId)) + "|" + outputMetric + "|" + outputPeriod);
            output = AnalysisContract.CreateObservedFact(snapshotId,
                outputMetric,
                operation == Count
                    ? AnalysisContract.IntegerValue
                    : AnalysisContract.DecimalValue,
                AnalysisContract.Format(result),
                AnalysisContract.Format(result),
                unit,
                currency,
                outputPeriod,
                new Dictionary<string, string>(),
                facts.SelectMany(fact => fact.Locators ??
                    new List<SourceLocator>()).ToList(),
                AnalysisContract.Verified);
            output.Kind = AnalysisContract.DerivedCalculation;
            output.FactId = AnalysisContract.ExpectedFactId(output);
            return new AnalysisCalculation
            {
                CalculationId = calculationId,
                Operation = operation,
                InputFactIds = facts.Select(fact => fact.FactId).ToList(),
                OutputFactId = output.FactId,
                NullPolicy = "reject_missing_or_error",
                RoundingPolicy = "preserve_source_precision",
                Status = AnalysisContract.Verified
            };
        }

        public static IList<RankedFact> RankDescending(
            IEnumerable<VerifiedFact> inputs)
        {
            var values = (inputs ?? new VerifiedFact[0])
                .Select(fact => new { Fact = fact,
                    Value = AnalysisContract.Decimal(fact) })
                .OrderByDescending(item => item.Value)
                .ThenBy(item => item.Fact.FactId, StringComparer.Ordinal)
                .ToList();
            var result = new List<RankedFact>();
            decimal? prior = null;
            var rank = 0;
            foreach (var item in values)
            {
                if (!prior.HasValue || prior.Value != item.Value) rank++;
                result.Add(new RankedFact
                {
                    FactId = item.Fact.FactId,
                    Rank = rank,
                    Value = AnalysisContract.Format(item.Value)
                });
                prior = item.Value;
            }
            return result;
        }

        private static decimal Divide(decimal numerator, decimal denominator)
        {
            if (denominator == 0m)
                throw new InvalidOperationException(
                    "ANALYSIS_ZERO_DENOMINATOR: Ratio calculations cannot divide by zero.");
            return numerator / denominator;
        }

        private static void RequireCount(
            ICollection<VerifiedFact> facts,
            int count,
            string operation)
        {
            if (facts.Count != count)
                throw new InvalidOperationException(
                    "ANALYSIS_INPUT_COUNT_INVALID: " + operation +
                    " requires " + count.ToString(CultureInfo.InvariantCulture) +
                    " inputs.");
        }

        private static void SameUnit(IList<VerifiedFact> facts)
        {
            if (facts.Any(fact => !string.Equals(fact.Unit, facts[0].Unit,
                    StringComparison.Ordinal) ||
                !string.Equals(fact.Currency, facts[0].Currency,
                    StringComparison.Ordinal)))
                throw new InvalidOperationException(
                    "ANALYSIS_UNIT_MISMATCH: Calculation inputs need the same unit and currency.");
        }
    }

    public static class WorkbookTypedCapture
    {
        public static TableDataset Capture(
            string tableId,
            string name,
            object values,
            object formulas,
            object numberFormats,
            object displayText,
            int rows,
            int columns,
            int firstRow,
            int firstColumn)
        {
            if (rows < 1 || columns < 1 || rows > 20000 || columns > 16384 ||
                firstRow < 1 || firstColumn < 1)
                throw new InvalidOperationException(
                    "ANALYSIS_RANGE_INVALID: Typed workbook capture needs a bounded positive range.");
            var table = new TableDataset
            {
                TableId = tableId,
                Name = name ?? string.Empty,
                Rows = rows,
                Columns = columns
            };
            for (var row = 0; row < rows; row++)
            {
                for (var column = 0; column < columns; column++)
                {
                    var raw = Matrix(values, row, column, rows, columns, false);
                    var formulaRaw = Matrix(formulas, row, column, rows, columns, false);
                    var formatRaw = Matrix(numberFormats, row, column, rows, columns, true);
                    var shown = Matrix(displayText, row, column, rows, columns, false);
                    var format = Convert.ToString(formatRaw,
                        CultureInfo.InvariantCulture) ?? string.Empty;
                    var formula = Convert.ToString(formulaRaw,
                        CultureInfo.InvariantCulture) ?? string.Empty;
                    if (!formula.StartsWith("=", StringComparison.Ordinal))
                        formula = string.Empty;
                    var cell = Cell(raw, format, formula, shown);
                    cell.Row = row;
                    cell.Column = column;
                    cell.Reference = Address(firstRow + row,
                        firstColumn + column);
                    table.Cells.Add(cell);
                }
            }
            return table;
        }

        private static DatasetCell Cell(
            object raw,
            string format,
            string formula,
            object shown)
        {
            var result = new DatasetCell
            {
                Formula = formula,
                NumberFormat = format,
                RawValue = Raw(raw),
                DisplayText = shown == null
                    ? string.Empty
                    : Convert.ToString(shown,
                        CultureInfo.InvariantCulture) ?? string.Empty,
                Status = AnalysisContract.Verified
            };
            var error = ExcelErrorValue.Text(raw);
            if (error != null)
            {
                result.ValueType = AnalysisContract.ErrorValue;
                result.RawValue = error;
                result.Value = string.Empty;
                result.DisplayText = result.DisplayText.Length == 0
                    ? error
                    : result.DisplayText;
                result.Status = AnalysisContract.Unresolved;
                return result;
            }
            if (raw == null)
            {
                result.ValueType = AnalysisContract.MissingValue;
                result.Value = string.Empty;
                result.Status = AnalysisContract.Unresolved;
                return result;
            }
            if (raw is bool)
            {
                result.ValueType = AnalysisContract.BooleanValue;
                result.Value = (bool)raw ? "true" : "false";
            }
            else if (raw is DateTime)
            {
                result.ValueType = AnalysisContract.DateValue;
                result.Value = ((DateTime)raw).ToString("O",
                    CultureInfo.InvariantCulture);
            }
            else if (IsNumber(raw))
            {
                var number = Convert.ToDouble(raw,
                    CultureInfo.InvariantCulture);
                if (LooksLikeDateFormat(format))
                {
                    result.ValueType = AnalysisContract.DateValue;
                    try
                    {
                        result.Value = DateTime.FromOADate(number).ToString(
                            "O", CultureInfo.InvariantCulture);
                    }
                    catch (ArgumentException)
                    {
                        result.ValueType = AnalysisContract.ErrorValue;
                        result.Value = string.Empty;
                        result.Status = AnalysisContract.Unresolved;
                    }
                }
                else
                {
                    result.ValueType = raw is byte || raw is sbyte ||
                        raw is short || raw is ushort || raw is int ||
                        raw is uint || raw is long || raw is ulong
                            ? AnalysisContract.IntegerValue
                            : AnalysisContract.DecimalValue;
                    result.Value = Raw(raw);
                }
            }
            else
            {
                result.ValueType = AnalysisContract.TextValue;
                result.Value = Convert.ToString(raw,
                    CultureInfo.InvariantCulture) ?? string.Empty;
            }
            if (result.DisplayText.Length == 0) result.DisplayText = result.Value;
            return result;
        }

        private static object Matrix(
            object value,
            int row,
            int column,
            int rows,
            int columns,
            bool repeatScalar)
        {
            var grid = value as object[,];
            if (grid != null)
            {
                if (grid.GetLength(0) != rows || grid.GetLength(1) != columns)
                    throw new InvalidOperationException(
                        "ANALYSIS_MATRIX_SIZE_MISMATCH: Workbook value, formula, format, and display matrices must align.");
                return grid[grid.GetLowerBound(0) + row,
                    grid.GetLowerBound(1) + column];
            }
            return rows == 1 && columns == 1 || repeatScalar ? value : null;
        }

        private static bool IsNumber(object value)
        {
            return value is byte || value is sbyte || value is short ||
                value is ushort || value is int || value is uint ||
                value is long || value is ulong || value is float ||
                value is double || value is decimal;
        }

        private static string Raw(object value)
        {
            if (value == null) return string.Empty;
            if (value is double) return ((double)value).ToString(
                "R", CultureInfo.InvariantCulture);
            if (value is float) return ((float)value).ToString(
                "R", CultureInfo.InvariantCulture);
            if (value is decimal) return ((decimal)value).ToString(
                CultureInfo.InvariantCulture);
            if (value is DateTime) return ((DateTime)value).ToString(
                "O", CultureInfo.InvariantCulture);
            return Convert.ToString(value,
                CultureInfo.InvariantCulture) ?? string.Empty;
        }

        private static bool LooksLikeDateFormat(string format)
        {
            if (string.IsNullOrWhiteSpace(format)) return false;
            var normalized = format.ToLowerInvariant();
            var quoted = false;
            var escaped = false;
            var bracketed = false;
            var tokens = new StringBuilder();
            foreach (var character in normalized)
            {
                if (escaped) { escaped = false; continue; }
                if (character == '\\') { escaped = true; continue; }
                if (character == '"') { quoted = !quoted; continue; }
                if (!quoted && character == '[') { bracketed = true; continue; }
                if (!quoted && character == ']') { bracketed = false; continue; }
                if (!quoted && !bracketed &&
                    "ymdhis".IndexOf(character) >= 0)
                    tokens.Append(character);
            }
            return tokens.Length > 0;
        }

        private static string Address(int row, int column)
        {
            var letters = string.Empty;
            var value = column;
            while (value > 0)
            {
                value--;
                letters = (char)('A' + value % 26) + letters;
                value /= 26;
            }
            return letters + row.ToString(CultureInfo.InvariantCulture);
        }
    }

    public sealed class SourceSnapshot
    {
        public int ContractVersion { get; set; }
        public string SnapshotId { get; set; }
        public string SourceInstanceId { get; set; }
        public string SourceType { get; set; }
        public string CaptureRevision { get; set; }
        public string CapturedUtc { get; set; }
        public string ContentHash { get; set; }
        public string Coverage { get; set; }
        public string CalculationState { get; set; }
        public List<SourceLocator> Locators { get; set; } =
            new List<SourceLocator>();
        public List<TableDataset> Tables { get; set; } =
            new List<TableDataset>();
    }

    public sealed class SourceLocator
    {
        public string Kind { get; set; }
        public string SourceInstanceId { get; set; }
        public string WorksheetIdentity { get; set; }
        public string Range { get; set; }
        public string Cell { get; set; }
        public string PresentationIdentity { get; set; }
        public string SlideId { get; set; }
        public string ShapeId { get; set; }
    }

    public sealed class TableDataset
    {
        public string TableId { get; set; }
        public string Name { get; set; }
        public int Rows { get; set; }
        public int Columns { get; set; }
        public List<DatasetCell> Cells { get; set; } =
            new List<DatasetCell>();
    }

    public sealed class DatasetCell
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public string Reference { get; set; }
        public string ValueType { get; set; }
        public string RawValue { get; set; }
        public string Value { get; set; }
        public string DisplayText { get; set; }
        public string Formula { get; set; }
        public string NumberFormat { get; set; }
        public string Status { get; set; }
    }

    public sealed class VerifiedFact
    {
        public string FactId { get; set; }
        public string SnapshotId { get; set; }
        public string Kind { get; set; }
        public string ValueType { get; set; }
        public string Value { get; set; }
        public string DisplayText { get; set; }
        public string Metric { get; set; }
        public string Unit { get; set; }
        public string Currency { get; set; }
        public string Period { get; set; }
        public Dictionary<string, string> Dimensions { get; set; } =
            new Dictionary<string, string>();
        public List<SourceLocator> Locators { get; set; } =
            new List<SourceLocator>();
        public string Status { get; set; }
    }

    public sealed class AnalysisCalculation
    {
        public string CalculationId { get; set; }
        public string Operation { get; set; }
        public List<string> InputFactIds { get; set; } = new List<string>();
        public string OutputFactId { get; set; }
        public string NullPolicy { get; set; }
        public string RoundingPolicy { get; set; }
        public string Status { get; set; }
    }

    public sealed class RankedFact
    {
        public string FactId { get; set; }
        public int Rank { get; set; }
        public string Value { get; set; }
    }

    public sealed class AnalysisArtifact
    {
        public int ContractVersion { get; set; }
        public string AnalysisId { get; set; }
        public List<SourceSnapshot> Snapshots { get; set; } =
            new List<SourceSnapshot>();
        public List<VerifiedFact> Facts { get; set; } =
            new List<VerifiedFact>();
        public List<AnalysisCalculation> Calculations { get; set; } =
            new List<AnalysisCalculation>();
        public List<string> Assumptions { get; set; } = new List<string>();
        public List<string> UnresolvedConflicts { get; set; } =
            new List<string>();
    }
}

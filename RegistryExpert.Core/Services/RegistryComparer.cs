using RegistryParser.Abstractions;

namespace RegistryExpert.Core.Services
{
    /// <summary>
    /// Stateless comparison service for registry hives.
    /// Extracted from Forms/CompareForm.cs comparison algorithm.
    /// </summary>
    public static class RegistryComparer
    {
        /// <summary>Normalized root name used as the top-level key in path dictionaries.</summary>
        public const string NormalizedRootName = "ROOT";

        /// <summary>
        /// Pre-computed diff status for a registry key path.
        /// </summary>
        public record struct DiffInfo(
            bool HasDifference,          // This key or any descendant has a difference
            bool IsUniqueToThisHive,     // GREEN: unique AND all descendants unique
            bool HasValueDifference      // RED: value diff, mixed descendants, or child diffs
        );

        /// <summary>
        /// Build a flat dictionary of all keys by normalized path.
        /// Walks the entire hive recursively.
        /// </summary>
        public static Dictionary<string, RegistryKey> BuildKeyIndex(RegistryKey? rootKey)
        {
            var result = new Dictionary<string, RegistryKey>(StringComparer.OrdinalIgnoreCase);
            if (rootKey == null) return result;
            BuildKeyIndexRecursive(rootKey, "", result);
            return result;
        }

        private static void BuildKeyIndexRecursive(RegistryKey key, string parentPath, Dictionary<string, RegistryKey> result)
        {
            var path = string.IsNullOrEmpty(parentPath) ? NormalizedRootName : $"{parentPath}\\{key.KeyName}";
            result[path] = key;

            if (key.SubKeys != null)
            {
                foreach (var subKey in key.SubKeys)
                    BuildKeyIndexRecursive(subKey, path, result);
            }
        }

        /// <summary>
        /// Pre-compute diff status for every key in one hive against the other hive's index.
        /// </summary>
        public static Dictionary<string, DiffInfo> ComputeDiff(
            RegistryKey? rootKey,
            Dictionary<string, RegistryKey> otherIndex,
            CancellationToken token = default)
        {
            var result = new Dictionary<string, DiffInfo>(StringComparer.OrdinalIgnoreCase);
            if (rootKey != null)
                ComputeDiffRecursive(rootKey, "", otherIndex, result, token);
            return result;
        }

        private static void ComputeDiffRecursive(
            RegistryKey key,
            string parentPath,
            Dictionary<string, RegistryKey> otherIndex,
            Dictionary<string, DiffInfo> result,
            CancellationToken token)
        {
            token.ThrowIfCancellationRequested();

            var path = string.IsNullOrEmpty(parentPath) ? NormalizedRootName : $"{parentPath}\\{key.KeyName}";

            bool isUnique = !otherIndex.TryGetValue(path, out var otherKey);
            bool hasValueDiff = false;

            if (!isUnique && otherKey != null)
                hasValueDiff = HasValueDifferences(key, otherKey);

            bool anyChildHasDiff = false;
            bool allChildrenUnique = true;

            if (key.SubKeys != null)
            {
                foreach (var sub in key.SubKeys)
                {
                    ComputeDiffRecursive(sub, path, otherIndex, result, token);
                    var childPath = $"{path}\\{sub.KeyName}";
                    if (result.TryGetValue(childPath, out var childDiff) && childDiff.HasDifference)
                    {
                        anyChildHasDiff = true;
                        if (!childDiff.IsUniqueToThisHive)
                            allChildrenUnique = false;
                    }
                }
            }

            bool hasDiff = isUnique || hasValueDiff || anyChildHasDiff;
            bool nodeIsUnique = false;
            bool nodeHasValueDiff = false;

            if (hasDiff)
            {
                if (isUnique && (!anyChildHasDiff || allChildrenUnique))
                    nodeIsUnique = true;
                else
                    nodeHasValueDiff = true;
            }

            result[path] = new DiffInfo(hasDiff, nodeIsUnique, nodeHasValueDiff);
        }

        /// <summary>
        /// Compare values between two keys. Returns true if any difference is found.
        /// </summary>
        public static bool HasValueDifferences(RegistryKey left, RegistryKey right)
        {
            var leftValues = left.Values ?? new List<KeyValue>();
            var rightValues = right.Values ?? new List<KeyValue>();

            if (leftValues.Count != rightValues.Count)
                return true;

            var leftDict = new Dictionary<string, KeyValue>(StringComparer.OrdinalIgnoreCase);
            foreach (var v in leftValues)
            {
                var name = v.ValueName ?? "(Default)";
                if (!leftDict.ContainsKey(name))
                    leftDict[name] = v;
            }

            foreach (var rv in rightValues)
            {
                var name = rv.ValueName ?? "(Default)";
                if (!leftDict.TryGetValue(name, out var lv))
                    return true;

                if (lv.ValueType != rv.ValueType)
                    return true;

                var leftData = lv.ValueData?.ToString() ?? "";
                var rightData = rv.ValueData?.ToString() ?? "";
                if (leftData != rightData)
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Get per-value diff details for display in the values grid.
        /// Returns list of ValueDiffItem with diff status for each value.
        /// </summary>
        public static List<ValueDiffItem> GetValueDiffs(
            RegistryKey key,
            string keyPath,
            Dictionary<string, RegistryKey> otherKeyIndex,
            bool showDifferencesOnly)
        {
            var items = new List<ValueDiffItem>();
            if (key.Values == null || key.Values.Count == 0)
                return items;

            // Build other side's value lookup
            Dictionary<string, KeyValue>? otherValues = null;
            if (otherKeyIndex.TryGetValue(keyPath, out var otherKey) && otherKey.Values != null)
            {
                otherValues = new Dictionary<string, KeyValue>(StringComparer.OrdinalIgnoreCase);
                foreach (var v in otherKey.Values)
                {
                    var name = v.ValueName ?? "(Default)";
                    if (!otherValues.ContainsKey(name))
                        otherValues[name] = v;
                }
            }

            foreach (var value in key.Values.OrderBy(v => v.ValueName ?? "(Default)", StringComparer.OrdinalIgnoreCase))
            {
                var name = value.ValueName ?? "(Default)";
                var type = value.ValueType ?? "";
                var data = FormatValue(value);

                ValueDiffStatus status;
                if (otherValues == null || !otherValues.ContainsKey(name))
                {
                    status = ValueDiffStatus.UniqueToThisHive; // GREEN
                }
                else
                {
                    var otherValue = otherValues[name];
                    var otherData = FormatValue(otherValue);
                    status = (type != otherValue.ValueType || data != otherData)
                        ? ValueDiffStatus.ValueDiffers  // RED
                        : ValueDiffStatus.Identical;
                }

                if (showDifferencesOnly && status == ValueDiffStatus.Identical)
                    continue;

                items.Add(new ValueDiffItem(name, type, data, status));
            }

            return items;
        }

        /// <summary>Format a KeyValue for display (same logic as CompareForm.FormatValue).</summary>
        public static string FormatValue(KeyValue value)
        {
            if (value.ValueData == null)
                return "(null)";

            try
            {
                var data = value.ValueData.ToString() ?? "";

                if (value.ValueType?.ToUpperInvariant() == "REGBINARY" && value.ValueDataRaw != null && value.ValueDataRaw.Length > 0)
                {
                    var hex = BitConverter.ToString(value.ValueDataRaw.Take(32).ToArray()).Replace("-", " ");
                    if (value.ValueDataRaw.Length > 32)
                        hex += $"... ({value.ValueDataRaw.Length} bytes)";
                    return hex;
                }

                if (data.Length > 200)
                    return data.Substring(0, 200) + "...";

                return data;
            }
            catch
            {
                return value.ValueData?.ToString() ?? "(error)";
            }
        }
    }

    /// <summary>Diff status for a single registry value.</summary>
    public enum ValueDiffStatus
    {
        Identical,
        UniqueToThisHive,   // GREEN
        ValueDiffers         // RED
    }

    /// <summary>A single value entry with diff status for display.</summary>
    public record ValueDiffItem(string Name, string Type, string Data, ValueDiffStatus Status);
}

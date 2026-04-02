using RegistryParser.Abstractions;

namespace RegistryExpert.Core
{
    /// <summary>
    /// Statistics about a single registry key subtree.
    /// </summary>
    public class KeyStatistics
    {
        public string KeyPath { get; set; } = "";
        public int SubKeyCount { get; set; }
        public int ValueCount { get; set; }
        public long TotalSize { get; set; }
    }

    /// <summary>
    /// Statistics about a single registry value (name, type, data size).
    /// Used by the Value Counts tab to list individual values within a key.
    /// </summary>
    public class ValueStatistic
    {
        public string ValueName { get; set; } = "";
        public string ValueType { get; set; } = "";
        public long DataSize { get; set; }
    }

    /// <summary>
    /// Analyzes registry hive structure for bloat detection and size reporting.
    /// Extracted from WinForms MainForm for reuse by WPF Statistics window.
    /// </summary>
    public static class RegistryStatisticsAnalyzer
    {
        /// <summary>
        /// Analyze all top-level keys under the hive root, computing subkey counts, value counts, and data sizes.
        /// </summary>
        public static List<KeyStatistics> AnalyzeTopLevelKeys(OfflineRegistryParser parser)
        {
            var results = new List<KeyStatistics>();
            var rootKey = parser.GetRootKey();
            if (rootKey?.SubKeys == null) return results;

            foreach (var topKey in rootKey.SubKeys)
            {
                var (subKeyCount, valueCount, totalSize) = CalculateKeyStatisticsRecursive(topKey);
                var stat = new KeyStatistics
                {
                    KeyPath = topKey.KeyName,
                    SubKeyCount = subKeyCount,
                    ValueCount = valueCount,
                    TotalSize = totalSize
                };
                results.Add(stat);
            }

            return results;
        }

        /// <summary>
        /// Analyze immediate children of a given key path, computing stats for each child subtree.
        /// </summary>
        public static List<KeyStatistics> GetChildKeyStats(OfflineRegistryParser parser, string parentPath, bool calculateSize)
        {
            var results = new List<KeyStatistics>();

            var parentKey = parser.GetKey(parentPath);
            if (parentKey?.SubKeys == null) return results;

            foreach (var subKey in parentKey.SubKeys)
            {
                var childPath = $"{parentPath}\\{subKey.KeyName}";
                // Get the fully loaded key to ensure SubKeys are populated
                var fullChildKey = parser.GetKey(childPath);

                var (subKeyCount, valueCount, totalSize) = CalculateKeyStatisticsRecursive(fullChildKey);
                var stat = new KeyStatistics
                {
                    KeyPath = childPath,
                    SubKeyCount = subKeyCount,
                    ValueCount = valueCount,
                    TotalSize = calculateSize ? totalSize : 0
                };
                results.Add(stat);
            }

            return results;
        }

        /// <summary>
        /// Single-pass recursive traversal that calculates subkey count, value count, and total size at once.
        /// </summary>
        public static (int subKeyCount, int valueCount, long totalSize) CalculateKeyStatisticsRecursive(RegistryKey? key)
        {
            if (key == null) return (0, 0, 0);

            // Calculate size for this key - only actual data, no overhead
            long size = ((long)(key.KeyName?.Length ?? 0)) * 2;  // Key name only
            int valueCount = key.Values?.Count ?? 0;

            // Add size of values
            if (key.Values != null)
            {
                foreach (var val in key.Values)
                {
                    size += ((long)(val.ValueName?.Length ?? 0)) * 2;
                    size += GetValueDataSize(val);
                }
            }

            // Leaf key (no children)
            if (key.SubKeys == null || key.SubKeys.Count == 0)
            {
                return (1, valueCount, size);
            }

            // Non-leaf: recursively accumulate from children (start at 1 to count this key itself)
            int subKeyCount = 1;
            foreach (var subKey in key.SubKeys)
            {
                var (childSubKeys, childValues, childSize) = CalculateKeyStatisticsRecursive(subKey);
                subKeyCount += childSubKeys;
                valueCount += childValues;
                size += childSize;
            }

            return (subKeyCount, valueCount, size);
        }

        /// <summary>
        /// Estimate the data size in bytes for a registry value based on its type.
        /// </summary>
        public static long GetValueDataSize(KeyValue val)
        {
            if (val.ValueData == null) return 0;

            // Try to get actual data size based on type
            var dataStr = val.ValueData.ToString() ?? "";

            // Check for binary data (hex string format from Registry library)
            if (val.ValueType == "RegBinary" || dataStr.Contains(' ') && IsHexString(dataStr))
            {
                // Binary data is typically shown as hex bytes separated by spaces
                var parts = dataStr.Split(new[] { ' ', '-' }, StringSplitOptions.RemoveEmptyEntries);
                return parts.Length;
            }

            // REG_DWORD
            if (val.ValueType == "RegDword")
                return 4;

            // REG_QWORD
            if (val.ValueType == "RegQword")
                return 8;

            // REG_SZ, REG_EXPAND_SZ - Unicode string
            if (val.ValueType == "RegSz" || val.ValueType == "RegExpandSz")
                return (dataStr.Length + 1) * 2; // Unicode + null terminator

            // REG_MULTI_SZ - Multiple strings
            if (val.ValueType == "RegMultiSz")
                return (dataStr.Length + 2) * 2; // Unicode + double null terminator

            // Default: estimate based on string representation
            return dataStr.Length;
        }

        /// <summary>
        /// Get individual value statistics for a specific key.
        /// Returns each value's name, type, and data size for display in the Value Counts tab.
        /// </summary>
        public static List<ValueStatistic> GetKeyValueStats(OfflineRegistryParser parser, string keyPath)
        {
            var results = new List<ValueStatistic>();

            var key = parser.GetKey(keyPath);
            if (key?.Values == null || key.Values.Count == 0) return results;

            foreach (var val in key.Values)
            {
                results.Add(new ValueStatistic
                {
                    ValueName = val.ValueName ?? "(Default)",
                    ValueType = val.ValueType ?? "Unknown",
                    DataSize = GetValueDataSize(val)
                });
            }

            return results;
        }

        /// <summary>
        /// Format a byte count as a human-readable string (B/K/M/G).
        /// </summary>
        public static string FormatSize(long bytes)
        {
            if (bytes >= 1_000_000_000) return $"{bytes / 1_000_000_000.0:F1}G";
            if (bytes >= 1_000_000) return $"{bytes / 1_000_000.0:F1}M";
            if (bytes >= 1_000) return $"{bytes / 1_000.0:F1}K";
            return bytes.ToString("N0");
        }

        /// <summary>
        /// Check if a string looks like hex data (digits and A-F only, ignoring spaces/dashes).
        /// </summary>
        public static bool IsHexString(string str)
        {
            if (string.IsNullOrEmpty(str)) return false;
            var trimmed = str.Replace(" ", "").Replace("-", "");
            return trimmed.Length > 0 && trimmed.All(c =>
                (c >= '0' && c <= '9') ||
                (c >= 'A' && c <= 'F') ||
                (c >= 'a' && c <= 'f'));
        }
    }
}

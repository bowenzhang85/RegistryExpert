using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Registry.Abstractions;

namespace RegistryExpert
{
    /// <summary>
    /// Type of difference found during comparison
    /// </summary>
    public enum DiffType
    {
        /// <summary>Item exists only in the right hive (added)</summary>
        Added,
        /// <summary>Item exists only in the left hive (removed)</summary>
        Removed,
        /// <summary>Item exists in both but with different values</summary>
        Modified,
        /// <summary>Item is identical in both hives</summary>
        Unchanged
    }

    /// <summary>
    /// Represents a difference in a registry key
    /// </summary>
    public class KeyDiff
    {
        public string KeyPath { get; set; } = "";
        public string KeyName { get; set; } = "";
        public DiffType ChangeType { get; set; }
        public DateTimeOffset? LeftTimestamp { get; set; }
        public DateTimeOffset? RightTimestamp { get; set; }
        public List<ValueDiff> ValueDiffs { get; set; } = new();
        public List<KeyDiff> SubKeyDiffs { get; set; } = new();

        /// <summary>
        /// Returns true if this key or any of its descendants have differences
        /// </summary>
        public bool HasDifferences =>
            ChangeType != DiffType.Unchanged ||
            ValueDiffs.Any(v => v.ChangeType != DiffType.Unchanged) ||
            SubKeyDiffs.Any(s => s.HasDifferences);
    }

    /// <summary>
    /// Represents a difference in a registry value
    /// </summary>
    public class ValueDiff
    {
        public string ValueName { get; set; } = "";
        public DiffType ChangeType { get; set; }
        public string? LeftValue { get; set; }
        public string? RightValue { get; set; }
        public string? LeftType { get; set; }
        public string? RightType { get; set; }
    }

    /// <summary>
    /// Result of comparing two registry hives
    /// </summary>
    public class ComparisonResult
    {
        public KeyDiff RootDiff { get; set; } = new();
        public int AddedKeyCount { get; set; }
        public int RemovedKeyCount { get; set; }
        public int ModifiedKeyCount { get; set; }
        public int AddedValueCount { get; set; }
        public int RemovedValueCount { get; set; }
        public int ModifiedValueCount { get; set; }

        public int TotalKeyChanges => AddedKeyCount + RemovedKeyCount + ModifiedKeyCount;
        public int TotalValueChanges => AddedValueCount + RemovedValueCount + ModifiedValueCount;
        public int TotalChanges => TotalKeyChanges + TotalValueChanges;

        /// <summary>
        /// Gets a flat list of all differences for the unified view
        /// </summary>
        public List<(DiffType Type, string Path, string ValueName, string? LeftValue, string? RightValue, string? LeftType, string? RightType, bool IsKey)> GetFlatDifferences()
        {
            var result = new List<(DiffType, string, string, string?, string?, string?, string?, bool)>();
            CollectDifferences(RootDiff, result);
            return result;
        }

        private void CollectDifferences(KeyDiff keyDiff, List<(DiffType, string, string, string?, string?, string?, string?, bool)> result)
        {
            // Add key-level difference if the key itself was added/removed
            if (keyDiff.ChangeType == DiffType.Added || keyDiff.ChangeType == DiffType.Removed)
            {
                result.Add((keyDiff.ChangeType, keyDiff.KeyPath, "(key)", null, null, null, null, true));
            }

            // Add value differences
            foreach (var valueDiff in keyDiff.ValueDiffs.Where(v => v.ChangeType != DiffType.Unchanged))
            {
                result.Add((valueDiff.ChangeType, keyDiff.KeyPath, valueDiff.ValueName, 
                    valueDiff.LeftValue, valueDiff.RightValue, valueDiff.LeftType, valueDiff.RightType, false));
            }

            // Recurse into subkeys
            foreach (var subKey in keyDiff.SubKeyDiffs)
            {
                CollectDifferences(subKey, result);
            }
        }
    }

    /// <summary>
    /// Compares two offline registry hives and finds differences
    /// </summary>
    public class RegistryComparer
    {
        private int _addedKeyCount;
        private int _removedKeyCount;
        private int _modifiedKeyCount;
        private int _addedValueCount;
        private int _removedValueCount;
        private int _modifiedValueCount;
        private IProgress<string>? _progress;

        /// <summary>
        /// Compares two registry hives and returns the differences
        /// </summary>
        public ComparisonResult Compare(OfflineRegistryParser left, OfflineRegistryParser right, IProgress<string>? progress = null)
        {
            _progress = progress;
            
            // Reset counters
            _addedKeyCount = 0;
            _removedKeyCount = 0;
            _modifiedKeyCount = 0;
            _addedValueCount = 0;
            _removedValueCount = 0;
            _modifiedValueCount = 0;

            _progress?.Report("Starting comparison...");
            
            var result = new ComparisonResult
            {
                RootDiff = CompareKeys(left.GetRootKey(), right.GetRootKey(), "", 0)
            };

            result.AddedKeyCount = _addedKeyCount;
            result.RemovedKeyCount = _removedKeyCount;
            result.ModifiedKeyCount = _modifiedKeyCount;
            result.AddedValueCount = _addedValueCount;
            result.RemovedValueCount = _removedValueCount;
            result.ModifiedValueCount = _modifiedValueCount;

            _progress?.Report("Comparison complete");
            return result;
        }

        /// <summary>
        /// Compares two registry hives asynchronously with parallel subkey processing
        /// </summary>
        public async Task<ComparisonResult> CompareAsync(OfflineRegistryParser left, OfflineRegistryParser right, IProgress<string>? progress = null, CancellationToken cancellationToken = default)
        {
            _progress = progress;
            
            // Reset counters
            _addedKeyCount = 0;
            _removedKeyCount = 0;
            _modifiedKeyCount = 0;
            _addedValueCount = 0;
            _removedValueCount = 0;
            _modifiedValueCount = 0;

            _progress?.Report("Starting comparison...");
            
            var result = new ComparisonResult();
            
            // Run comparison in background with parallel processing for top-level subkeys
            result.RootDiff = await Task.Run(() => CompareKeysParallel(left.GetRootKey(), right.GetRootKey(), "", cancellationToken), cancellationToken);

            result.AddedKeyCount = _addedKeyCount;
            result.RemovedKeyCount = _removedKeyCount;
            result.ModifiedKeyCount = _modifiedKeyCount;
            result.AddedValueCount = _addedValueCount;
            result.RemovedValueCount = _removedValueCount;
            result.ModifiedValueCount = _modifiedValueCount;

            _progress?.Report("Comparison complete");
            return result;
        }

        private KeyDiff CompareKeysParallel(RegistryKey? left, RegistryKey? right, string parentPath, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            
            var keyName = left?.KeyName ?? right?.KeyName ?? "";
            var keyPath = string.IsNullOrEmpty(parentPath) ? keyName : $"{parentPath}\\{keyName}";

            // Report progress for top-level keys
            if (parentPath.Count(c => c == '\\') <= 1)
            {
                _progress?.Report($"Comparing: {keyPath}");
            }

            var diff = new KeyDiff
            {
                KeyPath = keyPath,
                KeyName = keyName,
                LeftTimestamp = left?.LastWriteTime,
                RightTimestamp = right?.LastWriteTime
            };

            // Determine key-level change type
            if (left == null && right != null)
            {
                diff.ChangeType = DiffType.Added;
                Interlocked.Increment(ref _addedKeyCount);
                // Count all values in added key
                if (right.Values != null)
                    Interlocked.Add(ref _addedValueCount, right.Values.Count);
                // Recursively count added subkeys
                CountAddedKeysThreadSafe(right);
            }
            else if (left != null && right == null)
            {
                diff.ChangeType = DiffType.Removed;
                Interlocked.Increment(ref _removedKeyCount);
                // Count all values in removed key
                if (left.Values != null)
                    Interlocked.Add(ref _removedValueCount, left.Values.Count);
                // Recursively count removed subkeys
                CountRemovedKeysThreadSafe(left);
            }
            else if (left != null && right != null)
            {
                // Compare values
                diff.ValueDiffs = CompareValues(left.Values, right.Values);
                
                // Determine if key is modified (has value differences)
                if (diff.ValueDiffs.Any(v => v.ChangeType != DiffType.Unchanged))
                {
                    diff.ChangeType = DiffType.Modified;
                    Interlocked.Increment(ref _modifiedKeyCount);
                }
                else
                {
                    diff.ChangeType = DiffType.Unchanged;
                }

                // Compare subkeys - use parallel processing for deep comparisons
                var leftSubKeys = left.SubKeys?.ToDictionary(k => k.KeyName, StringComparer.OrdinalIgnoreCase) 
                    ?? new Dictionary<string, RegistryKey>();
                var rightSubKeys = right.SubKeys?.ToDictionary(k => k.KeyName, StringComparer.OrdinalIgnoreCase) 
                    ?? new Dictionary<string, RegistryKey>();

                var allKeyNames = leftSubKeys.Keys.Union(rightSubKeys.Keys, StringComparer.OrdinalIgnoreCase)
                    .OrderBy(k => k, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                // Use parallel processing if there are many subkeys
                if (allKeyNames.Count > 10 && parentPath.Count(c => c == '\\') < 3)
                {
                    var subDiffs = new ConcurrentBag<KeyDiff>();
                    
                    Parallel.ForEach(allKeyNames, new ParallelOptions 
                    { 
                        MaxDegreeOfParallelism = Environment.ProcessorCount,
                        CancellationToken = cancellationToken
                    }, 
                    subKeyName =>
                    {
                        leftSubKeys.TryGetValue(subKeyName, out var leftSubKey);
                        rightSubKeys.TryGetValue(subKeyName, out var rightSubKey);
                        var subDiff = CompareKeysParallel(leftSubKey, rightSubKey, keyPath, cancellationToken);
                        subDiffs.Add(subDiff);
                    });

                    // Sort results after parallel processing
                    diff.SubKeyDiffs = subDiffs.OrderBy(d => d.KeyName, StringComparer.OrdinalIgnoreCase).ToList();
                }
                else
                {
                    // Sequential processing for small number of subkeys
                    foreach (var subKeyName in allKeyNames)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        leftSubKeys.TryGetValue(subKeyName, out var leftSubKey);
                        rightSubKeys.TryGetValue(subKeyName, out var rightSubKey);
                        var subDiff = CompareKeysParallel(leftSubKey, rightSubKey, keyPath, cancellationToken);
                        diff.SubKeyDiffs.Add(subDiff);
                    }
                }
            }

            return diff;
        }

        private KeyDiff CompareKeys(RegistryKey? left, RegistryKey? right, string parentPath, int depth)
        {
            var keyName = left?.KeyName ?? right?.KeyName ?? "";
            var keyPath = string.IsNullOrEmpty(parentPath) ? keyName : $"{parentPath}\\{keyName}";

            var diff = new KeyDiff
            {
                KeyPath = keyPath,
                KeyName = keyName,
                LeftTimestamp = left?.LastWriteTime,
                RightTimestamp = right?.LastWriteTime
            };

            // Determine key-level change type
            if (left == null && right != null)
            {
                diff.ChangeType = DiffType.Added;
                _addedKeyCount++;
                // Count all values in added key
                if (right.Values != null)
                    _addedValueCount += right.Values.Count;
                // Recursively count added subkeys
                CountAddedKeys(right);
            }
            else if (left != null && right == null)
            {
                diff.ChangeType = DiffType.Removed;
                _removedKeyCount++;
                // Count all values in removed key
                if (left.Values != null)
                    _removedValueCount += left.Values.Count;
                // Recursively count removed subkeys
                CountRemovedKeys(left);
            }
            else if (left != null && right != null)
            {
                // Compare values
                diff.ValueDiffs = CompareValues(left.Values, right.Values);
                
                // Determine if key is modified (has value differences)
                if (diff.ValueDiffs.Any(v => v.ChangeType != DiffType.Unchanged))
                {
                    diff.ChangeType = DiffType.Modified;
                    _modifiedKeyCount++;
                }
                else
                {
                    diff.ChangeType = DiffType.Unchanged;
                }

                // Compare subkeys
                var leftSubKeys = left.SubKeys?.ToDictionary(k => k.KeyName, StringComparer.OrdinalIgnoreCase) 
                    ?? new Dictionary<string, RegistryKey>();
                var rightSubKeys = right.SubKeys?.ToDictionary(k => k.KeyName, StringComparer.OrdinalIgnoreCase) 
                    ?? new Dictionary<string, RegistryKey>();

                var allKeyNames = leftSubKeys.Keys.Union(rightSubKeys.Keys, StringComparer.OrdinalIgnoreCase);

                foreach (var subKeyName in allKeyNames.OrderBy(k => k, StringComparer.OrdinalIgnoreCase))
                {
                    leftSubKeys.TryGetValue(subKeyName, out var leftSubKey);
                    rightSubKeys.TryGetValue(subKeyName, out var rightSubKey);

                    var subDiff = CompareKeys(leftSubKey, rightSubKey, keyPath, depth + 1);
                    diff.SubKeyDiffs.Add(subDiff);
                }
            }

            return diff;
        }

        private void CountAddedKeysThreadSafe(RegistryKey key)
        {
            if (key.SubKeys == null) return;
            foreach (var subKey in key.SubKeys)
            {
                Interlocked.Increment(ref _addedKeyCount);
                if (subKey.Values != null)
                    Interlocked.Add(ref _addedValueCount, subKey.Values.Count);
                CountAddedKeysThreadSafe(subKey);
            }
        }

        private void CountRemovedKeysThreadSafe(RegistryKey key)
        {
            if (key.SubKeys == null) return;
            foreach (var subKey in key.SubKeys)
            {
                Interlocked.Increment(ref _removedKeyCount);
                if (subKey.Values != null)
                    Interlocked.Add(ref _removedValueCount, subKey.Values.Count);
                CountRemovedKeysThreadSafe(subKey);
            }
        }

        private void CountAddedKeys(RegistryKey key)
        {
            if (key.SubKeys == null) return;
            foreach (var subKey in key.SubKeys)
            {
                _addedKeyCount++;
                if (subKey.Values != null)
                    _addedValueCount += subKey.Values.Count;
                CountAddedKeys(subKey);
            }
        }

        private void CountRemovedKeys(RegistryKey key)
        {
            if (key.SubKeys == null) return;
            foreach (var subKey in key.SubKeys)
            {
                _removedKeyCount++;
                if (subKey.Values != null)
                    _removedValueCount += subKey.Values.Count;
                CountRemovedKeys(subKey);
            }
        }

        private List<ValueDiff> CompareValues(IList<KeyValue>? leftValues, IList<KeyValue>? rightValues)
        {
            var result = new List<ValueDiff>();

            // Use GroupBy to handle potential duplicate value names, taking the first occurrence
            var leftDict = new Dictionary<string, KeyValue>(StringComparer.OrdinalIgnoreCase);
            if (leftValues != null)
            {
                foreach (var v in leftValues)
                {
                    var key = v.ValueName ?? "(Default)";
                    if (!leftDict.ContainsKey(key))
                        leftDict[key] = v;
                }
            }

            var rightDict = new Dictionary<string, KeyValue>(StringComparer.OrdinalIgnoreCase);
            if (rightValues != null)
            {
                foreach (var v in rightValues)
                {
                    var key = v.ValueName ?? "(Default)";
                    if (!rightDict.ContainsKey(key))
                        rightDict[key] = v;
                }
            }

            var allValueNames = leftDict.Keys.Union(rightDict.Keys, StringComparer.OrdinalIgnoreCase);

            foreach (var valueName in allValueNames.OrderBy(n => n, StringComparer.OrdinalIgnoreCase))
            {
                leftDict.TryGetValue(valueName, out var leftValue);
                rightDict.TryGetValue(valueName, out var rightValue);

                var diff = new ValueDiff
                {
                    ValueName = valueName,
                    LeftValue = leftValue != null ? FormatValue(leftValue) : null,
                    RightValue = rightValue != null ? FormatValue(rightValue) : null,
                    LeftType = leftValue?.ValueType,
                    RightType = rightValue?.ValueType
                };

                if (leftValue == null && rightValue != null)
                {
                    diff.ChangeType = DiffType.Added;
                    Interlocked.Increment(ref _addedValueCount);
                }
                else if (leftValue != null && rightValue == null)
                {
                    diff.ChangeType = DiffType.Removed;
                    Interlocked.Increment(ref _removedValueCount);
                }
                else if (leftValue != null && rightValue != null)
                {
                    // Compare values
                    if (diff.LeftValue != diff.RightValue || diff.LeftType != diff.RightType)
                    {
                        diff.ChangeType = DiffType.Modified;
                        Interlocked.Increment(ref _modifiedValueCount);
                    }
                    else
                    {
                        diff.ChangeType = DiffType.Unchanged;
                    }
                }

                result.Add(diff);
            }

            return result;
        }

        private string FormatValue(KeyValue value)
        {
            if (value.ValueData == null)
                return "(null)";

            try
            {
                var data = value.ValueData.ToString() ?? "";
                
                // For binary data, show hex representation if available
                if (value.ValueType?.ToUpperInvariant() == "REGBINARY" && value.ValueDataRaw != null && value.ValueDataRaw.Length > 0)
                {
                    var hex = BitConverter.ToString(value.ValueDataRaw.Take(32).ToArray()).Replace("-", " ");
                    if (value.ValueDataRaw.Length > 32)
                        hex += $"... ({value.ValueDataRaw.Length} bytes)";
                    return hex;
                }

                // Truncate very long values
                if (data.Length > 200)
                    return data.Substring(0, 200) + "...";

                return data;
            }
            catch (Exception ex)
            {
                // Log the error for debugging but continue with fallback
                System.Diagnostics.Debug.WriteLine($"FormatValue error: {ex.Message}");
                return value.ValueData?.ToString() ?? "(error)";
            }
        }
    }
}

using System;
using System.Collections.Generic;

namespace RegistryExpert.Core
{
    /// <summary>
    /// Represents a change detected between a hive's current state and its state after
    /// replaying transaction logs.
    /// </summary>
    public class TransactionLogDiff
    {
        public string KeyPath { get; set; } = "";
        public string DisplayPath { get; set; } = "";
        public TransactionLogChangeType ChangeType { get; set; }
        public DateTime? OldTimestamp { get; set; }
        public DateTime? NewTimestamp { get; set; }
        public List<ValueChange> ValueChanges { get; set; } = new();

        /// <summary>
        /// Summary text for the grid column, e.g. "2 values modified" or "DeviceDesc changed"
        /// </summary>
        public string ChangeSummary
        {
            get
            {
                if (ValueChanges.Count == 0)
                {
                    return ChangeType switch
                    {
                        TransactionLogChangeType.KeyAdded => "New key",
                        TransactionLogChangeType.KeyDeleted => "Key removed",
                        _ => ""
                    };
                }
                int added = 0, removed = 0, modified = 0;
                foreach (var vc in ValueChanges)
                {
                    switch (vc.ChangeType)
                    {
                        case ValueChangeType.Added: added++; break;
                        case ValueChangeType.Removed: removed++; break;
                        case ValueChangeType.Modified: modified++; break;
                    }
                }
                var parts = new List<string>();
                if (added > 0) parts.Add($"{added} added");
                if (removed > 0) parts.Add($"{removed} removed");
                if (modified > 0) parts.Add($"{modified} modified");
                return string.Join(", ", parts);
            }
        }
    }

    public enum TransactionLogChangeType
    {
        KeyAdded,
        KeyDeleted,
        ValuesChanged
    }

    public class ValueChange
    {
        public string ValueName { get; set; } = "";
        public string ValueType { get; set; } = "";
        public ValueChangeType ChangeType { get; set; }
        public string? OldData { get; set; }
        public string? NewData { get; set; }
    }

    public enum ValueChangeType
    {
        Added,
        Removed,
        Modified
    }

    /// <summary>
    /// Analyzes registry transaction logs (.LOG1/.LOG2) by replaying them onto a hive
    /// and diffing the before/after states.
    /// </summary>
    public class TransactionLogAnalyzer
    {
        /// <summary>
        /// Detects transaction log files for a given hive file path.
        /// Looks for .LOG1 and .LOG2 files in the same directory with the same base name.
        /// </summary>
        public static List<string> DetectLogFiles(string hiveFilePath)
        {
            var logs = new List<string>();
            if (string.IsNullOrEmpty(hiveFilePath) || !System.IO.File.Exists(hiveFilePath))
                return logs;

            var dir = System.IO.Path.GetDirectoryName(hiveFilePath) ?? "";
            var baseName = System.IO.Path.GetFileName(hiveFilePath);

            var log1 = System.IO.Path.Combine(dir, baseName + ".LOG1");
            var log2 = System.IO.Path.Combine(dir, baseName + ".LOG2");

            if (System.IO.File.Exists(log1)) logs.Add(log1);
            if (System.IO.File.Exists(log2)) logs.Add(log2);

            return logs;
        }

        /// <summary>
        /// Analyzes transaction logs for the given hive, producing a list of diffs.
        /// </summary>
        /// <param name="hiveFilePath">Path to the hive file on disk</param>
        /// <param name="logFilePaths">Paths to .LOG1/.LOG2 files</param>
        /// <param name="progress">Optional progress callback (message string)</param>
        /// <param name="cancellationToken">Cancellation token</param>
        /// <returns>List of diffs between the current hive state and the replayed state</returns>
        public static async System.Threading.Tasks.Task<List<TransactionLogDiff>> AnalyzeAsync(
            string hiveFilePath,
            List<string> logFilePaths,
            Action<string>? progress = null,
            System.Threading.CancellationToken cancellationToken = default)
        {
            return await System.Threading.Tasks.Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                progress?.Invoke("Loading hive file...");

                // Load the hive fresh from disk for the "original" state
                using var originalHive = new RegistryParser.RegistryHive(hiveFilePath);
                originalHive.ParseHive();

                if (originalHive.Root == null)
                    throw new InvalidOperationException("Failed to parse hive: no root key");

                cancellationToken.ThrowIfCancellationRequested();
                progress?.Invoke("Loading transaction logs...");

                // Prepare log file infos
                var logInfos = new List<RegistryParser.TransactionLogFileInfo>();
                foreach (var logPath in logFilePaths)
                {
                    var logBytes = System.IO.File.ReadAllBytes(logPath);
                    if (logBytes.Length > 0)
                        logInfos.Add(new RegistryParser.TransactionLogFileInfo(logPath, logBytes));
                }

                if (logInfos.Count == 0)
                    throw new InvalidOperationException("No valid transaction log files found");

                cancellationToken.ThrowIfCancellationRequested();
                progress?.Invoke("Replaying transaction logs...");

                // Replay transaction logs onto a copy of the hive bytes
                byte[] replayedBytes;
                try
                {
                    replayedBytes = originalHive.ProcessTransactionLogs(logInfos, false);
                }
                catch (Exception ex) when (ex.Message.Contains("Sequence numbers match"))
                {
                    // Hive is not dirty -- try a fallback approach:
                    // Parse the logs independently and apply all entries regardless of sequence check
                    replayedBytes = ReplayLogsFallback(originalHive.FileBytes, logInfos);
                }

                cancellationToken.ThrowIfCancellationRequested();
                progress?.Invoke("Parsing replayed hive...");

                // Parse the replayed hive
                using var replayedHive = new RegistryParser.RegistryHive(replayedBytes, hiveFilePath);
                replayedHive.ParseHive();

                if (replayedHive.Root == null)
                    throw new InvalidOperationException("Failed to parse replayed hive: no root key");

                cancellationToken.ThrowIfCancellationRequested();
                progress?.Invoke("Comparing hive states...");

                // Build lookup dictionaries for both trees
                var originalKeys = new Dictionary<string, RegistryParser.Abstractions.RegistryKey>(StringComparer.OrdinalIgnoreCase);
                BuildKeyMap(originalHive.Root, originalKeys);

                var replayedKeys = new Dictionary<string, RegistryParser.Abstractions.RegistryKey>(StringComparer.OrdinalIgnoreCase);
                BuildKeyMap(replayedHive.Root, replayedKeys);

                cancellationToken.ThrowIfCancellationRequested();
                progress?.Invoke($"Diffing {originalKeys.Count:N0} vs {replayedKeys.Count:N0} keys...");

                var diffs = new List<TransactionLogDiff>();

                // Find keys modified or added in replayed state
                foreach (var kvp in replayedKeys)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var path = kvp.Key;
                    var replayedKey = kvp.Value;

                    if (!originalKeys.TryGetValue(path, out var originalKey))
                    {
                        // Key exists in replayed but not original = recovered/added by logs
                        var diff = new TransactionLogDiff
                        {
                            KeyPath = replayedKey.KeyPath,
                            DisplayPath = path,
                            ChangeType = TransactionLogChangeType.KeyAdded,
                            OldTimestamp = null,
                            NewTimestamp = replayedKey.LastWriteTime?.DateTime
                        };
                        // All values in the replayed key are "added"
                        if (replayedKey.Values != null)
                        {
                            foreach (var val in replayedKey.Values)
                            {
                                diff.ValueChanges.Add(new ValueChange
                                {
                                    ValueName = val.ValueName ?? "(Default)",
                                    ValueType = val.ValueType ?? "",
                                    ChangeType = ValueChangeType.Added,
                                    OldData = null,
                                    NewData = val.ValueData
                                });
                            }
                        }
                        diffs.Add(diff);
                        continue;
                    }

                    // Key exists in both -- check for value differences
                    var valueChanges = DiffValues(originalKey, replayedKey);

                    if (valueChanges.Count > 0)
                    {
                        diffs.Add(new TransactionLogDiff
                        {
                            KeyPath = replayedKey.KeyPath,
                            DisplayPath = path,
                            ChangeType = TransactionLogChangeType.ValuesChanged,
                            OldTimestamp = originalKey.LastWriteTime?.DateTime,
                            NewTimestamp = replayedKey.LastWriteTime?.DateTime,
                            ValueChanges = valueChanges
                        });
                    }
                }

                // Find keys deleted by log replay (exist in original but not replayed)
                foreach (var kvp in originalKeys)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!replayedKeys.ContainsKey(kvp.Key))
                    {
                        var originalKey = kvp.Value;
                        var diff = new TransactionLogDiff
                        {
                            KeyPath = originalKey.KeyPath,
                            DisplayPath = kvp.Key,
                            ChangeType = TransactionLogChangeType.KeyDeleted,
                            OldTimestamp = originalKey.LastWriteTime?.DateTime,
                            NewTimestamp = null
                        };
                        if (originalKey.Values != null)
                        {
                            foreach (var val in originalKey.Values)
                            {
                                diff.ValueChanges.Add(new ValueChange
                                {
                                    ValueName = val.ValueName ?? "(Default)",
                                    ValueType = val.ValueType ?? "",
                                    ChangeType = ValueChangeType.Removed,
                                    OldData = val.ValueData,
                                    NewData = null
                                });
                            }
                        }
                        diffs.Add(diff);
                    }
                }

                progress?.Invoke($"Analysis complete: {diffs.Count:N0} changes found");
                return diffs;
            }, cancellationToken);
        }

        /// <summary>
        /// Fallback: replay logs unconditionally (skipping the dirty-hive check)
        /// by applying all valid HvLE entries directly.
        /// </summary>
        private static byte[] ReplayLogsFallback(byte[] hiveBytes, List<RegistryParser.TransactionLogFileInfo> logInfos)
        {
            var result = new byte[hiveBytes.Length];
            Buffer.BlockCopy(hiveBytes, 0, result, 0, hiveBytes.Length);

            foreach (var logInfo in logInfos)
            {
                try
                {
                    var log = new RegistryParser.TransactionLog(logInfo.FileBytes, logInfo.FileName);
                    log.ParseLog();
                    result = log.UpdateHiveBytes(result);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Fallback log replay failed for {logInfo.FileName}: {ex.Message}");
                }
            }

            return result;
        }

        private static void BuildKeyMap(RegistryParser.Abstractions.RegistryKey key,
            Dictionary<string, RegistryParser.Abstractions.RegistryKey> map)
        {
            map[key.KeyPath] = key;
            if (key.SubKeys != null)
            {
                foreach (var sub in key.SubKeys)
                    BuildKeyMap(sub, map);
            }
        }

        private static List<ValueChange> DiffValues(
            RegistryParser.Abstractions.RegistryKey originalKey,
            RegistryParser.Abstractions.RegistryKey replayedKey)
        {
            var changes = new List<ValueChange>();

            var origValues = new Dictionary<string, RegistryParser.Abstractions.KeyValue>(StringComparer.OrdinalIgnoreCase);
            if (originalKey.Values != null)
            {
                foreach (var v in originalKey.Values)
                    origValues[v.ValueName ?? "(Default)"] = v;
            }

            var replayValues = new Dictionary<string, RegistryParser.Abstractions.KeyValue>(StringComparer.OrdinalIgnoreCase);
            if (replayedKey.Values != null)
            {
                foreach (var v in replayedKey.Values)
                    replayValues[v.ValueName ?? "(Default)"] = v;
            }

            // Check for added/modified values
            foreach (var kvp in replayValues)
            {
                if (!origValues.TryGetValue(kvp.Key, out var origVal))
                {
                    changes.Add(new ValueChange
                    {
                        ValueName = kvp.Key,
                        ValueType = kvp.Value.ValueType ?? "",
                        ChangeType = ValueChangeType.Added,
                        OldData = null,
                        NewData = kvp.Value.ValueData
                    });
                }
                else if (origVal.ValueData != kvp.Value.ValueData || origVal.ValueType != kvp.Value.ValueType)
                {
                    changes.Add(new ValueChange
                    {
                        ValueName = kvp.Key,
                        ValueType = kvp.Value.ValueType ?? "",
                        ChangeType = ValueChangeType.Modified,
                        OldData = origVal.ValueData,
                        NewData = kvp.Value.ValueData
                    });
                }
            }

            // Check for removed values
            foreach (var kvp in origValues)
            {
                if (!replayValues.ContainsKey(kvp.Key))
                {
                    changes.Add(new ValueChange
                    {
                        ValueName = kvp.Key,
                        ValueType = kvp.Value.ValueType ?? "",
                        ChangeType = ValueChangeType.Removed,
                        OldData = kvp.Value.ValueData,
                        NewData = null
                    });
                }
            }

            return changes;
        }
    }
}

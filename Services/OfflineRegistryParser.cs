using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using RegistryParser;
using RegistryParser.Abstractions;

namespace RegistryExpert
{
    /// <summary>
    /// Handles loading and parsing of offline registry hive files
    /// </summary>
    public class OfflineRegistryParser : IDisposable
    {
        private RegistryHive? _hive;
        private string? _filePath;
        private HiveType _hiveType;
        private bool _disposed;

        public bool IsLoaded => _hive != null;
        public string? FilePath => _filePath;
        public HiveType CurrentHiveType => _hiveType;
        public enum HiveType
        {
            Unknown,
            SAM,
            SECURITY,
            SOFTWARE,
            SYSTEM,
            NTUSER,
            USRCLASS,
            DEFAULT,
            AMCACHE,
            BCD,
            COMPONENTS
        }

        /// <summary>
        /// Load a registry hive file
        /// </summary>
        public bool LoadHive(string filePath, IProgress<(string phase, double percent)>? progress = null, CancellationToken cancellationToken = default)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(OfflineRegistryParser));

            try
            {
                // Dispose previous hive if it implements IDisposable
                (_hive as IDisposable)?.Dispose();
                _hive = null;

                _filePath = filePath;
                var newHive = new RegistryHive(filePath);
                
                if (!newHive.ParseHive(progress, cancellationToken))
                {
                    // Dispose the new hive if parsing failed
                    (newHive as IDisposable)?.Dispose();
                    throw new Exception("Failed to parse hive file");
                }

                _hive = newHive;
                _hiveType = DetectHiveType(filePath, _hive);
                return true;
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to load registry hive: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Detect the type of hive based on filename and content
        /// </summary>
        private HiveType DetectHiveType(string filePath, RegistryHive hive)
        {
            string fileName = Path.GetFileName(filePath).ToUpperInvariant();
            
            if (fileName.Contains("SAM")) return HiveType.SAM;
            if (fileName.Contains("SECURITY")) return HiveType.SECURITY;
            if (fileName.Contains("SOFTWARE")) return HiveType.SOFTWARE;
            if (fileName.Contains("SYSTEM")) return HiveType.SYSTEM;
            if (fileName.Contains("NTUSER")) return HiveType.NTUSER;
            if (fileName.Contains("USRCLASS")) return HiveType.USRCLASS;
            if (fileName.Contains("DEFAULT")) return HiveType.DEFAULT;
            if (fileName.Contains("AMCACHE")) return HiveType.AMCACHE;
            if (fileName.Contains("BCD")) return HiveType.BCD;
            if (fileName.Contains("COMPONENTS")) return HiveType.COMPONENTS;

            // Try to detect by content
            try
            {
                var root = hive.Root;
                if (root?.SubKeys != null)
                {
                    var subKeyNames = root.SubKeys.Select(k => k.KeyName.ToUpperInvariant()).ToList();
                    
                    if (subKeyNames.Contains("SAM")) return HiveType.SAM;
                    if (subKeyNames.Contains("POLICY")) return HiveType.SECURITY;
                    if (subKeyNames.Contains("MICROSOFT") && subKeyNames.Contains("CLASSES")) return HiveType.SOFTWARE;
                    if (subKeyNames.Contains("CONTROLSET001") || subKeyNames.Contains("SELECT")) return HiveType.SYSTEM;
                    if (subKeyNames.Contains("SOFTWARE") && subKeyNames.Contains("CONSOLE")) return HiveType.NTUSER;
                    // COMPONENTS hive has DerivedData and/or CanonicalData at root
                    if (subKeyNames.Contains("DERIVEDDATA") || subKeyNames.Contains("CANONICALDATA")) return HiveType.COMPONENTS;
                }
            }
            catch
            {
                // Silently fail content-based detection - filename detection already ran
                // Return Unknown below as fallback
            }

            return HiveType.Unknown;
        }

        /// <summary>
        /// Get the root key of the loaded hive
        /// </summary>
        public RegistryKey? GetRootKey()
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(OfflineRegistryParser));
            return _hive?.Root;
        }

        /// <summary>
        /// Get a specific key by path
        /// </summary>
        public RegistryKey? GetKey(string path)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(OfflineRegistryParser));
            try
            {
                return _hive?.GetKey(path);
            }
            catch
            {
                return null;
            }
        }

        private const int MaxSearchDepth = 100;

        /// <summary>
        /// Search for all matches at the value level. Each matching value gets its own result entry.
        /// Key name matches also produce a separate entry with MatchedValue = null.
        /// </summary>
        public List<SearchMatch> SearchAll(string pattern, bool caseSensitive = false, bool wholeWord = false)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(OfflineRegistryParser));

            var results = new List<SearchMatch>();
            var root = _hive?.Root;

            if (root == null) return results;

            SearchAllRecursive(root, pattern, caseSensitive, wholeWord, results, 0);
            return results;
        }

        private void SearchAllRecursive(RegistryKey key, string pattern, bool caseSensitive, bool wholeWord, List<SearchMatch> results, int depth)
        {
            if (depth > MaxSearchDepth) return;

            try
            {
                var comparison = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

                // Check key name
                bool keyNameMatch = wholeWord
                    ? IsWholeWordMatch(key.KeyName, pattern)
                    : key.KeyName.Contains(pattern, comparison);

                if (keyNameMatch)
                {
                    results.Add(new SearchMatch
                    {
                        Key = key,
                        MatchedValue = null,
                        MatchKind = "Key"
                    });
                }

                // Check each value — emit a separate result for every matching value
                foreach (var value in key.Values)
                {
                    bool nameMatch = wholeWord
                        ? IsWholeWordMatch(value.ValueName, pattern)
                        : value.ValueName.Contains(pattern, comparison);

                    if (nameMatch)
                    {
                        results.Add(new SearchMatch
                        {
                            Key = key,
                            MatchedValue = value,
                            MatchKind = "ValueName"
                        });
                        continue; // Don't double-count if data also matches
                    }

                    var valueData = value.ValueData?.ToString() ?? "";
                    bool dataMatch = wholeWord
                        ? IsWholeWordMatch(valueData, pattern)
                        : valueData.Contains(pattern, comparison);

                    if (dataMatch)
                    {
                        results.Add(new SearchMatch
                        {
                            Key = key,
                            MatchedValue = value,
                            MatchKind = "ValueData"
                        });
                    }
                }

                // Recurse into subkeys
                if (key.SubKeys != null)
                {
                    foreach (var subKey in key.SubKeys)
                    {
                        SearchAllRecursive(subKey, pattern, caseSensitive, wholeWord, results, depth + 1);
                    }
                }
            }
            catch
            {
                // Silently skip keys that fail to enumerate - continue searching remaining keys
            }
        }

        /// <summary>
        /// Checks if searchTerm exists in text as a whole word (bounded by non-word characters or string boundaries).
        /// Always case-insensitive.
        /// </summary>
        private static bool IsWholeWordMatch(string text, string searchTerm)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(searchTerm))
                return false;

            int index = 0;
            while ((index = text.IndexOf(searchTerm, index, StringComparison.OrdinalIgnoreCase)) >= 0)
            {
                bool startBoundary = index == 0 || !char.IsLetterOrDigit(text[index - 1]);
                int endPos = index + searchTerm.Length;
                bool endBoundary = endPos >= text.Length || !char.IsLetterOrDigit(text[endPos]);

                if (startBoundary && endBoundary)
                    return true;

                index += searchTerm.Length;
            }
            return false;
        }

        /// <summary>
        /// Get statistics about the loaded hive
        /// </summary>
        public HiveStatistics GetStatistics()
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(OfflineRegistryParser));
                
            var stats = new HiveStatistics();
            if (_hive?.Root == null) return stats;

            CountRecursive(_hive.Root, stats);
            if (!string.IsNullOrEmpty(_filePath) && File.Exists(_filePath))
                stats.FileSize = new FileInfo(_filePath).Length;
            stats.HiveType = _hiveType.ToString();
            
            return stats;
        }

        private void CountRecursive(RegistryKey key, HiveStatistics stats)
        {
            stats.TotalKeys++;
            stats.TotalValues += key.Values.Count;

            if (key.SubKeys != null)
            {
                foreach (var subKey in key.SubKeys)
                {
                    CountRecursive(subKey, stats);
                }
            }
        }

        /// <summary>
        /// Convert a KeyPath from ROOT\... to HIVENAME\... for display
        /// </summary>
        public string ConvertRootPath(string keyPath)
        {
            if (string.IsNullOrEmpty(keyPath)) return keyPath;
            
            var hiveName = _hiveType.ToString();
            
            if (keyPath.StartsWith("ROOT\\", StringComparison.OrdinalIgnoreCase))
            {
                return hiveName + keyPath.Substring(4); // "ROOT" is 4 chars
            }
            else if (keyPath.Equals("ROOT", StringComparison.OrdinalIgnoreCase))
            {
                return hiveName;
            }
            
            return keyPath;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // Release managed resources - dispose hive if it's disposable
                    (_hive as IDisposable)?.Dispose();
                    _hive = null;
                    _filePath = null;
                }

                _disposed = true;
            }
        }

        ~OfflineRegistryParser()
        {
            Dispose(false);
        }
    }

    /// <summary>
    /// Represents a single search match at the value level.
    /// For key name matches, MatchedValue is null.
    /// </summary>
    public class SearchMatch
    {
        public RegistryKey Key { get; set; } = null!;
        public KeyValue? MatchedValue { get; set; }
        /// <summary>"Key", "ValueName", or "ValueData"</summary>
        public string MatchKind { get; set; } = "";
    }

    public class HiveStatistics
    {
        public int TotalKeys { get; set; }
        public int TotalValues { get; set; }
        public long FileSize { get; set; }
        public string HiveType { get; set; } = "Unknown";

        public string FormattedFileSize
        {
            get
            {
                string[] sizes = { "B", "KB", "MB", "GB" };
                double len = FileSize;
                int order = 0;
                while (len >= 1024 && order < sizes.Length - 1)
                {
                    order++;
                    len /= 1024;
                }
                return $"{len:0.##} {sizes[order]}";
            }
        }
    }
}

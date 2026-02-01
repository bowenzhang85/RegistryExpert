using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Registry;
using Registry.Abstractions;

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
        public RegistryHive? Hive => _hive;

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
        public bool LoadHive(string filePath)
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
                
                if (!newHive.ParseHive())
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

        /// <summary>
        /// Search for keys matching a pattern
        /// </summary>
        public List<RegistryKey> SearchKeys(string pattern, bool caseSensitive = false)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(OfflineRegistryParser));
                
            var results = new List<RegistryKey>();
            var root = _hive?.Root;
            
            if (root == null) return results;

            // Use HashSet for O(1) duplicate checking
            var addedKeys = new HashSet<RegistryKey>();
            SearchKeysRecursive(root, pattern, caseSensitive, results, addedKeys, 0);
            return results;
        }

        private const int MaxSearchDepth = 100;
        
        private void SearchKeysRecursive(RegistryKey key, string pattern, bool caseSensitive, List<RegistryKey> results, HashSet<RegistryKey> addedKeys, int depth)
        {
            // Prevent stack overflow from malformed or deeply nested hives
            if (depth > MaxSearchDepth) return;
            
            try
            {
                var comparison = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
                
                if (key.KeyName.Contains(pattern, comparison))
                {
                    if (addedKeys.Add(key))  // O(1) check and add
                        results.Add(key);
                }

                foreach (var value in key.Values)
                {
                    if (value.ValueName.Contains(pattern, comparison))
                    {
                        if (addedKeys.Add(key))
                            results.Add(key);
                        break;
                    }
                    
                    var valueData = value.ValueData?.ToString() ?? "";
                    if (valueData.Contains(pattern, comparison))
                    {
                        if (addedKeys.Add(key))
                            results.Add(key);
                        break;
                    }
                }

                if (key.SubKeys != null)
                {
                    foreach (var subKey in key.SubKeys)
                    {
                        SearchKeysRecursive(subKey, pattern, caseSensitive, results, addedKeys, depth + 1);
                    }
                }
            }
            catch
            {
                // Silently skip keys that fail to enumerate - continue searching remaining keys
            }
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

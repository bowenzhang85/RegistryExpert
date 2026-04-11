using System.IO;
using System.Text.RegularExpressions;
using RegistryExpert.Core.Models;

namespace RegistryExpert.Core.Services
{
    /// <summary>
    /// Scans a folder for registry hive files.
    /// Supports multi-tier scanning:
    ///   - Fast path for Azure InspectIaaSDisk bundles (goes directly to device_0\Windows\System32\config)
    ///   - Fast path for TSS log bundles (goes directly to Setup_Report-*/ subfolder)
    ///   - Generic recursive scan for any other folder
    /// </summary>
    public class HiveBundleScanner
    {
        // Exact filenames (case-insensitive) that are known registry hives
        private static readonly HashSet<string> KnownHiveNames = new(StringComparer.OrdinalIgnoreCase)
        {
            "SYSTEM", "SOFTWARE", "SAM", "SECURITY", "DEFAULT", "BCD", "COMPONENTS"
        };

        // Filename prefixes that indicate registry hives (e.g., NTUSER.DAT, USRCLASS.DAT, Amcache.hve)
        private static readonly string[] KnownHivePrefixes = { "NTUSER", "USRCLASS", "AMCACHE" };

        // Minimum file size to consider (skip tiny files that can't be valid hives)
        private const long MinHiveFileSize = 4096;

        // InspectIaaSDisk folder name pattern
        private static readonly Regex InspectIaaSDiskPattern = new(
            @"-InspectIaaSDisk-", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// Scans a folder for registry hive files.
        /// </summary>
        /// <param name="folderPath">Root folder to scan.</param>
        /// <param name="progress">Reports status messages during scanning.</param>
        /// <param name="ct">Cancellation token.</param>
        /// <returns>List of discovered hive files.</returns>
        public async Task<List<DiscoveredHive>> ScanFolderAsync(
            string folderPath,
            IProgress<string>? progress = null,
            CancellationToken ct = default)
        {
            return await Task.Run(() => ScanFolder(folderPath, progress, ct), ct)
                .ConfigureAwait(false);
        }

        private List<DiscoveredHive> ScanFolder(
            string folderPath,
            IProgress<string>? progress,
            CancellationToken ct)
        {
            var results = new List<DiscoveredHive>();
            var folderName = Path.GetFileName(folderPath);

            // Fast path: InspectIaaSDisk bundle
            if (InspectIaaSDiskPattern.IsMatch(folderName ?? ""))
            {
                progress?.Report("Detected InspectIaaSDisk bundle...");
                ScanInspectIaaSDiskBundle(folderPath, results, progress, ct);
            }
            // Fast path: TSS log bundle
            else if ((folderName ?? "").StartsWith("TSS_", StringComparison.OrdinalIgnoreCase))
            {
                progress?.Report("Detected TSS log bundle...");
                ScanTssBundle(folderPath, results, progress, ct);
            }
            else
            {
                // Generic recursive scan
                progress?.Report("Scanning folder for registry hives...");
                ScanRecursive(folderPath, folderPath, results, progress, ct);
            }

            return results;
        }

        private void ScanInspectIaaSDiskBundle(
            string rootPath,
            List<DiscoveredHive> results,
            IProgress<string>? progress,
            CancellationToken ct)
        {
            // Known path: device_0\Windows\System32\config
            var configPath = Path.Combine(rootPath, "device_0", "Windows", "System32", "config");

            if (Directory.Exists(configPath))
            {
                progress?.Report("Scanning config folder...");
                ScanDirectory(configPath, rootPath, results, ct);
            }
            else
            {
                // Fallback: if the known path doesn't exist, do a recursive scan
                progress?.Report("InspectIaaSDisk config path not found, scanning recursively...");
                ScanRecursive(rootPath, rootPath, results, progress, ct);
            }
        }

        private void ScanTssBundle(
            string rootPath,
            List<DiscoveredHive> results,
            IProgress<string>? progress,
            CancellationToken ct)
        {
            // TSS bundles store .hiv files inside a Setup_Report-*/ subdirectory
            try
            {
                var setupReportDir = Directory.EnumerateDirectories(rootPath)
                    .FirstOrDefault(d => Path.GetFileName(d)
                        .StartsWith("Setup_Report-", StringComparison.OrdinalIgnoreCase));

                if (setupReportDir != null)
                {
                    progress?.Report($"Scanning {Path.GetFileName(setupReportDir)}...");
                    ScanDirectory(setupReportDir, rootPath, results, ct);
                }

                // If no Setup_Report folder or no hives found there, scan root too
                if (results.Count == 0)
                {
                    progress?.Report("Scanning TSS bundle root...");
                    ScanDirectory(rootPath, rootPath, results, ct);
                }

                // If still nothing found, fallback to recursive scan
                if (results.Count == 0)
                {
                    progress?.Report("No hives in expected TSS paths, scanning recursively...");
                    ScanRecursive(rootPath, rootPath, results, progress, ct);
                }
            }
            catch (UnauthorizedAccessException) { /* skip inaccessible directories */ }
            catch (DirectoryNotFoundException) { /* skip deleted directories */ }
        }

        private const int MaxScanDepth = 15;

        private void ScanRecursive(
            string currentPath,
            string rootPath,
            List<DiscoveredHive> results,
            IProgress<string>? progress,
            CancellationToken ct,
            int depth = 0)
        {
            ct.ThrowIfCancellationRequested();
            if (depth > MaxScanDepth) return; // Safety limit to prevent stack overflow on symlink loops

            ScanDirectory(currentPath, rootPath, results, ct);

            // Recurse into subdirectories
            try
            {
                foreach (var subDir in Directory.EnumerateDirectories(currentPath))
                {
                    ct.ThrowIfCancellationRequested();

                    var dirName = Path.GetFileName(subDir);

                    // Skip well-known non-hive directories to speed up scanning
                    if (dirName.StartsWith('.') ||
                        dirName.Equals("$Recycle.Bin", StringComparison.OrdinalIgnoreCase) ||
                        dirName.Equals("$WinREAgent", StringComparison.OrdinalIgnoreCase))
                        continue;

                    progress?.Report($"Scanning {Path.GetRelativePath(rootPath, subDir)}...");
                    ScanRecursive(subDir, rootPath, results, progress, ct, depth + 1);
                }
            }
            catch (UnauthorizedAccessException) { /* skip inaccessible directories */ }
            catch (DirectoryNotFoundException) { /* skip deleted directories */ }
        }

        private void ScanDirectory(
            string dirPath,
            string rootPath,
            List<DiscoveredHive> results,
            CancellationToken ct)
        {
            try
            {
                foreach (var filePath in Directory.EnumerateFiles(dirPath))
                {
                    ct.ThrowIfCancellationRequested();

                    if (TryMatchHiveFile(filePath, out var hiveType))
                    {
                        // Check minimum file size
                        try
                        {
                            var fileInfo = new FileInfo(filePath);
                            if (fileInfo.Length < MinHiveFileSize) continue;
                        }
                        catch { continue; }

                        // Skip transaction log files
                        var ext = Path.GetExtension(filePath).ToUpperInvariant();
                        if (ext == ".LOG1" || ext == ".LOG2" || ext == ".LOG") continue;

                        results.Add(new DiscoveredHive
                        {
                            FilePath = filePath,
                            DetectedType = hiveType,
                            RelativePath = Path.GetRelativePath(rootPath, filePath)
                        });
                    }
                }
            }
            catch (UnauthorizedAccessException) { /* skip inaccessible directories */ }
            catch (DirectoryNotFoundException) { /* skip deleted directories */ }
        }

        /// <summary>
        /// Checks if a file matches known registry hive filename patterns.
        /// </summary>
        private static bool TryMatchHiveFile(string filePath, out OfflineRegistryParser.HiveType hiveType)
        {
            var fileName = Path.GetFileName(filePath);
            var fileNameUpper = fileName.ToUpperInvariant();
            var nameWithoutExt = Path.GetFileNameWithoutExtension(fileName).ToUpperInvariant();

            // Check exact known names (e.g., SYSTEM, SOFTWARE, SAM)
            if (KnownHiveNames.Contains(nameWithoutExt) || KnownHiveNames.Contains(fileNameUpper))
            {
                hiveType = DetectTypeFromName(fileNameUpper);
                return true;
            }

            // Check known prefixes (e.g., NTUSER.DAT, USRCLASS.DAT, Amcache.hve)
            foreach (var prefix in KnownHivePrefixes)
            {
                if (fileNameUpper.StartsWith(prefix, StringComparison.Ordinal))
                {
                    hiveType = DetectTypeFromName(fileNameUpper);
                    return true;
                }
            }

            // Check .hiv extension
            var ext = Path.GetExtension(filePath);
            if (ext.Equals(".hiv", StringComparison.OrdinalIgnoreCase))
            {
                // Handle TSS pattern: {hostname}_reg_{name}.hiv
                // Extract only the part after "_reg_" to avoid false positives
                // (e.g., "DriverDatabase_System.hiv" should NOT match as SYSTEM)
                var regIndex = nameWithoutExt.IndexOf("_REG_", StringComparison.Ordinal);
                if (regIndex >= 0)
                {
                    var regName = nameWithoutExt.Substring(regIndex + 5);
                    hiveType = KnownHiveNames.Contains(regName)
                        ? DetectTypeFromName(regName)
                        : OfflineRegistryParser.HiveType.Unknown;
                }
                else
                {
                    hiveType = DetectTypeFromName(fileNameUpper);
                }
                return true;
            }

            hiveType = OfflineRegistryParser.HiveType.Unknown;
            return false;
        }

        /// <summary>
        /// Maps a filename to a HiveType. Mirrors the logic in OfflineRegistryParser.DetectHiveType.
        /// </summary>
        private static OfflineRegistryParser.HiveType DetectTypeFromName(string fileNameUpper)
        {
            if (fileNameUpper.Contains("SAM")) return OfflineRegistryParser.HiveType.SAM;
            if (fileNameUpper.Contains("SECURITY")) return OfflineRegistryParser.HiveType.SECURITY;
            if (fileNameUpper.Contains("SOFTWARE")) return OfflineRegistryParser.HiveType.SOFTWARE;
            if (fileNameUpper.Contains("SYSTEM")) return OfflineRegistryParser.HiveType.SYSTEM;
            if (fileNameUpper.Contains("NTUSER")) return OfflineRegistryParser.HiveType.NTUSER;
            if (fileNameUpper.Contains("USRCLASS")) return OfflineRegistryParser.HiveType.USRCLASS;
            if (fileNameUpper.Contains("DEFAULT")) return OfflineRegistryParser.HiveType.DEFAULT;
            if (fileNameUpper.Contains("AMCACHE")) return OfflineRegistryParser.HiveType.AMCACHE;
            if (fileNameUpper.Contains("BCD")) return OfflineRegistryParser.HiveType.BCD;
            if (fileNameUpper.Contains("COMPONENTS")) return OfflineRegistryParser.HiveType.COMPONENTS;
            return OfflineRegistryParser.HiveType.Unknown;
        }

        // ── Recent bundle discovery ─────────────────────────────────────────

        /// <summary>
        /// Scans the user's Downloads folder for IID/TSS bundles modified within the last N days.
        /// Returns an empty list if nothing is found or on any error (never throws).
        /// </summary>
        public static List<BundleInfo> FindRecentBundles(int maxAgeDays = 3)
        {
            var results = new List<BundleInfo>();

            try
            {
                var userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                var downloadsPath = Path.Combine(userProfile, "Downloads");

                if (!Directory.Exists(downloadsPath))
                    return results;

                var cutoff = DateTime.Now.AddDays(-maxAgeDays);

                foreach (var dir in Directory.EnumerateDirectories(downloadsPath))
                {
                    try
                    {
                        var folderName = Path.GetFileName(dir);
                        if (string.IsNullOrEmpty(folderName))
                            continue;

                        var lastWrite = Directory.GetLastWriteTime(dir);
                        if (lastWrite < cutoff)
                            continue;

                        string? bundleType = null;

                        if (InspectIaaSDiskPattern.IsMatch(folderName))
                            bundleType = "InspectIaaSDisk";
                        else if (folderName.StartsWith("TSS_", StringComparison.OrdinalIgnoreCase))
                            bundleType = "TSS";

                        if (bundleType != null)
                        {
                            results.Add(new BundleInfo
                            {
                                FolderPath = dir,
                                Name = folderName,
                                BundleType = bundleType,
                                ModifiedDate = lastWrite
                            });
                        }
                    }
                    catch
                    {
                        // Skip inaccessible folders
                    }
                }

                results.Sort((a, b) => b.ModifiedDate.CompareTo(a.ModifiedDate));
            }
            catch
            {
                // Never throw — return empty list
            }

            return results;
        }
    }
}

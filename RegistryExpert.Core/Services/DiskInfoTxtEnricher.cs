using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using RegistryExpert.Core.Models;

namespace RegistryExpert.Core.Services
{
    /// <summary>
    /// Detects and applies Azure InspectIaaSDisk bundle enrichment to a
    /// <see cref="DiskLayoutModel"/>. Microsoft's InspectIaaSDisk tool mounts a
    /// customer's Windows OS VHD on a Linux helper VM and dumps <c>df</c> /
    /// <c>statvfs</c> output to a <c>diskinfo.txt</c> file in the bundle root.
    /// When the loaded SYSTEM hive lives inside such a bundle, we can recover
    /// authoritative Capacity / Free / Used values for the OS volume — values
    /// otherwise impossible to determine from offline registry data.
    /// </summary>
    public static class DiskInfoTxtEnricher
    {
        /// <summary>
        /// Walk up parent directories from <paramref name="hivePath"/> looking
        /// for a folder whose name contains <c>InspectIaaSDisk</c> AND which
        /// contains a <c>diskinfo.txt</c> file. Returns the diskinfo.txt path
        /// when found, or null otherwise.
        /// </summary>
        public static string? DetectBundleDiskInfo(string hivePath)
        {
            try
            {
                var dir = Path.GetDirectoryName(hivePath);
                while (!string.IsNullOrEmpty(dir))
                {
                    var name = Path.GetFileName(dir);
                    if (!string.IsNullOrEmpty(name) &&
                        name.IndexOf("InspectIaaSDisk", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        var candidate = Path.Combine(dir, "diskinfo.txt");
                        if (File.Exists(candidate))
                            return candidate;
                    }
                    dir = Path.GetDirectoryName(dir);
                }
            }
            catch
            {
                // Defensive — never let path issues break enrichment
            }
            return null;
        }

        /// <summary>
        /// Parse <c>diskinfo.txt</c> at <paramref name="filePath"/> into one or
        /// more <see cref="DiskInfoTxtData"/> records (typically one per Windows
        /// drive letter present in the dump, usually just C:).
        /// </summary>
        public static List<DiskInfoTxtData> Parse(string filePath)
        {
            var results = new List<DiskInfoTxtData>();
            if (!File.Exists(filePath)) return results;

            string text;
            try { text = File.ReadAllText(filePath); }
            catch { return results; }

            var fileTime = File.GetLastWriteTimeUtc(filePath);

            // ── 1. Find Windows-drive-letter ↔ Linux-device mappings ─────────
            // Example line:  "C: /dev/sda4"
            var letterMappings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var letterRegex = new Regex(@"^\s*([A-Z]):\s+(/dev/\S+)\s*$",
                RegexOptions.IgnoreCase | RegexOptions.Multiline);
            foreach (Match m in letterRegex.Matches(text))
            {
                var letter = m.Groups[1].Value.ToUpperInvariant() + ":";
                var device = m.Groups[2].Value;
                if (!letterMappings.ContainsKey(letter))
                    letterMappings[letter] = device;
            }

            // ── 2. Find df rows with /sysroot mountpoint — that's the Windows volume
            //    Format:  "/dev/sda4       127G   63G   65G  50% /sysroot"
            var dfRowsByDevice = new Dictionary<string, (string SizeH, string UsedH, string AvailH, string UsePct)>(
                StringComparer.OrdinalIgnoreCase);
            var dfRegex = new Regex(
                @"^\s*(?<dev>/dev/\S+)\s+(?<size>\S+)\s+(?<used>\S+)\s+(?<avail>\S+)\s+(?<pct>\S+%)\s+(?<mp>/\S+)\s*$",
                RegexOptions.Multiline);
            foreach (Match m in dfRegex.Matches(text))
            {
                var mp = m.Groups["mp"].Value;
                if (!mp.Equals("/sysroot", StringComparison.OrdinalIgnoreCase)) continue;
                var dev = m.Groups["dev"].Value;
                if (!dfRowsByDevice.ContainsKey(dev))
                {
                    dfRowsByDevice[dev] = (
                        m.Groups["size"].Value,
                        m.Groups["used"].Value,
                        m.Groups["avail"].Value,
                        m.Groups["pct"].Value);
                }
            }

            // ── 3. Find statvfs blocks ───────────────────────────────────────
            //    Format:
            //      [Device: /dev/sda4, mountpoint: / ]
            //      bsize: 4096
            //      frsize: 4096
            //      blocks: 33147379
            //      bfree: 16811096   ← raw free blocks (includes reserved)
            //      bavail: 16811096  ← user-accessible free blocks (excludes reserved)
            //      ...
            //
            // We use `bavail` for the user-facing "Free Space" because it matches
            // Windows' notion of free space (Windows excludes filesystem-reserved
            // blocks from free-space reports). Falls back to `bfree` if `bavail`
            // is missing for any reason.
            var statvfsBlockRegex = new Regex(
                @"\[Device:\s*(?<dev>/dev/\S+?),\s*mountpoint:\s*(?<mp>/\S*)\s*\](?<body>.*?)(?=\[Device:|\z)",
                RegexOptions.Singleline);
            foreach (Match block in statvfsBlockRegex.Matches(text))
            {
                var dev = block.Groups["dev"].Value;
                var body = block.Groups["body"].Value;

                int? bsize = ParseStatvfsInt(body, "bsize");
                long? blocks = ParseStatvfsLong(body, "blocks");
                long? bavail = ParseStatvfsLong(body, "bavail");
                long? bfree = ParseStatvfsLong(body, "bfree");
                long? freeBlocks = bavail ?? bfree;
                long? files = ParseStatvfsLong(body, "files");
                long? ffree = ParseStatvfsLong(body, "ffree");

                // Find which Windows drive letter (if any) corresponds to this Linux device
                string? matchedLetter = letterMappings
                    .FirstOrDefault(kvp => string.Equals(kvp.Value, dev, StringComparison.OrdinalIgnoreCase))
                    .Key;

                // Even if there's no matching letter mapping, we may still want to record
                // the statvfs data for diagnostics. Skip silently if essential fields missing.
                if (!bsize.HasValue || !blocks.HasValue) continue;

                var record = new DiskInfoTxtData
                {
                    WindowsDriveLetter = matchedLetter ?? "",
                    LinuxDevice = dev,
                    ClusterSizeBytes = bsize,
                    TotalBytes = blocks.Value * bsize.Value,
                    FreeBytes = freeBlocks.HasValue ? freeBlocks.Value * bsize.Value : null,
                    TotalInodes = files,
                    FreeInodes = ffree,
                    ExtractedAt = fileTime,
                    SourcePath = filePath,
                };

                if (record.TotalBytes.HasValue && record.FreeBytes.HasValue)
                    record.UsedBytes = record.TotalBytes.Value - record.FreeBytes.Value;

                // Derive disk and partition numbers from the Linux device name
                // (e.g. "/dev/sda4" → disk 0, partition 4; "/dev/sdb12" → disk 1, partition 12)
                var devName = dev.StartsWith("/dev/", StringComparison.OrdinalIgnoreCase)
                    ? dev.Substring(5)
                    : dev;
                if (devName.StartsWith("sd", StringComparison.OrdinalIgnoreCase) && devName.Length >= 3)
                {
                    char diskChar = char.ToLowerInvariant(devName[2]);
                    if (diskChar >= 'a' && diskChar <= 'z')
                        record.DiskNumber = diskChar - 'a';

                    var trailing = new string(devName.SkipWhile(c => !char.IsDigit(c)).ToArray());
                    if (int.TryParse(trailing, NumberStyles.Integer, CultureInfo.InvariantCulture, out int pn))
                        record.PartitionNumber = pn;
                }

                // Filesystem hint — bsize=4096 is the default NTFS cluster size; flag accordingly
                if (bsize.Value == 4096)
                    record.FilesystemHint = "NTFS (likely)";
                else if (bsize.Value == 65536)
                    record.FilesystemHint = "ReFS (likely)";
                else
                    record.FilesystemHint = $"Unknown (cluster {bsize.Value})";

                results.Add(record);
            }

            return results;
        }

        /// <summary>
        /// Apply parsed diskinfo.txt records to a <see cref="DiskLayoutModel"/>:
        /// for each record with a Windows drive letter, find the matching
        /// partition and populate its authoritative Capacity / Free / Used fields.
        /// </summary>
        /// <returns>Number of partitions that received enrichment.</returns>
        public static int Enrich(DiskLayoutModel model, IEnumerable<DiskInfoTxtData> records)
        {
            int enrichedCount = 0;
            foreach (var rec in records)
            {
                if (string.IsNullOrEmpty(rec.WindowsDriveLetter)) continue;

                var partition = model.Disks
                    .SelectMany(d => d.Partitions)
                    .FirstOrDefault(p => string.Equals(p.DriveLetter, rec.WindowsDriveLetter,
                        StringComparison.OrdinalIgnoreCase));
                if (partition == null) continue;

                if (rec.TotalBytes.HasValue)
                {
                    partition.EstimatedLengthBytes = rec.TotalBytes;
                    partition.LengthIsEstimated = false;
                    partition.CapacityFromExternalSource = true;
                }
                partition.FreeSpaceBytes = rec.FreeBytes;
                partition.UsedSpaceBytes = rec.UsedBytes;
                partition.ClusterSizeBytes = rec.ClusterSizeBytes;
                partition.TotalInodes = rec.TotalInodes;
                partition.FreeInodes = rec.FreeInodes;

                // Upgrade filesystem inference when hint is available
                if (!string.IsNullOrEmpty(rec.FilesystemHint))
                {
                    // Replace boot-volume's inferred "NTFS" with the cluster-size-based
                    // hint if the hint is more specific than what we already have.
                    if (string.IsNullOrEmpty(partition.FilesystemType) ||
                        partition.FilesystemIsInferred)
                    {
                        // Strip parenthetical for clean display: "NTFS (likely)" → "NTFS"
                        var fs = rec.FilesystemHint;
                        var parenIdx = fs.IndexOf('(');
                        if (parenIdx > 0)
                            fs = fs.Substring(0, parenIdx).Trim();
                        if (!fs.StartsWith("Unknown", StringComparison.OrdinalIgnoreCase))
                        {
                            partition.FilesystemType = fs;
                            partition.FilesystemIsInferred = true; // still inferred (cluster heuristic)
                        }
                    }
                }

                partition.RawRegistryLocations["diskinfo.txt"] =
                    $"{rec.SourcePath} — {rec.WindowsDriveLetter} → {rec.LinuxDevice}, " +
                    $"{FormatBytes(rec.TotalBytes)} total, {FormatBytes(rec.FreeBytes)} free";

                enrichedCount++;
            }

            if (enrichedCount > 0)
                model.Sources |= DiskLayoutSourceFlags.DiskInfoTxt;

            return enrichedCount;
        }

        // ── helpers ──────────────────────────────────────────────────────────

        private static int? ParseStatvfsInt(string body, string field)
        {
            var rx = new Regex($@"^\s*{Regex.Escape(field)}:\s*(\d+)\s*$", RegexOptions.Multiline);
            var m = rx.Match(body);
            if (m.Success && int.TryParse(m.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int v))
                return v;
            return null;
        }

        private static long? ParseStatvfsLong(string body, string field)
        {
            var rx = new Regex($@"^\s*{Regex.Escape(field)}:\s*(\d+)\s*$", RegexOptions.Multiline);
            var m = rx.Match(body);
            if (m.Success && long.TryParse(m.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long v))
                return v;
            return null;
        }

        private static string FormatBytes(long? bytes)
        {
            if (!bytes.HasValue) return "?";
            double v = bytes.Value;
            string[] units = { "B", "KB", "MB", "GB", "TB", "PB" };
            int i = 0;
            while (v >= 1024 && i < units.Length - 1)
            {
                v /= 1024;
                i++;
            }
            return $"{v:F2} {units[i]}";
        }
    }
}

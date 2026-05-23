using System;
using System.Collections.Generic;
using System.Linq;
using RegistryExpert.Core.Models;

namespace RegistryExpert.Core.Services
{
    /// <summary>
    /// Applies BCD-derived Boot/System role tagging to a <see cref="DiskLayoutModel"/>.
    /// Replaces the heuristic "C: is Boot+System" assumption with authoritative
    /// information from the BCD store when available.
    /// </summary>
    public static class DiskLayoutBcdEnricher
    {
        /// <summary>
        /// Walk parent directories from <paramref name="hivePath"/> looking for a
        /// <c>Boot/BCD</c> sibling. Returns the BCD path when found, else null.
        /// </summary>
        /// <remarks>
        /// Per OQ-Phase-A-1 answer, this is *not* called automatically — BCD must
        /// be explicitly loaded by the user. This helper exists for the test
        /// harness and for the (deferred) F1 multi-hive correlation feature.
        /// </remarks>
        public static string? DiscoverBcdNearHive(string hivePath)
        {
            try
            {
                var dir = System.IO.Path.GetDirectoryName(hivePath);
                while (!string.IsNullOrEmpty(dir))
                {
                    // SYSTEM is at "...\Windows\System32\config\SYSTEM"
                    // BCD is at  "...\Boot\BCD" — siblings of "Windows" folder.
                    var bootDir = System.IO.Path.Combine(dir, "Boot");
                    var bcd = System.IO.Path.Combine(bootDir, "BCD");
                    if (System.IO.File.Exists(bcd))
                        return bcd;
                    dir = System.IO.Path.GetDirectoryName(dir);
                }
            }
            catch { }
            return null;
        }

        /// <summary>
        /// Apply BCD-derived roles to the model:
        ///  - Match the active Boot Loader's <c>osdevice</c> against model partitions → set <see cref="PartitionRoleFlags.Boot"/>.
        ///  - Match the active Boot Loader's <c>device</c> against model partitions → set <see cref="PartitionRoleFlags.System"/>.
        ///  - Clear the heuristic Boot/System assignment from partitions that don't match BCD evidence.
        /// </summary>
        /// <returns>Number of partitions that received an authoritative role from BCD.</returns>
        public static int Enrich(DiskLayoutModel model, IEnumerable<BcdBootEntry> entries)
        {
            int taggedCount = 0;
            bool sawAuthoritative = false;

            // First pass: collect all targets BCD references so we know which
            // heuristic assignments should be reverted (any partition currently
            // tagged Boot/System but not referenced by BCD).
            //
            // We treat Boot Loader's osdevice as primary Boot evidence. When the
            // Boot Loader's descriptor is empty (common when it inherits from
            // unresolved objects), we fall back to the Resume Loader's osdevice
            // — winresume.exe points at the same OS partition.
            var bcdBootTargets = new List<BcdDeviceDescriptor>();
            var bcdSystemTargets = new List<BcdDeviceDescriptor>();

            foreach (var entry in entries)
            {
                if (entry.Type == BcdParser.TypeBootLoader)
                {
                    if (entry.OsDevice?.HasUsableTarget == true)
                    {
                        bcdBootTargets.Add(entry.OsDevice);
                        sawAuthoritative = true;
                    }
                    if (entry.Device?.HasUsableTarget == true)
                    {
                        bcdSystemTargets.Add(entry.Device);
                        sawAuthoritative = true;
                    }
                }
            }

            // Fallback to Resume Loader when the active Boot Loaders provided nothing
            if (bcdBootTargets.Count == 0)
            {
                foreach (var entry in entries)
                {
                    if (entry.Type == BcdParser.TypeResumeLoader &&
                        entry.OsDevice?.HasUsableTarget == true)
                    {
                        bcdBootTargets.Add(entry.OsDevice);
                        sawAuthoritative = true;
                    }
                }
            }

            // If BCD has nothing usable (e.g. all-zero descriptors only — common when
            // the loader inherits from objects we couldn't resolve), leave the
            // heuristic in place and report 0.
            if (!sawAuthoritative) return 0;

            // Build a fast lookup: (sig, offset) → partition for MBR; partition GUID → partition for GPT.
            var partitions = model.Disks.SelectMany(d => d.Partitions).ToList();

            foreach (var target in bcdBootTargets)
            {
                var p = FindMatch(partitions, target);
                if (p == null) continue;
                p.Roles |= PartitionRoleFlags.Boot;
                p.RawRegistryLocations["BCD osdevice"] = "BCD: Boot Loader's osdevice points here (Windows installation)";
                if (string.IsNullOrEmpty(p.FilesystemType))
                {
                    p.FilesystemType = "NTFS";
                    p.FilesystemIsInferred = true;
                }
                taggedCount++;
            }

            foreach (var target in bcdSystemTargets)
            {
                var p = FindMatch(partitions, target);
                if (p == null) continue;
                p.Roles |= PartitionRoleFlags.System;
                p.RawRegistryLocations["BCD device"] = "BCD: Boot Loader's device points here (boot loader location)";
                taggedCount++;
            }

            // Clear heuristic Boot/System tags on partitions BCD didn't point at —
            // BUT ONLY when at least one BCD match actually succeeded. If BCD
            // produced "authoritative" descriptors that didn't match anything in
            // the model (e.g. GPT systems where our GPT-GUID matching is structurally
            // limited — see TODO below), we'd otherwise wipe out the heuristic-assigned
            // C: roles and leave the user with NO Boot/System tags at all.
            //
            // TODO Phase B: properly match BCD's per-partition unique GUID against
            // the partition's unique GUID (currently we compare against the partition
            // TYPE GUID which is a different concept). When that lands, this guard
            // can be removed.
            if (taggedCount > 0)
            {
                foreach (var p in partitions)
                {
                    bool bcdSaysBoot = bcdBootTargets.Any(t => Matches(p, t));
                    bool bcdSaysSystem = bcdSystemTargets.Any(t => Matches(p, t));
                    if (!bcdSaysBoot && (p.Roles & PartitionRoleFlags.Boot) != 0)
                    {
                        p.Roles &= ~PartitionRoleFlags.Boot;
                    }
                    if (!bcdSaysSystem && (p.Roles & PartitionRoleFlags.System) != 0)
                    {
                        p.Roles &= ~PartitionRoleFlags.System;
                    }
                }
            }

            if (taggedCount > 0)
                model.Sources |= DiskLayoutSourceFlags.Bcd;

            return taggedCount;
        }

        private static DiskLayoutPartition? FindMatch(
            List<DiskLayoutPartition> partitions,
            BcdDeviceDescriptor target)
        {
            foreach (var p in partitions)
            {
                if (Matches(p, target)) return p;
            }
            return null;
        }

        private static bool Matches(DiskLayoutPartition p, BcdDeviceDescriptor target)
        {
            // MBR match: signature equals
            if (target.MbrDiskSignature.HasValue && p.MbrDiskSignature.HasValue)
            {
                if (target.MbrDiskSignature.Value == p.MbrDiskSignature.Value)
                    return true;
            }
            // GPT match: partition GUID equals (when both sides have it)
            if (!string.IsNullOrEmpty(target.GptPartitionGuid) &&
                !string.IsNullOrEmpty(p.GptPartitionTypeGuid))
            {
                if (string.Equals(target.GptPartitionGuid, p.GptPartitionTypeGuid,
                    StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }
    }
}

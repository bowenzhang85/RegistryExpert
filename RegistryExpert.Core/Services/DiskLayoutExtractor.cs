using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using RegistryExpert.Core.Models;
using RegistryParser.Abstractions;

namespace RegistryExpert.Core.Services
{
    /// <summary>
    /// Builds a <see cref="DiskLayoutModel"/> from a loaded SYSTEM hive. Reconstructs a
    /// diskmgmt.msc-style view: enumerates physical disks, walks their partitions via
    /// <c>Enum\STORAGE\Volume</c>, estimates partition lengths by offset subtraction,
    /// cross-references MountedDevices to assign drive letters, and infers per-partition
    /// roles (Boot, System, Pagefile, CrashDump, ESP, MSR, Recovery, Temp).
    /// </summary>
    /// <remarks>
    /// SYSTEM-only by design. BCD-based role enrichment and InspectIaaSDisk diskinfo.txt
    /// enrichment are layered on by separate services (DiskLayoutBcdEnricher,
    /// DiskInfoTxtEnricher) so the extractor degrades gracefully when only SYSTEM is
    /// available.
    /// </remarks>
    public class DiskLayoutExtractor
    {
        private readonly OfflineRegistryParser _parser;

        // Reuse the existing extractor for MountedDevices parsing (already battle-tested)
        private readonly RegistryInfoExtractor _registryInfo;

        // Well-known DEVPROPKEY GUIDs (Plug-and-Play device properties)
        // Property key format under Enum: ...\Properties\{GUID}\{HexId}
        // For the data type prefix, see https://learn.microsoft.com/en-us/windows-hardware/drivers/install/devpropkey
        private const string DevPkeyDeviceTimestamps  = "{83da6326-97a6-4088-9453-a1923f573b29}";
        private const string DevPkeyDeviceFriendlyEtc = "{a8b865dd-2e3d-4094-ad97-e593a70c75d6}";
        private const string DevPkeyDeviceLocationPath = "{a45c254e-df1c-4efd-8020-67d146a850e0}";
        private const string DevPkeyDevicePartitionInfo = "{540b947e-8b40-45bc-a8a2-6a0b894cbda2}";
        private const string DevPkeyDeviceGptInfo      = "{3464f7a4-2444-40b1-980a-e0903cb6d912}";

        // Property IDs within DevPkeyDeviceTimestamps
        private const string PidInstalledTimestamp  = "0064";
        private const string PidArrivalTimestamp    = "0066";

        // Property IDs within DevPkeyDeviceFriendlyEtc
        private const string PidFriendlyName        = "000A"; // DEVPKEY_NAME (friendly)
        private const string PidManufacturer        = "000D";

        // Property IDs within DevPkeyDeviceLocationPath
        private const string PidLocationPath        = "0025";

        // Known GPT partition type GUIDs
        private static readonly Dictionary<string, PartitionRoleFlags> GptTypeRoles = new(StringComparer.OrdinalIgnoreCase)
        {
            ["{c12a7328-f81f-11d2-ba4b-00a0c93ec93b}"] = PartitionRoleFlags.ESP,
            ["{e3c9e316-0b5c-4db8-817d-f92df00215ae}"] = PartitionRoleFlags.MSR,
            ["{de94bba4-06d1-4d40-a16a-bfd50179d6ac}"] = PartitionRoleFlags.Recovery,
        };

        public DiskLayoutExtractor(OfflineRegistryParser parser)
        {
            _parser = parser ?? throw new ArgumentNullException(nameof(parser));
            _registryInfo = new RegistryInfoExtractor(parser);
        }

        /// <summary>
        /// Build the complete <see cref="DiskLayoutModel"/> from the loaded SYSTEM hive.
        /// Safe to call when the hive is partial or missing expected keys — the model
        /// will reflect only what could be recovered and record diagnostics for gaps.
        /// </summary>
        public DiskLayoutModel BuildModel()
        {
            var model = new DiskLayoutModel
            {
                Sources = DiskLayoutSourceFlags.System,
                ExtractedAt = DateTime.UtcNow,
            };

            // ── 1. Enumerate physical disks ──────────────────────────────────
            try
            {
                EnumerateDisks(model);
            }
            catch (Exception ex)
            {
                model.Diagnostics.Add($"Disk enumeration failed: {ex.Message}");
                Debug.WriteLine($"DiskLayoutExtractor disk enum: {ex}");
            }

            // ── 2. Enumerate partitions via STORAGE\Volume ───────────────────
            //    Each registered volume goes either to a parent disk's Partitions
            //    list (matched by DiskId) or to the OrphanPartitions list.
            try
            {
                EnumeratePartitions(model);
            }
            catch (Exception ex)
            {
                model.Diagnostics.Add($"Partition enumeration failed: {ex.Message}");
                Debug.WriteLine($"DiskLayoutExtractor partition enum: {ex}");
            }

            // ── 3. Sort partitions on each disk and compute estimated lengths
            foreach (var disk in model.Disks)
            {
                disk.Partitions.Sort((a, b) => a.PartitionOffsetBytes.CompareTo(b.PartitionOffsetBytes));
                ComputeEstimatedLengths(disk);
            }

            // ── 4. Cross-reference MountedDevices for drive letters ──────────
            try
            {
                ApplyMountedDevicesCrossRef(model);
            }
            catch (Exception ex)
            {
                model.Diagnostics.Add($"MountedDevices cross-reference failed: {ex.Message}");
                Debug.WriteLine($"DiskLayoutExtractor mountdev xref: {ex}");
            }

            // ── 4.5. Filter stale STORAGE\Volume registrations ───────────────
            //    Heavy hives (e.g. long-lived VMware-on-Azure machines) accumulate
            //    historical STORAGE\Volume entries from past partitioning operations.
            //    Each repartition creates a new entry; Windows never garbage-collects
            //    them. A registration is "active" only if it has a MountedDevices
            //    counterpart. The rest are residue and would clutter the UI.
            try
            {
                FilterStaleRegistrations(model);

                // Re-compute estimated lengths because some partitions' "next neighbour"
                // may have been removed, changing the estimate.
                foreach (var disk in model.Disks)
                {
                    // Reset any prior estimates so the recomputation isn't biased
                    foreach (var p in disk.Partitions)
                    {
                        if (p.LengthIsEstimated)
                        {
                            p.EstimatedLengthBytes = null;
                            p.LengthIsEstimated = false;
                        }
                    }
                    ComputeEstimatedLengths(disk);
                }
            }
            catch (Exception ex)
            {
                model.Diagnostics.Add($"Stale registration filter failed: {ex.Message}");
                Debug.WriteLine($"DiskLayoutExtractor stale filter: {ex}");
            }

            // ── 5. Infer partition roles ─────────────────────────────────────
            try
            {
                InferPartitionRoles(model);
            }
            catch (Exception ex)
            {
                model.Diagnostics.Add($"Role inference failed: {ex.Message}");
                Debug.WriteLine($"DiskLayoutExtractor role inference: {ex}");
            }

            // ── 6. Roll up disk-level roles from partition roles ─────────────
            foreach (var disk in model.Disks)
            {
                ComputeDiskRoles(disk);
            }

            // ── 7. Compute partition / disk statuses ─────────────────────────
            ComputeStatuses(model);

            // ── 8. Renumber disks (synthetic DiskNumber, 0-based, by bus order)
            RenumberDisks(model);

            // ── 9. Build RawRegistryLocations strings (final pass) ───────────
            BuildRawLocations(model);

            return model;
        }

        // ════════════════════════════════════════════════════════════════════
        // Disk enumeration
        // ════════════════════════════════════════════════════════════════════

        private void EnumerateDisks(DiskLayoutModel model)
        {
            // Walk Enum\IDE and Enum\SCSI for physical disk-class devices.
            // Note: Enum\STORAGE\Disk is intentionally skipped — those entries are
            // PnP re-projections of the underlying IDE/SCSI disks (used for filter
            // driver attachment) and would double-count if enumerated.
            EnumerateDiskBus(model, "ControlSet001\\Enum\\IDE",  "IDE");
            EnumerateDiskBus(model, "ControlSet001\\Enum\\SCSI", "SCSI");

            if (model.Disks.Count == 0)
            {
                model.Diagnostics.Add("No disks found under ControlSet001\\Enum\\{IDE,SCSI}");
            }
        }

        /// <summary>
        /// Walk a bus key (e.g. Enum\IDE) and create a DiskLayoutDisk for each
        /// instance under each device class. Skips CdRom and floppy classes.
        /// </summary>
        private void EnumerateDiskBus(DiskLayoutModel model, string busKeyPath, string busType)
        {
            var busKey = _parser.GetKey(busKeyPath);
            if (busKey == null) return;

            foreach (var deviceClass in busKey.SubKeys)
            {
                // Skip CdRom / floppy / non-disk classes
                var className = deviceClass.KeyName;
                if (IsCdRomClass(className) || IsFloppyClass(className))
                    continue;

                // For STORAGE\Disk, the immediate children ARE the instances.
                // For IDE/SCSI, deviceClass children are instances.
                foreach (var instance in deviceClass.SubKeys)
                {
                    var disk = TryBuildDisk(instance, busType, deviceClass.KeyName);
                    if (disk != null)
                        model.Disks.Add(disk);
                }
            }
        }

        private static bool IsCdRomClass(string className) =>
            className.IndexOf("CdRom", StringComparison.OrdinalIgnoreCase) >= 0;

        private static bool IsFloppyClass(string className) =>
            className.IndexOf("Floppy", StringComparison.OrdinalIgnoreCase) >= 0 ||
            className.StartsWith("FDC", StringComparison.OrdinalIgnoreCase);

        private DiskLayoutDisk? TryBuildDisk(RegistryKey instanceKey, string busType, string deviceClassName)
        {
            var disk = new DiskLayoutDisk
            {
                BusType = busType,
                FriendlyName = TrimHardwareDescription(deviceClassName),
                EnumKeyPath = instanceKey.KeyPath,
            };

            // Read DiskId from Device Parameters\Partmgr
            var partmgrKey = FindSubkey(instanceKey, "Device Parameters", "Partmgr");
            if (partmgrKey != null)
            {
                var diskIdVal = partmgrKey.GetValue("DiskId") as string;
                if (!string.IsNullOrEmpty(diskIdVal))
                {
                    disk.DiskId = NormalizeGuid(diskIdVal);
                }
            }

            // Bus location
            var locInfo = instanceKey.GetValue("LocationInformation") as string;
            if (!string.IsNullOrEmpty(locInfo))
            {
                disk.BusLocation = locInfo;
            }

            // Manufacturer (Mfg)
            var mfg = instanceKey.GetValue("Mfg") as string;
            if (!string.IsNullOrEmpty(mfg))
            {
                disk.Manufacturer = StripInfPrefix(mfg);
            }

            // FriendlyName overrides class-derived name
            var fnRaw = instanceKey.GetValue("FriendlyName") as string;
            if (!string.IsNullOrEmpty(fnRaw))
            {
                disk.FriendlyName = StripInfPrefix(fnRaw);
            }

            // Properties dictionary timestamps + ACPI path
            var propsKey = FindSubkey(instanceKey, "Properties");
            if (propsKey != null)
            {
                disk.InstalledAt = ReadPropertyAsFileTime(propsKey, DevPkeyDeviceTimestamps, PidInstalledTimestamp);
                disk.LastArrivalAt = ReadPropertyAsFileTime(propsKey, DevPkeyDeviceTimestamps, PidArrivalTimestamp);
                disk.AcpiPath = ReadPropertyAsString(propsKey, DevPkeyDeviceLocationPath, PidLocationPath);
            }

            return disk;
        }

        // ════════════════════════════════════════════════════════════════════
        // Partition enumeration
        // ════════════════════════════════════════════════════════════════════

        private void EnumeratePartitions(DiskLayoutModel model)
        {
            var storageVolumeKey = _parser.GetKey("ControlSet001\\Enum\\STORAGE\\Volume");
            if (storageVolumeKey == null)
            {
                model.Diagnostics.Add("ControlSet001\\Enum\\STORAGE\\Volume not present — no partitions can be enumerated");
                return;
            }

            // Build index of DiskId → DiskLayoutDisk for O(1) lookup
            var diskById = new Dictionary<string, DiskLayoutDisk>(StringComparer.OrdinalIgnoreCase);
            foreach (var d in model.Disks)
            {
                if (!string.IsNullOrEmpty(d.DiskId))
                    diskById[d.DiskId] = d;
            }

            foreach (var volKey in storageVolumeKey.SubKeys)
            {
                // Key name format: {DiskIdGuid}#{HexOffset}
                // e.g. {0f8bc731-d943-11e7-a93c-806e6f6e6963}#0000000000100000
                var keyName = volKey.KeyName;
                var hashIdx = keyName.IndexOf('#');
                if (hashIdx < 1 || hashIdx == keyName.Length - 1)
                    continue;

                var diskIdToken = keyName.Substring(0, hashIdx).Trim();
                var offsetToken = keyName.Substring(hashIdx + 1).Trim();

                var normalizedDiskId = NormalizeGuid(diskIdToken);
                if (!long.TryParse(offsetToken, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out long offset))
                    continue;

                var partition = new DiskLayoutPartition
                {
                    PartitionOffsetBytes = offset,
                    StorageVolumeKey = volKey.KeyPath,
                };

                // The volume GUID associated with this volume is not in the key name —
                // the key name contains the *parent disk's* DiskId GUID. The volume's
                // own GUID would come from MountedDevices cross-reference. For now,
                // tag with a synthetic placeholder using the disk-id + offset.
                partition.VolumeGuid = $"{normalizedDiskId}#{offset:X16}";

                // Read installation/arrival timestamps for this volume registration
                var propsKey = FindSubkey(volKey, "Properties");
                if (propsKey != null)
                {
                    partition.InstalledAt = ReadPropertyAsFileTime(propsKey, DevPkeyDeviceTimestamps, PidInstalledTimestamp);
                    partition.LastArrivalAt = ReadPropertyAsFileTime(propsKey, DevPkeyDeviceTimestamps, PidArrivalTimestamp);

                    // Try to recover GPT partition type GUID. The DEVPROPKEY blob
                    // format is [4-byte type prefix][16-byte GUID]. Without skipping
                    // the 4-byte prefix the parsed GUID is shifted and never matches
                    // any well-known type (ESP, MSR, Recovery).
                    var gptTypeRaw = ReadPropertyRaw(propsKey, DevPkeyDeviceGptInfo, "000A");
                    if (gptTypeRaw != null && gptTypeRaw.Length >= 20)
                    {
                        try
                        {
                            var guidBytes = new byte[16];
                            Array.Copy(gptTypeRaw, 4, guidBytes, 0, 16);
                            var gptGuid = new Guid(guidBytes);
                            partition.GptPartitionTypeGuid = gptGuid.ToString("B");
                        }
                        catch { /* ignore */ }
                    }
                }

                // Attach to parent disk or treat as orphan
                if (diskById.TryGetValue(normalizedDiskId, out var parentDisk))
                {
                    parentDisk.Partitions.Add(partition);
                }
                else
                {
                    // Parent disk not found — orphan
                    model.OrphanPartitions.Add(partition);
                }
            }
        }

        // ════════════════════════════════════════════════════════════════════
        // Length estimation
        // ════════════════════════════════════════════════════════════════════

        /// <summary>
        /// For each partition except the last one on the disk, set EstimatedLengthBytes
        /// to (next partition's offset − this partition's offset). The last partition's
        /// length cannot be inferred from the registry alone (no disk size known).
        /// </summary>
        private static void ComputeEstimatedLengths(DiskLayoutDisk disk)
        {
            for (int i = 0; i < disk.Partitions.Count - 1; i++)
            {
                var current = disk.Partitions[i];
                var next = disk.Partitions[i + 1];
                long length = next.PartitionOffsetBytes - current.PartitionOffsetBytes;
                if (length > 0)
                {
                    current.EstimatedLengthBytes = length;
                    current.LengthIsEstimated = true;
                }
            }
            // Last partition: leave EstimatedLengthBytes null (cannot infer).
        }

        // ════════════════════════════════════════════════════════════════════
        // MountedDevices cross-reference
        // ════════════════════════════════════════════════════════════════════

        /// <summary>
        /// Walk MountedDevices and link each entry to its matching DiskLayoutPartition.
        /// MBR entries match by (disk signature, partition offset) — but since we don't
        /// know each disk's signature up front, we first deduce signatures by matching
        /// known partition offsets, then apply drive letters.
        /// </summary>
        private void ApplyMountedDevicesCrossRef(DiskLayoutModel model)
        {
            var mountedDevices = _registryInfo.GetMountedDevices();

            // Step 1: Collect MBR (signature, offset) → MountedDeviceEntry index
            // grouped by signature so we can match a signature to a disk by overlapping offsets.
            var mbrEntriesBySig = new Dictionary<uint, List<MountedDeviceEntry>>();
            var gptEntriesByGuidBytes = new List<MountedDeviceEntry>();
            var volumeGuidEntries = new List<MountedDeviceEntry>();

            foreach (var md in mountedDevices)
            {
                // Skip floppy/CD entries
                if (IsFloppyMountedEntry(md) || IsCdRomMountedEntry(md))
                    continue;

                bool isLetterOrGuid = md.MountType == "Drive Letter" || md.MountType == "Volume GUID";
                if (!isLetterOrGuid) continue;

                if (md.MountType == "Volume GUID")
                    volumeGuidEntries.Add(md);

                if (md.PartitionStyle == "MBR")
                {
                    uint? sig = TryParseMbrSignature(md.DiskSignature);
                    if (sig.HasValue)
                    {
                        if (!mbrEntriesBySig.TryGetValue(sig.Value, out var list))
                        {
                            list = new List<MountedDeviceEntry>();
                            mbrEntriesBySig[sig.Value] = list;
                        }
                        list.Add(md);
                    }
                }
                else if (md.PartitionStyle == "GPT")
                {
                    gptEntriesByGuidBytes.Add(md);
                }
            }

            // Step 2: For each MBR signature in MountedDevices, find the best
            // candidate disk in our model and assign uniquely. A disk that has
            // already received a signature is no longer a candidate.
            //
            // Scoring: a disk is a candidate for a signature when every offset
            // recorded for that signature exists in the disk's partition list.
            // We then prefer disks whose partition count exactly matches the
            // signature's offset count (tightest fit), then earlier DiskNumber.
            var disksAssignedSig = new HashSet<DiskLayoutDisk>();

            foreach (var (sig, entries) in mbrEntriesBySig)
            {
                // Distinct offsets that this signature is known to reference
                var sigOffsets = entries
                    .Select(e => ParseLbaOffset(e.PartitionOffset))
                    .Where(o => o.HasValue)
                    .Select(o => o!.Value)
                    .Distinct()
                    .ToHashSet();
                if (sigOffsets.Count == 0) continue;

                // Find candidate disks: those not yet assigned, whose partition
                // offsets are a superset of sigOffsets.
                var candidates = model.Disks
                    .Where(d => !disksAssignedSig.Contains(d))
                    .Where(d => sigOffsets.IsSubsetOf(d.Partitions.Select(p => p.PartitionOffsetBytes)))
                    .OrderBy(d => Math.Abs(d.Partitions.Count - sigOffsets.Count))
                    .ThenBy(d => d.EnumKeyPath, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                var chosen = candidates.FirstOrDefault();
                if (chosen == null) continue;

                disksAssignedSig.Add(chosen);

                // Adopt this MBR signature for the chosen disk
                chosen.MbrDiskSignature = sig;
                chosen.PartitionStyle = "MBR";

                foreach (var p in chosen.Partitions)
                {
                    p.PartitionStyle = "MBR";
                    p.MbrDiskSignature = sig;
                }

                // Assign drive letters / volume GUIDs for the matching offsets
                foreach (var md in entries)
                {
                    long? off = ParseLbaOffset(md.PartitionOffset);
                    if (!off.HasValue) continue;
                    var matchedPartition = chosen.Partitions.FirstOrDefault(p => p.PartitionOffsetBytes == off.Value);
                    if (matchedPartition == null) continue;

                    if (md.MountType == "Drive Letter")
                    {
                        matchedPartition.DriveLetter = md.MountPoint;
                        matchedPartition.RawRegistryLocations["MountedDevices (letter)"] =
                            $"MountedDevices :: {md.RegistryValueName}";
                    }
                    else if (md.MountType == "Volume GUID")
                    {
                        var guid = ExtractVolumeGuidFromMountPoint(md.MountPoint);
                        if (!string.IsNullOrEmpty(guid))
                            matchedPartition.VolumeGuid = guid;
                        matchedPartition.RawRegistryLocations["MountedDevices (volume)"] =
                            $"MountedDevices :: {md.RegistryValueName}";
                    }
                }
            }

            // Step 3: GPT MountedDevices entries don't carry a usable partition
            // offset, and our model doesn't yet know each partition's GPT GUID
            // (Phase B will add PartitionTableCache parsing for that). For now,
            // when we have unassigned GPT drive-letter entries AND disks whose
            // last partition has no drive letter, pair them in encounter order.
            // The "last partition on the disk" is the conventional location for
            // the data volume on Gen 2 GPT layouts (OS disk: WinRE, ESP, MSR, OS;
            // data disks: single large partition). This is a heuristic.
            var gptLetterEntries = gptEntriesByGuidBytes
                .Where(md => md.MountType == "Drive Letter")
                .ToList();

            if (gptLetterEntries.Count > 0)
            {
                // Candidate disks: those whose last partition (by offset) has no
                // drive letter assigned yet. Note: a disk that already received
                // an MBR signature is excluded (it's MBR, not GPT).
                var candidateDisks = model.Disks
                    .Where(d => !disksAssignedSig.Contains(d) && d.Partitions.Count > 0)
                    .Where(d => string.IsNullOrEmpty(d.Partitions[^1].DriveLetter))
                    .OrderBy(d => d.EnumKeyPath, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                int pairCount = Math.Min(gptLetterEntries.Count, candidateDisks.Count);
                for (int i = 0; i < pairCount; i++)
                {
                    var disk = candidateDisks[i];
                    var entry = gptLetterEntries[i];
                    // Pick the largest-offset partition (the "data" partition by convention)
                    var partition = disk.Partitions[^1];

                    partition.DriveLetter = entry.MountPoint;
                    partition.PartitionStyle = "GPT";
                    disk.PartitionStyle = "GPT";
                    partition.RawRegistryLocations["MountedDevices (letter, heuristic)"] =
                        $"MountedDevices :: {entry.RegistryValueName} (paired by encounter order; assigned to last partition)";

                    // Mark other partitions on this disk as GPT too
                    foreach (var p in disk.Partitions)
                    {
                        if (string.IsNullOrEmpty(p.PartitionStyle))
                            p.PartitionStyle = "GPT";
                    }
                }
                if (pairCount > 0)
                    model.Diagnostics.Add(
                        $"GPT drive letters mapped to disks by encounter order (heuristic). " +
                        $"Phase B partition-table-cache parsing required for authoritative matching.");
            }

            // Step 4: Volume-GUID-only MountedDevices entries that didn't match any
            // disk partition above become orphans (if not already present).
            foreach (var md in volumeGuidEntries)
            {
                var guid = ExtractVolumeGuidFromMountPoint(md.MountPoint);
                if (string.IsNullOrEmpty(guid)) continue;

                // Already known on a disk partition?
                bool knownOnDisk = model.Disks.SelectMany(d => d.Partitions)
                    .Any(p => string.Equals(p.VolumeGuid, guid, StringComparison.OrdinalIgnoreCase));
                if (knownOnDisk) continue;

                // Already in orphan list?
                bool inOrphans = model.OrphanPartitions
                    .Any(o => string.Equals(o.VolumeGuid, guid, StringComparison.OrdinalIgnoreCase));
                if (inOrphans) continue;

                var orphan = new DiskLayoutPartition
                {
                    VolumeGuid = guid,
                    PartitionStyle = md.PartitionStyle,
                    MbrDiskSignature = TryParseMbrSignature(md.DiskSignature),
                    Status = PartitionStatus.Stale,
                };
                orphan.RawRegistryLocations["MountedDevices"] = $"MountedDevices :: {md.RegistryValueName}";
                model.OrphanPartitions.Add(orphan);
            }
        }

        /// <summary>
        /// Remove STORAGE\Volume registrations that have no MountedDevices
        /// counterpart and look like residue from past partitioning operations.
        /// </summary>
        /// <remarks>
        /// A registration is filtered when ALL of the following hold:
        ///  - It has no DriveLetter (no MountedDevices letter assignment)
        ///  - It has no MountedDevices entry in RawRegistryLocations
        ///  - Its estimated length is small (&lt; 50 MB) — characteristic of
        ///    historical partition-table artifacts (GPT primary header at LBA 34,
        ///    1-block aligned 17 KB / 17 MB / 1 MB ghosts from past resize ops)
        ///
        /// We deliberately KEEP partitions that:
        ///  - Have a drive letter (active)
        ///  - Have a MountedDevices entry (Volume{GUID} match)
        ///  - Are &gt;= 50 MB (real ESP, MSR, Recovery, OS, data partitions)
        ///  - Are the last partition on a disk (estimated length unknown — can't
        ///    judge size, so play it safe and keep)
        ///
        /// This is conservative: false negatives (showing a stale entry) are
        /// preferable to false positives (hiding a real partition). On hives
        /// with no stale residue (e.g. EUAZRQUESTCA1, GZ-ZWESQLPWV011), the
        /// filter is effectively a no-op.
        /// </remarks>
        private void FilterStaleRegistrations(DiskLayoutModel model)
        {
            const long StaleSizeThresholdBytes = 50L * 1024 * 1024; // 50 MB
            int totalFiltered = 0;

            foreach (var disk in model.Disks)
            {
                if (disk.Partitions.Count <= 1) continue;

                var kept = new List<DiskLayoutPartition>();
                var filtered = new List<DiskLayoutPartition>();

                foreach (var p in disk.Partitions)
                {
                    bool hasDriveLetter = !string.IsNullOrEmpty(p.DriveLetter);
                    bool hasMountedDevicesEntry = p.RawRegistryLocations.Keys
                        .Any(k => k.StartsWith("MountedDevices", StringComparison.OrdinalIgnoreCase));
                    bool isLargeEnough = p.EstimatedLengthBytes.HasValue &&
                                          p.EstimatedLengthBytes.Value >= StaleSizeThresholdBytes;
                    bool isLastOnDisk = !p.EstimatedLengthBytes.HasValue;

                    if (hasDriveLetter || hasMountedDevicesEntry || isLargeEnough || isLastOnDisk)
                        kept.Add(p);
                    else
                        filtered.Add(p);
                }

                if (filtered.Count > 0)
                {
                    disk.Partitions.Clear();
                    foreach (var p in kept) disk.Partitions.Add(p);
                    totalFiltered += filtered.Count;
                    // Diagnostic uses BusType + truncated DiskId because final
                    // synthetic DiskNumber isn't assigned until step 8 (renumber).
                    var idShort = string.IsNullOrEmpty(disk.DiskId) ? "(no DiskId)"
                        : disk.DiskId.Substring(0, Math.Min(disk.DiskId.Length, 10));
                    model.Diagnostics.Add(
                        $"{disk.BusType} disk {idShort}: filtered {filtered.Count} stale STORAGE\\Volume " +
                        $"registration(s) (< 50 MB, no MountedDevices counterpart)");
                }
            }

            if (totalFiltered > 0)
            {
                model.Diagnostics.Add(
                    $"Total stale partition registrations filtered: {totalFiltered}. " +
                    "These are residue of past partitioning operations and are not currently mounted.");
            }
        }

        private static bool IsFloppyMountedEntry(MountedDeviceEntry md) =>
            string.Equals(md.DeviceService, "flpydisk", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(md.BusType, "FDC", StringComparison.OrdinalIgnoreCase) ||
            (md.EnumPath?.StartsWith("FDC\\", StringComparison.OrdinalIgnoreCase) ?? false);

        private static bool IsCdRomMountedEntry(MountedDeviceEntry md) =>
            string.Equals(md.DeviceService, "cdrom", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(md.DeviceClass, "CDROM", StringComparison.OrdinalIgnoreCase) ||
            (md.EnumPath?.IndexOf("\\CdRom", StringComparison.OrdinalIgnoreCase) ?? -1) >= 0 ||
            (md.DevicePath?.IndexOf("#CdRom", StringComparison.OrdinalIgnoreCase) ?? -1) >= 0;

        // ════════════════════════════════════════════════════════════════════
        // Role inference
        // ════════════════════════════════════════════════════════════════════

        private void InferPartitionRoles(DiskLayoutModel model)
        {
            // Read pagefile drive(s)
            var pagefileLetters = ReadPagefileLetters();

            // Read crash-dump drive
            var dumpLetter = ReadDumpLetter();

            // Read hibernation enabled flag
            // Hibernation file always lives on the boot/system drive. We tag it on C: when enabled.
            bool hibernationEnabled = ReadHibernationEnabled();

            foreach (var disk in model.Disks)
            {
                foreach (var p in disk.Partitions)
                {
                    // Drive-letter heuristic: C: is Boot+System unless BCD enricher overrides later
                    if (string.Equals(p.DriveLetter, "C:", StringComparison.OrdinalIgnoreCase))
                    {
                        p.Roles |= PartitionRoleFlags.Boot | PartitionRoleFlags.System;
                        if (string.IsNullOrEmpty(p.FilesystemType))
                        {
                            p.FilesystemType = "NTFS";
                            p.FilesystemIsInferred = true;
                        }
                        if (hibernationEnabled)
                        {
                            p.Roles |= PartitionRoleFlags.Hibernation;
                        }
                    }

                    // Pagefile role: any drive letter present in PagingFiles
                    if (!string.IsNullOrEmpty(p.DriveLetter) &&
                        pagefileLetters.Contains(p.DriveLetter, StringComparer.OrdinalIgnoreCase))
                    {
                        p.Roles |= PartitionRoleFlags.Pagefile;
                    }

                    // Crash-dump role
                    if (!string.IsNullOrEmpty(p.DriveLetter) &&
                        string.Equals(p.DriveLetter, dumpLetter, StringComparison.OrdinalIgnoreCase))
                    {
                        p.Roles |= PartitionRoleFlags.CrashDump;
                    }

                    // GPT partition type → ESP/MSR/Recovery
                    if (!string.IsNullOrEmpty(p.GptPartitionTypeGuid) &&
                        GptTypeRoles.TryGetValue(p.GptPartitionTypeGuid, out var typeRole))
                    {
                        p.Roles |= typeRole;
                        if ((typeRole & PartitionRoleFlags.ESP) != 0 && string.IsNullOrEmpty(p.FilesystemType))
                        {
                            p.FilesystemType = "FAT32";
                            p.FilesystemIsInferred = true;
                        }
                        if ((typeRole & PartitionRoleFlags.Recovery) != 0 && string.IsNullOrEmpty(p.FilesystemType))
                        {
                            p.FilesystemType = "NTFS";
                            p.FilesystemIsInferred = true;
                        }
                    }

                    // Unmounted: no drive letter and no other significant role
                    if (string.IsNullOrEmpty(p.DriveLetter) && p.Roles == PartitionRoleFlags.None)
                    {
                        p.Roles |= PartitionRoleFlags.Unmounted;
                    }
                }
            }

            // Temp-disk heuristic (Gen 1 Azure):
            // If a disk is on IDE channel 0 target 1 (key ends with "0.1.0") AND it has a
            // partition with a pagefile, mark that partition as Temp.
            foreach (var disk in model.Disks.Where(d => d.BusType == "IDE"))
            {
                if (disk.EnumKeyPath.EndsWith("0.1.0", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (var p in disk.Partitions)
                    {
                        if ((p.Roles & PartitionRoleFlags.Pagefile) != 0)
                        {
                            p.Roles |= PartitionRoleFlags.Temp;
                        }
                    }
                }
            }
        }

        private List<string> ReadPagefileLetters()
        {
            var letters = new List<string>();
            var mmKey = _parser.GetKey("ControlSet001\\Control\\Session Manager\\Memory Management");
            if (mmKey == null) return letters;

            // PagingFiles is REG_MULTI_SZ: "C:\pagefile.sys 0 0", "D:\pagefile.sys 1024 2048", ...
            var pf = mmKey.GetValue("PagingFiles");
            if (pf == null) return letters;

            string text = pf is string s
                ? s
                : pf is IEnumerable<string> arr
                    ? string.Join("\n", arr)
                    : pf.ToString() ?? "";

            foreach (var line in text.Split(new[] { '\n', '\0' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var trimmed = line.Trim();
                if (trimmed.Length >= 2 && trimmed[1] == ':')
                {
                    var letter = trimmed.Substring(0, 2).ToUpperInvariant();
                    if (!letters.Contains(letter))
                        letters.Add(letter);
                }
            }
            return letters;
        }

        private string? ReadDumpLetter()
        {
            var ccKey = _parser.GetKey("ControlSet001\\Control\\CrashControl");
            var dumpFile = ccKey?.GetValue("DumpFile") as string;
            if (string.IsNullOrEmpty(dumpFile)) return null;

            // Expand %SystemRoot% → assume C: when system path is unresolved offline
            var expanded = dumpFile.Replace("%SystemRoot%", "C:\\Windows", StringComparison.OrdinalIgnoreCase);
            if (expanded.Length >= 2 && expanded[1] == ':')
            {
                return expanded.Substring(0, 2).ToUpperInvariant();
            }
            return null;
        }

        private bool ReadHibernationEnabled()
        {
            var hibKey = _parser.GetKey("ControlSet001\\Control\\Power");
            var value = hibKey?.GetValue("HibernateEnabled");
            if (value is int i) return i != 0;
            if (value is uint u) return u != 0;
            // Also check HibernateEnabledDefault as fallback
            value = hibKey?.GetValue("HiberFileSizePercent");
            return value != null;
        }

        // ════════════════════════════════════════════════════════════════════
        // Disk-level role roll-up
        // ════════════════════════════════════════════════════════════════════

        private static void ComputeDiskRoles(DiskLayoutDisk disk)
        {
            bool hasBoot = disk.Partitions.Any(p => (p.Roles & PartitionRoleFlags.Boot) != 0);
            bool hasTemp = disk.Partitions.Any(p => (p.Roles & PartitionRoleFlags.Temp) != 0);
            bool anyMounted = disk.Partitions.Any(p => !string.IsNullOrEmpty(p.DriveLetter));
            bool anyPartitions = disk.Partitions.Count > 0;

            if (hasBoot)
                disk.Roles |= DiskRoleFlags.OsDisk;
            if (hasTemp && !hasBoot)
                disk.Roles |= DiskRoleFlags.TempDisk;
            if (!hasBoot && !hasTemp && anyMounted)
                disk.Roles |= DiskRoleFlags.DataDisk;
            if (anyPartitions && !anyMounted)
                disk.Roles |= DiskRoleFlags.Unmounted;
        }

        // ════════════════════════════════════════════════════════════════════
        // Status computation
        // ════════════════════════════════════════════════════════════════════

        private void ComputeStatuses(DiskLayoutModel model)
        {
            // A partition is Online when its STORAGE\Volume registration is present
            // (which it is, since we built it from STORAGE\Volume). Mark accordingly.
            foreach (var disk in model.Disks)
            {
                foreach (var p in disk.Partitions)
                {
                    if (p.Status == PartitionStatus.Unknown)
                        p.Status = PartitionStatus.Online;
                }
            }

            // Disk status: Online if any partition is online; Unknown otherwise.
            foreach (var disk in model.Disks)
            {
                if (disk.Partitions.Any(p => p.Status == PartitionStatus.Online))
                    disk.Status = DiskStatus.Online;
                else if (disk.Partitions.Count > 0)
                    disk.Status = DiskStatus.Offline;
                else
                    disk.Status = DiskStatus.Unknown;
            }

            // Orphans default to Stale
            foreach (var orphan in model.OrphanPartitions)
            {
                if (orphan.Status == PartitionStatus.Unknown)
                    orphan.Status = PartitionStatus.Stale;
            }
        }

        // ════════════════════════════════════════════════════════════════════
        // Numbering
        // ════════════════════════════════════════════════════════════════════

        private static void RenumberDisks(DiskLayoutModel model)
        {
            // Order: IDE first (sorted by EnumKeyPath), then SCSI, then STORAGE
            int n = 0;
            var ordered = model.Disks
                .OrderBy(d => d.BusType switch { "IDE" => 0, "SCSI" => 1, "STORAGE" => 2, _ => 3 })
                .ThenBy(d => d.EnumKeyPath, StringComparer.OrdinalIgnoreCase)
                .ToList();

            model.Disks.Clear();
            foreach (var d in ordered)
            {
                d.DiskNumber = n++;
                model.Disks.Add(d);
                foreach (var p in d.Partitions)
                    p.ParentDiskNumber = d.DiskNumber;
            }
        }

        // ════════════════════════════════════════════════════════════════════
        // Raw locations
        // ════════════════════════════════════════════════════════════════════

        private static void BuildRawLocations(DiskLayoutModel model)
        {
            foreach (var d in model.Disks)
            {
                d.RawRegistryLocations["Enum"] = d.EnumKeyPath;
                foreach (var p in d.Partitions)
                {
                    if (!string.IsNullOrEmpty(p.StorageVolumeKey))
                        p.RawRegistryLocations["STORAGE\\Volume"] = p.StorageVolumeKey;
                }
            }
            foreach (var orphan in model.OrphanPartitions)
            {
                if (!string.IsNullOrEmpty(orphan.StorageVolumeKey))
                    orphan.RawRegistryLocations["STORAGE\\Volume"] = orphan.StorageVolumeKey;
            }
        }

        // ════════════════════════════════════════════════════════════════════
        // Helpers — registry navigation
        // ════════════════════════════════════════════════════════════════════

        private static RegistryKey? FindSubkey(
            RegistryKey parent, params string[] pathParts)
        {
            var current = parent;
            foreach (var part in pathParts)
            {
                var next = current.SubKeys.FirstOrDefault(k =>
                    string.Equals(k.KeyName, part, StringComparison.OrdinalIgnoreCase));
                if (next == null) return null;
                current = next;
            }
            return current;
        }

        private static DateTime? ReadPropertyAsFileTime(
            RegistryKey propsKey, string devPkeyGuid, string propertyId)
        {
            var pkey = FindSubkey(propsKey, devPkeyGuid, propertyId);
            if (pkey == null) return null;
            var defaultVal = pkey.Values.FirstOrDefault(IsDefaultValue);
            if (defaultVal?.ValueDataRaw == null || defaultVal.ValueDataRaw.Length < 8) return null;
            try
            {
                long fileTime = BitConverter.ToInt64(defaultVal.ValueDataRaw, 0);
                if (fileTime <= 0) return null;
                return DateTime.FromFileTimeUtc(fileTime);
            }
            catch { return null; }
        }

        private static string? ReadPropertyAsString(
            RegistryKey propsKey, string devPkeyGuid, string propertyId)
        {
            var pkey = FindSubkey(propsKey, devPkeyGuid, propertyId);
            var defaultVal = pkey?.Values.FirstOrDefault(IsDefaultValue);
            if (defaultVal?.ValueDataRaw == null || defaultVal.ValueDataRaw.Length < 6) return null;

            // The data starts with a 4-byte type prefix followed by null-terminated UTF-16.
            // Skip the prefix and decode.
            try
            {
                int offset = 4;
                int count = defaultVal.ValueDataRaw.Length - offset;
                if (count <= 0) return null;
                var s = Encoding.Unicode.GetString(defaultVal.ValueDataRaw, offset, count);
                return s.TrimEnd('\0');
            }
            catch { return null; }
        }

        private static byte[]? ReadPropertyRaw(
            RegistryKey propsKey, string devPkeyGuid, string propertyId)
        {
            var pkey = FindSubkey(propsKey, devPkeyGuid, propertyId);
            var defaultVal = pkey?.Values.FirstOrDefault(IsDefaultValue);
            return defaultVal?.ValueDataRaw;
        }

        /// <summary>
        /// The vendored registry library exposes the default value with the literal
        /// name "(default)" rather than an empty string. Match either form.
        /// </summary>
        private static bool IsDefaultValue(KeyValue v) =>
            string.IsNullOrEmpty(v.ValueName) ||
            string.Equals(v.ValueName, "(default)", StringComparison.OrdinalIgnoreCase);

        // ════════════════════════════════════════════════════════════════════
        // Helpers — parsing
        // ════════════════════════════════════════════════════════════════════

        private static string NormalizeGuid(string raw)
        {
            var t = raw.Trim();
            if (!t.StartsWith("{")) t = "{" + t;
            if (!t.EndsWith("}")) t = t + "}";
            return t;
        }

        private static string TrimHardwareDescription(string hardwareIdLikeName)
        {
            // Names like "DiskVirtual_HD______________________________1.1.0___" are ugly;
            // strip underscores and excess whitespace for friendlier presentation.
            var clean = hardwareIdLikeName.Replace('_', ' ').Trim();
            // Collapse runs of whitespace
            while (clean.Contains("  "))
                clean = clean.Replace("  ", " ");
            return clean;
        }

        private static string StripInfPrefix(string raw)
        {
            // Many values are "@disk.inf,%vhd_friendlyname%;Microsoft Virtual Disk"
            // We want the part after the last ';'
            var idx = raw.LastIndexOf(';');
            if (idx >= 0 && idx < raw.Length - 1)
                return raw.Substring(idx + 1).Trim();
            return raw.Trim();
        }

        private static string? ExtractVolumeGuidFromMountPoint(string mountPoint)
        {
            // mountPoint may be like "Volume{007498b5-f45b-11e7-a944-806e6f6e6963}"
            var start = mountPoint.IndexOf('{');
            var end = mountPoint.LastIndexOf('}');
            if (start >= 0 && end > start)
                return mountPoint.Substring(start, end - start + 1);
            return null;
        }

        private static long? ParseLbaOffset(string partitionOffsetText)
        {
            // PartitionOffset from MountedDeviceEntry is formatted like
            // "1,048,576 bytes (LBA 2,048)". Parse the leading number.
            if (string.IsNullOrEmpty(partitionOffsetText)) return null;
            var firstSpace = partitionOffsetText.IndexOf(' ');
            var numeric = firstSpace > 0 ? partitionOffsetText.Substring(0, firstSpace) : partitionOffsetText;
            numeric = numeric.Replace(",", "").Trim();
            if (long.TryParse(numeric, NumberStyles.Integer, CultureInfo.InvariantCulture, out long bytes))
                return bytes;
            return null;
        }

        private static uint? TryParseMbrSignature(string sigText)
        {
            // sigText might be "0x50E1C300"
            if (string.IsNullOrEmpty(sigText)) return null;
            var t = sigText.Trim();
            if (t.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
                t = t.Substring(2);
            if (uint.TryParse(t, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out uint val))
                return val;
            return null;
        }
    }
}

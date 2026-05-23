using System;
using System.Collections.Generic;

namespace RegistryExpert.Core.Models
{
    /// <summary>
    /// One partition (volume) belonging to a physical disk, or an orphan volume
    /// whose parent disk is no longer present. Models all data recoverable
    /// from the offline registry plus optional external enrichments (BCD,
    /// diskinfo.txt).
    /// </summary>
    public class DiskLayoutPartition
    {
        /// <summary>Persistent Volume GUID, e.g. "{091c7ef3-7ef3-7ef3-95b3-4de42cf2df86}".</summary>
        public string VolumeGuid { get; set; } = "";

        /// <summary>Drive letter with colon (e.g. "C:") or null when no letter is assigned.</summary>
        public string? DriveLetter { get; set; }

        /// <summary>Human-readable volume label (e.g. "Windows", "Temporary Storage").
        /// Null when label could not be determined.</summary>
        public string? VolumeLabel { get; set; }

        /// <summary>True when VolumeLabel was inferred (e.g. boot volume → "Windows")
        /// rather than read from a definitive source.</summary>
        public bool VolumeLabelIsInferred { get; set; }

        /// <summary>"MBR", "GPT", or "" if not determined.</summary>
        public string PartitionStyle { get; set; } = "";

        /// <summary>Byte offset of partition from start of disk (parsed from
        /// MountedDevices payload bytes 4-11 for MBR, or from the
        /// STORAGE\Volume key name's hex suffix after #).</summary>
        public long PartitionOffsetBytes { get; set; }

        /// <summary>Estimated partition length in bytes — computed by subtracting
        /// consecutive partition offsets on the same disk. Null when length
        /// cannot be inferred (last partition on disk, no enrichment source).
        /// Replaced by authoritative value when diskinfo.txt enricher applies.</summary>
        public long? EstimatedLengthBytes { get; set; }

        /// <summary>True when EstimatedLengthBytes was inferred via offset subtraction.
        /// False when set authoritatively from an enricher (e.g. diskinfo.txt).</summary>
        public bool LengthIsEstimated { get; set; }

        /// <summary>MBR disk signature (for MBR-style partitions only). Null for GPT.</summary>
        public uint? MbrDiskSignature { get; set; }

        /// <summary>GPT partition type GUID when known (e.g. "{C12A7328-F81F-11D2-BA4B-00A0C93EC93B}"
        /// for ESP). Null for MBR or when not recoverable.</summary>
        public string? GptPartitionTypeGuid { get; set; }

        /// <summary>Filesystem type: "NTFS", "FAT32", "ReFS", etc. Null if unknown.</summary>
        /// <remarks>
        /// Populated by <c>DiskLayoutExtractor</c>, <c>DiskInfoTxtEnricher</c>, and
        /// <c>DiskLayoutBcdEnricher</c>, but NOT currently displayed in any UI
        /// surface — the File System column was removed in Phase A.6 after the
        /// explorer investigation determined that offline registry data cannot
        /// reliably identify filesystem type (Services\Ntfs\Instances is empty
        /// offline; no per-volume DEVPROPKEY contains FS). Retained on the model
        /// for potential future re-add (e.g. when BitLocker detection or NTUSER
        /// MountPoints2 correlation lands).
        /// </remarks>
        public string? FilesystemType { get; set; }

        /// <summary>True when FilesystemType was inferred from role rather than read from
        /// an authoritative source. Always honor this flag when displaying.</summary>
        /// <remarks>
        /// See <see cref="FilesystemType"/> remarks — currently not displayed.
        /// </remarks>
        public bool FilesystemIsInferred { get; set; }

        /// <summary>Bitwise flags of partition roles.</summary>
        public PartitionRoleFlags Roles { get; set; } = PartitionRoleFlags.None;

        /// <summary>When this volume was first registered under STORAGE\Volume.
        /// From DEVPKEY {83da6326}\0064. Null if unavailable.</summary>
        public DateTime? InstalledAt { get; set; }

        /// <summary>When this volume was last seen online (most recent arrival).
        /// From DEVPKEY {83da6326}\0066. Null if unavailable.</summary>
        public DateTime? LastArrivalAt { get; set; }

        /// <summary>Inferred status from cross-referencing MountedDevices vs STORAGE\Volume.</summary>
        public PartitionStatus Status { get; set; } = PartitionStatus.Unknown;

        /// <summary>Synthetic disk number of the owning disk (null when orphan).</summary>
        public int? ParentDiskNumber { get; set; }

        /// <summary>Full registry path of the matching STORAGE\Volume key (when present).</summary>
        public string StorageVolumeKey { get; set; } = "";

        // ── Enrichments from diskinfo.txt (InspectIaaSDisk bundles) ────────────

        /// <summary>Authoritative free bytes from diskinfo.txt. Null when no enrichment available.</summary>
        public long? FreeSpaceBytes { get; set; }

        /// <summary>Authoritative used bytes from diskinfo.txt.</summary>
        public long? UsedSpaceBytes { get; set; }

        /// <summary>NTFS cluster / filesystem block size from diskinfo.txt.</summary>
        public int? ClusterSizeBytes { get; set; }

        /// <summary>Total file/inode capacity (statvfs files).</summary>
        public long? TotalInodes { get; set; }

        /// <summary>Free inodes available.</summary>
        public long? FreeInodes { get; set; }

        /// <summary>True when capacity values came from an authoritative external source
        /// (e.g. diskinfo.txt) rather than being estimated from the registry.</summary>
        /// <remarks>
        /// Currently only used by the Volume Details Expander for an annotation;
        /// no longer drives any column-level rendering since the volume table
        /// shows capacity uniformly. Retained for potential UI re-use.
        /// </remarks>
        public bool CapacityFromExternalSource { get; set; }

        /// <summary>Map of registry/file locations that contributed to this partition's
        /// data. Used for the details pane "Raw registry locations" row.
        /// Key = source label (e.g. "MountedDevices", "STORAGE\\Volume", "BCD osdevice",
        /// "diskinfo.txt"). Value = the actual path/key.</summary>
        public Dictionary<string, string> RawRegistryLocations { get; set; } = new();
    }
}

using System;
using System.Collections.Generic;

namespace RegistryExpert.Core.Models
{
    /// <summary>
    /// One physical (or virtual) disk enumerated from the offline registry,
    /// along with its partitions and inferred role.
    /// </summary>
    public class DiskLayoutDisk
    {
        /// <summary>Synthetic 0-based disk number, ordered by bus then index
        /// (IDE first by channel/target, then SCSI by LUN).</summary>
        public int DiskNumber { get; set; }

        /// <summary>PnP Disk GUID from ControlSet001\Enum\…\Device Parameters\Partmgr\DiskId.
        /// Empty string if not available (e.g. hives missing the Partmgr cache).</summary>
        public string DiskId { get; set; } = "";

        /// <summary>FriendlyName from the disk's Properties dictionary
        /// (DEVPKEY {a8b865dd}\0002). e.g. "Virtual HD ATA Device", "Microsoft Virtual Disk".</summary>
        public string FriendlyName { get; set; } = "";

        /// <summary>Manufacturer string (DEVPKEY {a8b865dd}\0003) e.g. "Microsoft".</summary>
        public string Manufacturer { get; set; } = "";

        /// <summary>"IDE", "SCSI", "STORAGE", "NVMe", etc. — derived from the
        /// Enum sub-bus and CompatibleIDs.</summary>
        public string BusType { get; set; } = "";

        /// <summary>Human-friendly bus position from Enum LocationInformation
        /// (e.g. "Channel 0, Target 0, LUN 0" or "Bus Number 0, Target Id 0, LUN 1").</summary>
        public string BusLocation { get; set; } = "";

        /// <summary>"MBR", "GPT", or "Unknown" if no partitions or undetermined.</summary>
        public string PartitionStyle { get; set; } = "Unknown";

        /// <summary>MBR disk signature (when partition style is MBR).</summary>
        public uint? MbrDiskSignature { get; set; }

        /// <summary>Partitions belonging to this disk, sorted by PartitionOffsetBytes ascending.</summary>
        public List<DiskLayoutPartition> Partitions { get; set; } = new();

        /// <summary>Inferred online/offline/stale status.</summary>
        public DiskStatus Status { get; set; } = DiskStatus.Unknown;

        /// <summary>When this disk was first registered (DEVPKEY {83da6326}\0064).</summary>
        public DateTime? InstalledAt { get; set; }

        /// <summary>When this disk was last seen online (DEVPKEY {83da6326}\0066).</summary>
        public DateTime? LastArrivalAt { get; set; }

        /// <summary>ACPI / hardware path string when recoverable (DEVPKEY {a45c254e}\0025
        /// after parsing the 4-byte header prefix). e.g. "\_SB.PCI0.IDE0.CHN0.DRV0".</summary>
        public string? AcpiPath { get; set; }

        /// <summary>Full registry path of the disk's Enum key (for "Raw locations" details).</summary>
        public string EnumKeyPath { get; set; } = "";

        /// <summary>Bitwise role flags.</summary>
        public DiskRoleFlags Roles { get; set; } = DiskRoleFlags.None;

        /// <summary>Map of registry locations that contributed to this disk's data.</summary>
        public Dictionary<string, string> RawRegistryLocations { get; set; } = new();
    }
}

using System;

namespace RegistryExpert.Core.Models
{
    /// <summary>
    /// Parsed contents of an Azure InspectIaaSDisk bundle's diskinfo.txt file.
    /// Microsoft's InspectIaaSDisk tool mounts a customer's Windows OS VHD on a
    /// Linux helper VM and dumps `df` / `statvfs` output here. This gives us
    /// authoritative Capacity / Free / Used for the C: volume — data otherwise
    /// not present in the offline registry.
    /// </summary>
    /// <remarks>
    /// One instance represents one filesystem entry (typically one Windows volume,
    /// usually C:). A bundle may produce multiple instances if InspectIaaSDisk
    /// ever ships data for multiple partitions (the parser is flexible).
    /// </remarks>
    public class DiskInfoTxtData
    {
        /// <summary>Windows drive letter as reported (e.g. "C:").</summary>
        public string WindowsDriveLetter { get; set; } = "";

        /// <summary>Linux device path as reported (e.g. "/dev/sda4").</summary>
        public string LinuxDevice { get; set; } = "";

        /// <summary>Disk number derived from the Linux device name
        /// ("sda" → 0, "sdb" → 1). Null if unparseable.</summary>
        public int? DiskNumber { get; set; }

        /// <summary>Partition number derived from trailing digits
        /// ("sda4" → 4). Null if unparseable.</summary>
        public int? PartitionNumber { get; set; }

        /// <summary>Authoritative total partition size in bytes (blocks × bsize from statvfs).</summary>
        public long? TotalBytes { get; set; }

        /// <summary>Authoritative free bytes (bfree × bsize from statvfs).</summary>
        public long? FreeBytes { get; set; }

        /// <summary>Used bytes (TotalBytes − FreeBytes).</summary>
        public long? UsedBytes { get; set; }

        /// <summary>Filesystem block size (statvfs bsize). For NTFS this is the cluster size.</summary>
        public int? ClusterSizeBytes { get; set; }

        /// <summary>Filesystem hint inferred from bsize ("NTFS (likely)" when bsize=4096).
        /// Always treat as a hint, never authoritative.</summary>
        public string? FilesystemHint { get; set; }

        /// <summary>Total inodes (NTFS: MFT entries cap).</summary>
        public long? TotalInodes { get; set; }

        /// <summary>Free inodes available for new files.</summary>
        public long? FreeInodes { get; set; }

        /// <summary>Timestamp of the diskinfo.txt file itself (proxy for when InspectIaaSDisk ran).</summary>
        public DateTime ExtractedAt { get; set; }

        /// <summary>Full path of the source diskinfo.txt file (for "Raw locations" display).</summary>
        public string SourcePath { get; set; } = "";
    }
}

using System;
using System.Collections.Generic;

namespace RegistryExpert.Core.Models
{
    /// <summary>
    /// Top-level container produced by <c>DiskLayoutExtractor</c>. Represents a
    /// diskmgmt.msc-style view of all storage data recoverable from the loaded
    /// hives, including disks, partitions, and orphan volume registrations.
    /// </summary>
    public class DiskLayoutModel
    {
        /// <summary>All disks enumerated from the registry, ordered by DiskNumber.</summary>
        public List<DiskLayoutDisk> Disks { get; set; } = new();

        /// <summary>Volume registrations whose parent disk could not be located.
        /// Typically residue of disks that were once attached but have since been
        /// removed (e.g. detached Azure data disks, replaced VMware disks).</summary>
        public List<DiskLayoutPartition> OrphanPartitions { get; set; } = new();

        /// <summary>When this model was extracted (for UI timestamp display).</summary>
        public DateTime ExtractedAt { get; set; } = DateTime.UtcNow;

        /// <summary>Bitwise flags recording which data sources contributed to this model.</summary>
        public DiskLayoutSourceFlags Sources { get; set; } = DiskLayoutSourceFlags.None;

        /// <summary>Diagnostic messages produced during extraction (e.g.
        /// "BCD references partition not found in model", "Partmgr DiskId cache empty").
        /// Empty list when extraction completed cleanly.</summary>
        public List<string> Diagnostics { get; set; } = new();
    }
}

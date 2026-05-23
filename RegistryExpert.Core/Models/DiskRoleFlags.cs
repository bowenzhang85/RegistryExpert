using System;

namespace RegistryExpert.Core.Models
{
    /// <summary>
    /// Roles that a physical disk may play in the system. Computed from the
    /// roles assigned to the disk's partitions plus disk-level signals.
    /// </summary>
    [Flags]
    public enum DiskRoleFlags
    {
        None = 0,

        /// <summary>Hosts the boot/OS partition (Windows is installed here).</summary>
        OsDisk = 1 << 0,

        /// <summary>Azure resource disk / Hyper-V temp disk (IDE 0:1 on Gen 1; heuristic on Gen 2).</summary>
        TempDisk = 1 << 1,

        /// <summary>General data disk (mounted, has drive letter or folder mount, not OS or Temp).</summary>
        DataDisk = 1 << 2,

        /// <summary>Member of a Storage Spaces pool (deferred handling, flag only).</summary>
        PoolMember = 1 << 3,

        /// <summary>Disk is enumerated but has no mounted partition.</summary>
        Unmounted = 1 << 4,
    }
}

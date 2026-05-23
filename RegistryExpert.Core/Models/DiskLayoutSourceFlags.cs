using System;

namespace RegistryExpert.Core.Models
{
    /// <summary>
    /// Tracks which data sources contributed to a DiskLayoutModel so the UI can
    /// honestly report (and the user can verify) provenance of inferred values.
    /// </summary>
    [Flags]
    public enum DiskLayoutSourceFlags
    {
        None = 0,

        /// <summary>SYSTEM hive — required baseline source.</summary>
        System = 1 << 0,

        /// <summary>SOFTWARE hive — adds MountPoints2 / VolumeCaches.</summary>
        Software = 1 << 1,

        /// <summary>NTUSER.DAT — adds per-user Explorer MountPoints2 labels.</summary>
        NtUser = 1 << 2,

        /// <summary>BCD hive — adds Boot/System role identification.</summary>
        Bcd = 1 << 3,

        /// <summary>COMPONENTS hive — reserved for future Component-based servicing lookups.</summary>
        Components = 1 << 4,

        /// <summary>InspectIaaSDisk bundle diskinfo.txt — adds exact Capacity/Free/Used for the OS volume.</summary>
        DiskInfoTxt = 1 << 5,
    }
}

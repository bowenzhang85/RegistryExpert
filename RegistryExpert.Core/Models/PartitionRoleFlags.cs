using System;

namespace RegistryExpert.Core.Models
{
    /// <summary>
    /// Roles that a partition may play in the system. A single partition may
    /// have multiple roles (e.g. Boot + System + CrashDump on a typical C:).
    /// </summary>
    [Flags]
    public enum PartitionRoleFlags
    {
        None = 0,

        /// <summary>Hosts the Windows installation (BCD osdevice target).</summary>
        Boot = 1 << 0,

        /// <summary>Holds the boot loader (BCD device target). On UEFI systems
        /// this is the ESP; on BIOS systems often the System Reserved partition.</summary>
        System = 1 << 1,

        /// <summary>Contains a page file (per Session Manager\Memory Management\PagingFiles).</summary>
        Pagefile = 1 << 2,

        /// <summary>Configured crash dump destination (per CrashControl\DumpFile).</summary>
        CrashDump = 1 << 3,

        /// <summary>EFI System Partition (FAT32, GPT type guid C12A7328-...).</summary>
        ESP = 1 << 4,

        /// <summary>Microsoft Reserved Partition (GPT, no file system).</summary>
        MSR = 1 << 5,

        /// <summary>Windows Recovery Environment partition.</summary>
        Recovery = 1 << 6,

        /// <summary>MBR active/boot flag set in the partition table.</summary>
        Active = 1 << 7,

        /// <summary>Azure resource / Hyper-V temp partition (typically D: on Gen 1).</summary>
        Temp = 1 << 8,

        /// <summary>Partition is registered but has no drive letter or folder mount.</summary>
        Unmounted = 1 << 9,

        /// <summary>Contains the hibernation file (hiberfil.sys).</summary>
        Hibernation = 1 << 10,
    }
}

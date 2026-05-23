namespace RegistryExpert.Core.Models
{
    /// <summary>
    /// Status of a partition / volume as inferred from offline registry data.
    /// </summary>
    public enum PartitionStatus
    {
        /// <summary>Status could not be determined from the registry.</summary>
        Unknown = 0,

        /// <summary>Volume registration exists under STORAGE\Volume and parent disk is online.</summary>
        Online = 1,

        /// <summary>Parent disk is offline.</summary>
        Offline = 2,

        /// <summary>MountedDevices entry exists but no matching STORAGE\Volume registration
        /// — typically a residue of a removed disk.</summary>
        Stale = 3,
    }
}

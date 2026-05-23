namespace RegistryExpert.Core.Models
{
    /// <summary>
    /// Status of a physical disk as inferred from offline registry data.
    /// </summary>
    public enum DiskStatus
    {
        /// <summary>Status could not be determined from the registry.</summary>
        Unknown = 0,

        /// <summary>Disk is currently registered and reachable.</summary>
        Online = 1,

        /// <summary>Disk is registered but marked offline (e.g. via partmgr Attributes bit).</summary>
        Offline = 2,

        /// <summary>Disk was previously registered but is no longer present
        /// (e.g. detached Azure data disk that left registrations behind).</summary>
        Stale = 3,
    }
}

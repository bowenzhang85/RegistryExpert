using RegistryExpert.Core;

namespace RegistryExpert.Core.Models
{
    /// <summary>
    /// Represents a registry hive file discovered during folder scanning.
    /// </summary>
    public class DiscoveredHive
    {
        /// <summary>Full path to the hive file.</summary>
        public required string FilePath { get; init; }

        /// <summary>Predicted hive type based on filename pattern.</summary>
        public required OfflineRegistryParser.HiveType DetectedType { get; init; }

        /// <summary>Path relative to the scanned root folder (for display).</summary>
        public required string RelativePath { get; init; }
    }
}

namespace RegistryExpert.Core.Models
{
    public class PhysicalDiskEntry
    {
        // Identification
        public string DeviceId { get; set; } = "";        // e.g. "SCSI\Disk&Ven_Msft&Prod_Virtual_Disk\000000"
        public string FriendlyName { get; set; } = "";     // e.g. "Microsoft Virtual Disk"
        public string DeviceDesc { get; set; } = "";       // e.g. "Disk drive"

        // Bus/Location
        public string BusType { get; set; } = "";          // IDE, SCSI, USBSTOR, NVME
        public string LocationInfo { get; set; } = "";     // e.g. "Bus 0, Target 0, LUN 0"

        // Hardware
        public string HardwareId { get; set; } = "";       // First hardware ID string
        public string CompatibleIds { get; set; } = "";    // Compatible IDs
        public string Manufacturer { get; set; } = "";     // Mfg value (cleaned)
        public string Service { get; set; } = "";          // Driver service name
        public string DeviceClass { get; set; } = "";      // Class GUID description

        // Partmgr
        public string DiskId { get; set; } = "";           // Partmgr DiskId GUID
        public int VolumeCount { get; set; }               // Number of STORAGE\Volume entries for this DiskId
        public string PoolStatus { get; set; } = "";       // "Probable Pool Member", "Normal", or ""

        // Drive letter mapping
        public string DriveLetters { get; set; } = "";     // Mapped drive letters (e.g. "C:, D:")
        public string PartitionStyle { get; set; } = "";   // MBR, GPT, or Unknown

        // Registry
        public string EnumPath { get; set; } = "";         // Full enum path
        public string RegistryPath { get; set; } = "";     // Display path

        // Status
        public string DeviceStatus { get; set; } = "";     // Enabled/Disabled from ConfigFlags
    }
}

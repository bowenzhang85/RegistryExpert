namespace RegistryExpert.Core.Models
{
    public class MountedDeviceEntry
    {
        public string MountPoint { get; set; } = "";
        public string MountType { get; set; } = "";
        public string PartitionStyle { get; set; } = "";
        public string Identifier { get; set; } = "";
        public string DiskSignature { get; set; } = "";
        public string PartitionOffset { get; set; } = "";
        public string PartitionGuid { get; set; } = "";
        public string DevicePath { get; set; } = "";
        public string BusType { get; set; } = "";
        public string Vendor { get; set; } = "";
        public string Product { get; set; } = "";
        public string Serial { get; set; } = "";
        public string RegistryPath { get; set; } = "MountedDevices";
        public string RegistryValueName { get; set; } = "";
        public int DataLength { get; set; }

        // Enum-enriched properties (from ControlSet001\Enum\{Bus}\{Device}\{Instance})
        public string FriendlyName { get; set; } = "";
        public string DeviceClass { get; set; } = "";
        public string DeviceService { get; set; } = "";
        public string Manufacturer { get; set; } = "";
        public string LocationInfo { get; set; } = "";
        public string DeviceStatus { get; set; } = "";
        public string EnumPath { get; set; } = "";
        public string DiskId { get; set; } = "";
        public string StaleStatus { get; set; } = "";
    }
}

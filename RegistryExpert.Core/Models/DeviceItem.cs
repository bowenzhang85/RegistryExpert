namespace RegistryExpert.Core.Models
{
    public class DeviceItem
    {
        public string DisplayName { get; set; } = "";
        public string RegistryPath { get; set; } = "";
        public List<DevicePropertyItem> Properties { get; set; } = new();
        public List<DevicePropertyItem> DriverProperties { get; set; } = new();
        public string DriverRegistryPath { get; set; } = "";

        public override string ToString() => DisplayName;
    }
}

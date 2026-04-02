namespace RegistryExpert.Core.Models
{
    public class DeviceClassItem
    {
        public string ClassName { get; set; } = "";
        public string ClassGuid { get; set; } = "";
        public string RegistryPath { get; set; } = "";
        public List<DeviceItem> Devices { get; set; } = new();

        public override string ToString() => ClassName;
    }
}

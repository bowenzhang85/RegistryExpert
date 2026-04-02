namespace RegistryExpert.Core.Models
{
    public class NetworkAdapterItem
    {
        public string DisplayName { get; set; } = "";
        public string RegistryPath { get; set; } = "";
        public string FullGuid { get; set; } = "";
        public List<NetworkPropertyItem> Properties { get; set; } = new();

        public override string ToString() => DisplayName;
    }
}

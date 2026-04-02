namespace RegistryExpert.Core.Models
{
    public class RoleFeatureItem
    {
        public string KeyName { get; set; } = "";
        public string DisplayName { get; set; } = "";
        public string Description { get; set; } = "";
        public int ServerComponentType { get; set; }   // 0=Role, 1=RoleService, 2=Feature
        public int InstallState { get; set; }           // 0=Not Installed, 1=Installed
        public string ParentName { get; set; } = "";
        public int NumericId { get; set; }
        public string RegistryPath { get; set; } = "";
        public string SystemServices { get; set; } = "";
        public string Dependencies { get; set; } = "";
        public int MajorVersion { get; set; }
        public int MinorVersion { get; set; }

        public string ComponentTypeName => ServerComponentType switch
        {
            0 => "Role",
            1 => "Role Service",
            2 => "Feature",
            _ => $"Unknown ({ServerComponentType})"
        };

        public string InstallStateName => InstallState == 1 ? "Installed" : "Not Installed";

        public override string ToString() => DisplayName;
    }
}

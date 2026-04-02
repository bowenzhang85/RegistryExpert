namespace RegistryExpert.Core.Models
{
    public class AnalysisItem
    {
        public string Name { get; set; } = "";
        public string Value { get; set; } = "";
        public string RegistryPath { get; set; } = "";
        public string RegistryValue { get; set; } = "";
        public bool IsSubSection { get; set; } = false;
        public bool IsWarning { get; set; } = false;
        public List<AnalysisItem>? SubItems { get; set; }
    }
}

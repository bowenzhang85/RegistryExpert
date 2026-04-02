namespace RegistryExpert.Core.Models
{
    public class AnalysisSection
    {
        public string Title { get; set; } = "";
        public List<AnalysisItem> Items { get; set; } = new();
        public object? Tag { get; set; }
    }
}

namespace RegistryExpert.Core.Models
{
    /// <summary>
    /// Represents a detected IID or TSS log bundle folder.
    /// </summary>
    public class BundleInfo
    {
        public required string FolderPath { get; init; }
        public required string Name { get; init; }
        public required string BundleType { get; init; }
        public required DateTime ModifiedDate { get; init; }

        /// <summary>
        /// Human-friendly date string (e.g. "Today 14:30", "Yesterday", "Apr 7").
        /// </summary>
        public string DisplayDate
        {
            get
            {
                var now = DateTime.Now;
                if (ModifiedDate.Date == now.Date)
                    return $"Today {ModifiedDate:HH:mm}";
                if (ModifiedDate.Date == now.Date.AddDays(-1))
                    return $"Yesterday {ModifiedDate:HH:mm}";
                return ModifiedDate.ToString("MMM d, HH:mm");
            }
        }
    }
}

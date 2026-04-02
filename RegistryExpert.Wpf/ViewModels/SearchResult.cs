namespace RegistryExpert.Wpf.ViewModels
{
    /// <summary>
    /// One row in the search results DataGrid.
    /// </summary>
    public class SearchResult
    {
        /// <summary>Display path: hive-prefixed (e.g. "SYSTEM\ControlSet001\Services").</summary>
        public string KeyPath { get; init; } = "";

        /// <summary>Raw key path from parser (for navigation).</summary>
        public string RawKeyPath { get; init; } = "";

        /// <summary>"Key" for key matches, or the value type string (e.g. "RegSz", "RegDword").</summary>
        public string MatchType { get; init; } = "";

        /// <summary>Key name (key match) or value name (value match). "(Default)" for empty names.</summary>
        public string Details { get; init; } = "";

        /// <summary>Raw value name for navigation (empty string for default value).</summary>
        public string ValueName { get; init; } = "";

        /// <summary>Cleaned value data, truncated for grid display (max 200 chars).</summary>
        public string Data { get; init; } = "";

        /// <summary>Full multi-line preview text (Name + Type + Data, or "Key: path").</summary>
        public string FullValue { get; init; } = "";

        /// <summary>Image key for the row icon: "folder" for keys, "reg_str"/"reg_num"/"reg_bin" for values.</summary>
        public string ImageKey { get; init; } = "";

        /// <summary>The hive type name this result belongs to (e.g. "SYSTEM").</summary>
        public string HiveTypeName { get; init; } = "";
    }
}

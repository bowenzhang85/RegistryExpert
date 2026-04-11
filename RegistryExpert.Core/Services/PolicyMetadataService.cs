namespace RegistryExpert.Core.Services;

/// <summary>
/// Provides ADMX policy metadata (display names, descriptions, category paths)
/// by looking up registry paths against an embedded policy definitions database
/// derived from the Windows 11 25H2 Policy Settings spreadsheet.
/// </summary>
public class PolicyMetadataService
{
    private static PolicyMetadataService? _instance;
    private static readonly object _lock = new();

    private readonly Dictionary<string, PolicyInfo> _policies;

    public static PolicyMetadataService Instance
    {
        get
        {
            if (_instance == null)
            {
                lock (_lock)
                {
                    _instance ??= new PolicyMetadataService();
                }
            }
            return _instance;
        }
    }

    private PolicyMetadataService()
    {
        _policies = new Dictionary<string, PolicyInfo>(StringComparer.OrdinalIgnoreCase);
        LoadEmbeddedMetadata();
    }

    /// <summary>
    /// Try to find ADMX metadata for a registry path + value name.
    /// Handles two key formats from the spreadsheet:
    ///   1. "path!valueName" (Administrative Templates format)
    ///   2. "path\valueName" (Security Options format — value as last path segment)
    /// </summary>
    /// <param name="registryPath">Path without hive prefix, e.g. "Policies\Microsoft\Windows\BITS"</param>
    /// <param name="valueName">Value name, e.g. "EnableBITSMaxBandwidth"</param>
    /// <param name="info">The policy metadata if found</param>
    /// <param name="isKeyLevelMatch">True if the match came from a key-level lookup (enabledList target key)</param>
    public bool TryGetPolicyInfo(string registryPath, string valueName, out PolicyInfo info, out bool isKeyLevelMatch)
    {
        var normalizedPath = NormalizePath(registryPath);
        isKeyLevelMatch = false;

        // Try format 1: path!valueName (most common — Administrative Templates)
        if (!string.IsNullOrEmpty(valueName))
        {
            var key1 = $"{normalizedPath}!{valueName}";
            if (_policies.TryGetValue(key1, out info!))
                return true;

            // Try format 2: path\valueName (Security Options — value embedded in path)
            var key2 = $@"{normalizedPath}\{valueName}";
            if (_policies.TryGetValue(key2, out info!))
                return true;
        }

        // Try path-only lookup (enabledList target keys or key-level policies)
        if (_policies.TryGetValue(normalizedPath, out info!))
        {
            isKeyLevelMatch = true;
            return true;
        }

        return false;
    }

    /// <summary>
    /// Try to find the ADMX category path for a registry path by checking
    /// if any policy under this path has a known category.
    /// </summary>
    public bool TryGetCategoryPath(string registryPath, out string categoryPath)
    {
        var normalized = NormalizePath(registryPath);
        foreach (var kvp in _policies)
        {
            var policyPath = kvp.Key.Contains('!')
                ? kvp.Key[..kvp.Key.IndexOf('!')]
                : kvp.Key;
            if (policyPath.Equals(normalized, StringComparison.OrdinalIgnoreCase) ||
                policyPath.StartsWith(normalized + @"\", StringComparison.OrdinalIgnoreCase))
            {
                categoryPath = kvp.Value.CategoryPath;
                return true;
            }
        }
        categoryPath = "";
        return false;
    }

    public int PolicyCount => _policies.Count;

    private static string NormalizePath(string path)
    {
        var p = path.Replace("/", @"\");
        // Remove hive prefixes
        if (p.StartsWith(@"HKLM\", StringComparison.OrdinalIgnoreCase))
            p = p[5..];
        if (p.StartsWith(@"HKEY_LOCAL_MACHINE\", StringComparison.OrdinalIgnoreCase))
            p = p[19..];
        if (p.StartsWith(@"HKCU\", StringComparison.OrdinalIgnoreCase))
            p = p[5..];
        if (p.StartsWith(@"HKEY_CURRENT_USER\", StringComparison.OrdinalIgnoreCase))
            p = p[18..];
        // Remove SOFTWARE prefix (our hive paths don't include it)
        if (p.StartsWith(@"SOFTWARE\", StringComparison.OrdinalIgnoreCase))
            p = p[9..];
        return p.TrimStart('\\');
    }

    private void LoadEmbeddedMetadata()
    {
        try
        {
            var assembly = typeof(PolicyMetadataService).Assembly;
            var resourceName = assembly.GetManifestResourceNames()
                .FirstOrDefault(n => n.EndsWith("PolicyMetadata.json", StringComparison.OrdinalIgnoreCase));
            if (resourceName == null) return;

            using var stream = assembly.GetManifestResourceStream(resourceName);
            if (stream == null) return;

            using var reader = new System.IO.StreamReader(stream);
            var json = reader.ReadToEnd();

            var doc = System.Text.Json.JsonDocument.Parse(json);
            if (doc.RootElement.TryGetProperty("policies", out var policies))
            {
                foreach (var prop in policies.EnumerateObject())
                {
                    var policyInfo = new PolicyInfo
                    {
                        Name = prop.Value.GetProperty("n").GetString() ?? "",
                        CategoryPath = prop.Value.GetProperty("p").GetString() ?? "",
                        Description = prop.Value.TryGetProperty("d", out var d) ? d.GetString() ?? "" : "",
                        SupportedOn = prop.Value.TryGetProperty("s", out var s) ? s.GetString() ?? "" : "",
                        Scope = prop.Value.TryGetProperty("sc", out var sc) ? sc.GetString() ?? "" : ""
                    };

                    // Load ADMX value definitions
                    if (prop.Value.TryGetProperty("ev", out var ev) && ev.ValueKind == System.Text.Json.JsonValueKind.Number)
                        policyInfo.EnabledValue = ev.GetInt32();
                    if (prop.Value.TryGetProperty("dv", out var dv) && dv.ValueKind == System.Text.Json.JsonValueKind.Number)
                        policyInfo.DisabledValue = dv.GetInt32();
                    if (prop.Value.TryGetProperty("b", out var b) && b.ValueKind == System.Text.Json.JsonValueKind.True)
                        policyInfo.IsBare = true;
                    if (prop.Value.TryGetProperty("e", out var e) && e.ValueKind == System.Text.Json.JsonValueKind.Object)
                    {
                        policyInfo.EnumValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        foreach (var enumProp in e.EnumerateObject())
                        {
                            var enumVal = enumProp.Value.GetString();
                            if (!string.IsNullOrEmpty(enumVal))
                                policyInfo.EnumValues[enumProp.Name] = enumVal;
                        }
                    }

                    // Normalize the JSON key to match hive-relative paths
                    var key = NormalizeJsonKey(prop.Name);
                    _policies.TryAdd(key, policyInfo);
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Failed to load policy metadata: {ex.Message}");
        }
    }

    /// <summary>
    /// Normalizes a JSON key from the metadata file to match the format used by TryGetPolicyInfo lookups.
    /// Handles three key formats from the spreadsheet:
    ///   1. "policies\microsoft\windows\bits!valueName" — Administrative Templates (already normalized)
    ///   2. "machine\software\microsoft\...\key\valueName" — Security Options (strip machine\software\ prefix)
    ///   3. "machine\software\...\key, value=valueName" — Security Options variant (convert to key\valueName)
    /// </summary>
    private static string NormalizeJsonKey(string key)
    {
        var k = key;

        // Handle ", value=" format → convert to backslash separator
        // e.g. "machine\software\...\system, value=foo" → "machine\software\...\system\foo"
        var commaIdx = k.IndexOf(", value=", StringComparison.OrdinalIgnoreCase);
        if (commaIdx >= 0)
        {
            var path = k[..commaIdx];
            var valName = k[(commaIdx + 8)..]; // 8 = ", value=".Length
            k = $@"{path}\{valName}";
        }

        // Strip "machine\software\" prefix (Security Options keys)
        if (k.StartsWith(@"machine\software\", StringComparison.OrdinalIgnoreCase))
            k = k[17..];
        // Strip "machine\" prefix for any remaining edge cases
        else if (k.StartsWith(@"machine\", StringComparison.OrdinalIgnoreCase))
            k = k[8..];

        return k;
    }
}

/// <summary>
/// ADMX policy metadata for a single policy setting.
/// </summary>
public class PolicyInfo
{
    /// <summary>Display name (e.g., "Limit the maximum network bandwidth for BITS background transfers")</summary>
    public string Name { get; set; } = "";
    /// <summary>GPResult-style category path (e.g., "Network/Background Intelligent Transfer Service (BITS)")</summary>
    public string CategoryPath { get; set; } = "";
    /// <summary>Full description / help text from ADMX</summary>
    public string Description { get; set; } = "";
    /// <summary>Minimum supported OS (e.g., "At least Windows Vista")</summary>
    public string SupportedOn { get; set; } = "";
    /// <summary>Machine or User scope</summary>
    public string Scope { get; set; } = "";

    /// <summary>DWORD value written when policy is Enabled (null = not defined in ADMX)</summary>
    public int? EnabledValue { get; set; }
    /// <summary>DWORD value written when policy is Disabled (null = not defined in ADMX)</summary>
    public int? DisabledValue { get; set; }
    /// <summary>True for bare policies (implicit DWORD 1=Enabled, absent=Not Configured)</summary>
    public bool IsBare { get; set; }
    /// <summary>Enum value map: DWORD value → display string (e.g., "1" → "Quick scan")</summary>
    public Dictionary<string, string>? EnumValues { get; set; }
}

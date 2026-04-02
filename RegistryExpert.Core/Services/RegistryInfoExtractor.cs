using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using RegistryParser.Abstractions;
using RegistryParser.Cells;
using RegistryExpert.Core.Models;

namespace RegistryExpert.Core
{
    /// <summary>
    /// Extracts useful information from registry hives
    /// </summary>
    public class RegistryInfoExtractor : IDisposable
    {
        private readonly OfflineRegistryParser _parser;
        
        // Cache for frequently accessed registry keys
        private readonly Dictionary<string, RegistryKey?> _keyCache = new();
        private readonly Dictionary<string, string?> _valueCache = new();
        private readonly string _currentControlSet;

        // Lazy-cached disk partition registry and active storage volume DiskIds
        private List<DiskPartitionInfo>? _diskPartitionRegistry;
        private HashSet<string>? _activeStorageVolumeDiskIds;

        public RegistryInfoExtractor(OfflineRegistryParser parser)
        {
            _parser = parser;

            // Determine the active control set from Select\Current
            _currentControlSet = "ControlSet001"; // default
            try
            {
                var selectKey = _parser.GetKey("Select");
                var currentValue = selectKey?.Values.FirstOrDefault(v => v.ValueName == "Current")?.ValueData?.ToString();
                if (int.TryParse(currentValue, out int csNumber) && csNumber >= 1 && csNumber <= 3)
                {
                    _currentControlSet = $"ControlSet{csNumber:D3}";
                }
            }
            catch { /* default to ControlSet001 */ }
        }

        public void Dispose()
        {
            _keyCache.Clear();
            _valueCache.Clear();
            _diskPartitionRegistry = null;
            _activeStorageVolumeDiskIds = null;
        }

        /// <summary>
        /// Get a cached registry key
        /// </summary>
        private RegistryKey? GetCachedKey(string path)
        {
            if (_keyCache.TryGetValue(path, out var cached))
                return cached;
            var key = _parser.GetKey(path);
            _keyCache[path] = key;
            return key;
        }


        #region Helper Methods

        private string? GetValue(string keyPath, string valueName)
        {
            // Use cache key that includes value name
            var cacheKey = $"{keyPath}\\{valueName}";
            if (_valueCache.TryGetValue(cacheKey, out var cached))
                return cached;
            string? result = null;
            try
            {
                var key = GetCachedKey(keyPath);
                result = key?.Values.FirstOrDefault(v => v.ValueName == valueName)?.ValueData?.ToString();
            }
            catch { }
            _valueCache[cacheKey] = result;
            return result;
        }

        private byte[]? GetBinaryValue(string keyPath, string valueName)
        {
            try
            {
                var key = GetCachedKey(keyPath);
                return key?.Values.FirstOrDefault(v => v.ValueName == valueName)?.ValueDataRaw;
            }
            catch
            {
                return null;
            }
        }

        private string GetStartTypeName(string startType)
        {
            return startType switch
            {
                "0" => "Boot",
                "1" => "System",
                "2" => "Automatic",
                "3" => "Manual",
                "4" => "Disabled",
                _ => startType
            };
        }

        private string GetNetworkCategory(string category)
        {
            return category switch
            {
                "0" => "Public",
                "1" => "Private",
                "2" => "Domain",
                _ => category
            };
        }

        private string FormatInstallDate(string dateStr)
        {
            if (dateStr.Length == 8 && DateTime.TryParseExact(dateStr, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out var date))
            {
                return date.ToString("MM/dd/yyyy");
            }
            return dateStr;
        }

        private string FormatMilliseconds(string msStr)
        {
            if (!long.TryParse(msStr, out var ms) || ms <= 0)
                return msStr;

            var timeSpan = TimeSpan.FromMilliseconds(ms);
            if (timeSpan.TotalDays >= 1)
                return $"{timeSpan.TotalDays:F1} days ({msStr} ms)";
            if (timeSpan.TotalHours >= 1)
                return $"{timeSpan.TotalHours:F1} hours ({msStr} ms)";
            if (timeSpan.TotalMinutes >= 1)
                return $"{timeSpan.TotalMinutes:F1} minutes ({msStr} ms)";
            return $"{timeSpan.TotalSeconds:F1} seconds ({msStr} ms)";
        }



        private string TruncatePath(string? path, int maxLength)
        {
            if (string.IsNullOrEmpty(path) || path.Length <= maxLength)
                return path ?? "";
            return path.Substring(0, maxLength - 3) + "...";
        }

        private string ExtractAgentVersionInfo(string? imagePath)
        {
            if (string.IsNullOrWhiteSpace(imagePath))
                return "";

            try
            {
                // Pattern: <version>_<YYYY-MM-DD>_<HHMMSS>
                var m = System.Text.RegularExpressions.Regex.Match(
                    imagePath,
                    @"(?<ver>\d{1,4}(?:\.\d{1,6}){1,3})_(?<date>\d{4}-\d{2}-\d{2})_(?<time>\d{6})"
                );
                if (m.Success)
                {
                    var ver = m.Groups["ver"].Value;
                    var date = m.Groups["date"].Value;
                    var time = m.Groups["time"].Value;
                    var timeFmt = time.Length == 6 ? $"{time.Substring(0,2)}:{time.Substring(2,2)}:{time.Substring(4,2)}" : time;
                    return $"{ver} ({date} {timeFmt})";
                }

                // Fallback: just dotted version anywhere
                var v2 = System.Text.RegularExpressions.Regex.Match(imagePath, @"\d{1,4}(?:\.\d{1,6}){1,3}");
                if (v2.Success)
                    return v2.Value;
            }
            catch { }
            return "";
        }

        private string GetProtocolStatus(string? enabled, string? disabledByDefault)
        {
            // Enabled=0 means explicitly disabled
            // Enabled=1 (or not set with DisabledByDefault=0) means enabled
            // DisabledByDefault=1 means disabled unless Enabled=1
            
            if (enabled == "0")
                return "Disabled";
            if (enabled == "1")
                return "Enabled";
            if (disabledByDefault == "1")
                return "Disabled";
            if (disabledByDefault == "0")
                return "Enabled";
            
            return "Not Configured";
        }

        private string DetermineOverallProtocolStatus(RegistryParser.Abstractions.RegistryKey protocolKey)
        {
            if (protocolKey.SubKeys == null) return "Not Configured";
            
            bool clientEnabled = true;
            bool serverEnabled = true;
            bool hasExplicitSettings = false;
            
            var clientKey = protocolKey.SubKeys.FirstOrDefault(s => s.KeyName.Equals("Client", StringComparison.OrdinalIgnoreCase));
            if (clientKey != null)
            {
                var enabled = clientKey.Values.FirstOrDefault(v => v.ValueName == "Enabled")?.ValueData?.ToString();
                var disabledByDefault = clientKey.Values.FirstOrDefault(v => v.ValueName == "DisabledByDefault")?.ValueData?.ToString();
                
                if (enabled == "0" || disabledByDefault == "1")
                {
                    clientEnabled = false;
                    hasExplicitSettings = true;
                }
                else if (enabled == "1" || disabledByDefault == "0")
                {
                    hasExplicitSettings = true;
                }
            }
            
            var serverKey = protocolKey.SubKeys.FirstOrDefault(s => s.KeyName.Equals("Server", StringComparison.OrdinalIgnoreCase));
            if (serverKey != null)
            {
                var enabled = serverKey.Values.FirstOrDefault(v => v.ValueName == "Enabled")?.ValueData?.ToString();
                var disabledByDefault = serverKey.Values.FirstOrDefault(v => v.ValueName == "DisabledByDefault")?.ValueData?.ToString();
                
                if (enabled == "0" || disabledByDefault == "1")
                {
                    serverEnabled = false;
                    hasExplicitSettings = true;
                }
                else if (enabled == "1" || disabledByDefault == "0")
                {
                    hasExplicitSettings = true;
                }
            }
            
            if (!hasExplicitSettings)
                return "Not Configured";
            if (clientEnabled && serverEnabled)
                return "Enabled";
            if (!clientEnabled && !serverEnabled)
                return "Disabled";
            return "Partial";
        }

        #endregion

        #region Structured Analysis Methods

        /// <summary>
        /// Get full analysis as structured sections
        /// </summary>
        public List<AnalysisSection> GetFullAnalysis()
        {
            var sections = new List<AnalysisSection>();
            sections.AddRange(GetSystemAnalysis());
            sections.AddRange(GetUserAnalysis());
            sections.AddRange(GetServicesSummary()); // Summary only, not full list
            sections.AddRange(GetNetworkAnalysis());
            return sections;
        }

        /// <summary>
        /// Get services summary (not full list) for Full Analysis view
        /// </summary>
        private List<AnalysisSection> GetServicesSummary()
        {
            var sections = new List<AnalysisSection>();
            var services = GetAllServices();
            
            if (services.Count > 0)
            {
                var summarySection = new AnalysisSection { Title = "🔧 Services Summary" };
                summarySection.Items.Add(new AnalysisItem { Name = "Total Services", Value = services.Count.ToString() });
                summarySection.Items.Add(new AnalysisItem { Name = "Disabled", Value = services.Count(s => s.IsDisabled).ToString() });
                summarySection.Items.Add(new AnalysisItem { Name = "Auto-Start", Value = services.Count(s => s.IsAutoStart).ToString() });
                summarySection.Items.Add(new AnalysisItem { Name = "Boot/System", Value = services.Count(s => s.IsBoot || s.IsSystem).ToString() });
                summarySection.Items.Add(new AnalysisItem { Name = "Manual", Value = services.Count(s => s.IsManual).ToString() });
                sections.Add(summarySection);
            }
            
            return sections;
        }

        /// <summary>
        /// Get system information as structured sections
        /// </summary>
        public List<AnalysisSection> GetSystemAnalysis()
        {
            var sections = new List<AnalysisSection>();

            // Hive Info Section
            var hiveSection = new AnalysisSection { Title = "📁 Hive Information" };
            hiveSection.Items.Add(new AnalysisItem 
            { 
                Name = "Hive Type", 
                Value = _parser.CurrentHiveType.ToString(),
                RegistryPath = _parser.FilePath ?? "Unknown",
                RegistryValue = $"Detected hive type: {_parser.CurrentHiveType}"
            });
            
            // Control Sets (SYSTEM hive)
            var controlSets = new[] { "ControlSet001", "ControlSet002", "ControlSet003" };
            var presentControlSets = new List<string>();
            foreach (var cs in controlSets)
            {
                var key = _parser.GetKey(cs);
                if (key != null)
                    presentControlSets.Add(cs);
            }
            if (presentControlSets.Count > 0)
                hiveSection.Items.Add(new AnalysisItem 
                { 
                    Name = "Control Sets", 
                    Value = string.Join(", ", presentControlSets),
                    RegistryPath = "Root",
                    RegistryValue = $"Present control sets: {string.Join(", ", presentControlSets)}"
                });

            var selectKey = _parser.GetKey("Select");
            if (selectKey != null)
            {
                var current = selectKey.Values.FirstOrDefault(v => v.ValueName == "Current")?.ValueData?.ToString();
                if (!string.IsNullOrEmpty(current))
                    hiveSection.Items.Add(new AnalysisItem 
                    { 
                        Name = "Current Control Set", 
                        Value = $"ControlSet00{current}",
                        RegistryPath = "Select",
                        RegistryValue = $"Current = {current}"
                    });
            }
            
            sections.Add(hiveSection);

            // Computer Info Section
            var computerSection = new AnalysisSection { Title = "💻 Computer Information" };
            var computerNamePath = $@"{_currentControlSet}\Control\ComputerName\ComputerName";
            var computerName = GetValue(computerNamePath, "ComputerName");
            if (string.IsNullOrEmpty(computerName))
            {
                computerNamePath = $@"{_currentControlSet}\Control\ComputerName\ActiveComputerName";
                computerName = GetValue(computerNamePath, "ComputerName");
            }
            if (!string.IsNullOrEmpty(computerName))
                computerSection.Items.Add(new AnalysisItem 
                { 
                    Name = "Computer Name", 
                    Value = computerName,
                    RegistryPath = computerNamePath,
                    RegistryValue = $"ComputerName = {computerName}"
                });

            var timezonePath = $@"{_currentControlSet}\Control\TimeZoneInformation";
            var timezone = GetValue(timezonePath, "TimeZoneKeyName");
            var timezoneValueName = "TimeZoneKeyName";
            if (string.IsNullOrEmpty(timezone))
            {
                timezone = GetValue(timezonePath, "StandardName");
                timezoneValueName = "StandardName";
            }
            if (!string.IsNullOrEmpty(timezone))
                computerSection.Items.Add(new AnalysisItem 
                { 
                    Name = "Time Zone", 
                    Value = timezone,
                    RegistryPath = timezonePath,
                    RegistryValue = $"{timezoneValueName} = {timezone}"
                });

            var shutdownPath = $@"{_currentControlSet}\Control\Windows";
            var shutdownTime = GetBinaryValue(shutdownPath, "ShutdownTime");
            if (shutdownTime != null && shutdownTime.Length >= 8)
            {
                try
                {
                    long ticks = BitConverter.ToInt64(shutdownTime, 0);
                    if (ticks > 0)
                    {
                        var dt = DateTime.FromFileTime(ticks);
                        computerSection.Items.Add(new AnalysisItem 
                        { 
                            Name = "Last Shutdown", 
                            Value = dt.ToString("yyyy-MM-dd HH:mm:ss"),
                            RegistryPath = shutdownPath,
                            RegistryValue = $"ShutdownTime = {dt:yyyy-MM-dd HH:mm:ss} (FILETIME: {ticks})"
                        });
                    }
                }
                catch { }
            }

            // Add BIOS Mode, BIOS Version/Date, VM Generation, TPM to Computer Information
            if (_parser.CurrentHiveType == OfflineRegistryParser.HiveType.SYSTEM)
            {
                computerSection.Items.AddRange(GetBiosAndTpmItems());
            }

            // Always add for UI consistency (will appear greyed out if empty)
            sections.Add(computerSection);

            // OS Information Section (SOFTWARE hive) - Always include for UI consistency
            string osRegPath = @"Microsoft\Windows NT\CurrentVersion";
            var osSection = new AnalysisSection { Title = "🪟 Build Information" };
            var productName = GetValue(osRegPath, "ProductName");
            var buildNumber = GetValue(osRegPath, "CurrentBuild");  // Fetch early for Windows 11 detection

            // Windows 11 detection: Build 22000+ is Windows 11, even if ProductName says "Windows 10"
            var displayProductName = productName;
            if (!string.IsNullOrEmpty(buildNumber) && 
                int.TryParse(buildNumber, out int buildNum) && 
                buildNum >= 22000 &&
                displayProductName != null && 
                displayProductName.Contains("Windows 10", StringComparison.OrdinalIgnoreCase))
            {
                displayProductName = displayProductName.Replace("Windows 10", "Windows 11", StringComparison.OrdinalIgnoreCase);
            }

            if (!string.IsNullOrEmpty(productName))
                osSection.Items.Add(new AnalysisItem 
                { 
                    Name = "Product Name", 
                    Value = displayProductName!,           // Shows "Windows 11 ..." if build >= 22000
                    RegistryPath = osRegPath,
                    RegistryValue = $"ProductName = {productName}"  // Shows real registry value
                });

            var editionId = GetValue(osRegPath, "EditionID");
            if (!string.IsNullOrEmpty(editionId))
                osSection.Items.Add(new AnalysisItem 
                { 
                    Name = "Edition", 
                    Value = editionId,
                    RegistryPath = osRegPath,
                    RegistryValue = $"EditionID = {editionId}"
                });

            var displayVersion = GetValue(osRegPath, "DisplayVersion")
                ?? GetValue(osRegPath, "ReleaseId");
            var displayVersionName = GetValue(osRegPath, "DisplayVersion") != null ? "DisplayVersion" : "ReleaseId";
            if (!string.IsNullOrEmpty(displayVersion))
                osSection.Items.Add(new AnalysisItem 
                { 
                    Name = "Version", 
                    Value = displayVersion,
                    RegistryPath = osRegPath,
                    RegistryValue = $"{displayVersionName} = {displayVersion}"
                });

            // buildNumber already fetched above for Windows 11 detection
            var ubr = GetValue(osRegPath, "UBR");
            if (!string.IsNullOrEmpty(buildNumber))
            {
                var fullBuild = !string.IsNullOrEmpty(ubr) ? $"{buildNumber}.{ubr}" : buildNumber;
                osSection.Items.Add(new AnalysisItem 
                { 
                    Name = "Build", 
                    Value = fullBuild,
                    RegistryPath = osRegPath,
                    RegistryValue = !string.IsNullOrEmpty(ubr) ? $"CurrentBuild = {buildNumber}, UBR = {ubr}" : $"CurrentBuild = {buildNumber}"
                });
            }

            var installDate = GetValue(osRegPath, "InstallDate");
            if (!string.IsNullOrEmpty(installDate) && long.TryParse(installDate, out long timestamp))
            {
                var dt = DateTimeOffset.FromUnixTimeSeconds(timestamp).LocalDateTime;
                osSection.Items.Add(new AnalysisItem 
                { 
                    Name = "Install Date", 
                    Value = dt.ToString("yyyy-MM-dd HH:mm:ss"),
                    RegistryPath = osRegPath,
                    RegistryValue = $"InstallDate = {installDate} (Unix timestamp)"
                });
            }

            var registeredOwner = GetValue(osRegPath, "RegisteredOwner");
            if (!string.IsNullOrEmpty(registeredOwner))
                osSection.Items.Add(new AnalysisItem 
                { 
                    Name = "Registered Owner", 
                    Value = registeredOwner,
                    RegistryPath = osRegPath,
                    RegistryValue = $"RegisteredOwner = {registeredOwner}"
                });

            var productId = GetValue(osRegPath, "ProductId");
            if (!string.IsNullOrEmpty(productId))
                osSection.Items.Add(new AnalysisItem 
                { 
                    Name = "Product ID", 
                    Value = productId,
                    RegistryPath = osRegPath,
                    RegistryValue = $"ProductId = {productId}"
                });

            var lcuVer = GetValue(osRegPath, "LCUVer");
            if (!string.IsNullOrEmpty(lcuVer))
                osSection.Items.Add(new AnalysisItem 
                { 
                    Name = "LCUVer", 
                    Value = lcuVer,
                    RegistryPath = osRegPath,
                    RegistryValue = $"LCUVer = {lcuVer}"
                });

            // Always add Build Information section so it appears greyed out when not available
            sections.Add(osSection);

            // CPU Hyper-Threading Section
            string memMgmtPath = $@"{_currentControlSet}\Control\Session Manager\Memory Management";
            var memMgmtKey = _parser.GetKey(memMgmtPath);
            var htSection = new AnalysisSection { Title = "🔄 CPU Hyper-Threading" };
            if (memMgmtKey != null)
            {
                // FeatureSettings
                var featureSettings = memMgmtKey.Values.FirstOrDefault(v => v.ValueName == "FeatureSettings")?.ValueData?.ToString() ?? "";
                htSection.Items.Add(new AnalysisItem 
                { 
                    Name = "FeatureSettings", 
                    Value = string.IsNullOrEmpty(featureSettings) ? "Not Configured" : featureSettings,
                    RegistryPath = memMgmtPath,
                    RegistryValue = $"FeatureSettings = {featureSettings}"
                });

                // FeatureSettingsOverride
                var featureOverride = memMgmtKey.Values.FirstOrDefault(v => v.ValueName == "FeatureSettingsOverride")?.ValueData?.ToString() ?? "";
                htSection.Items.Add(new AnalysisItem 
                { 
                    Name = "FeatureSettingsOverride", 
                    Value = string.IsNullOrEmpty(featureOverride) ? "Not Configured" : featureOverride,
                    RegistryPath = memMgmtPath,
                    RegistryValue = $"FeatureSettingsOverride = {featureOverride}"
                });

                // FeatureSettingsOverrideMask
                var featureMask = memMgmtKey.Values.FirstOrDefault(v => v.ValueName == "FeatureSettingsOverrideMask")?.ValueData?.ToString() ?? "";
                htSection.Items.Add(new AnalysisItem 
                { 
                    Name = "FeatureSettingsOverrideMask", 
                    Value = string.IsNullOrEmpty(featureMask) ? "Not Configured" : featureMask,
                    RegistryPath = memMgmtPath,
                    RegistryValue = $"FeatureSettingsOverrideMask = {featureMask}"
                });
            }
            sections.Add(htSection); // Always add for UI consistency

            // Crash Dump Configuration (SYSTEM hive)
            var crashControlKey = _parser.GetKey($@"{_currentControlSet}\Control\CrashControl");
            var dumpSection = new AnalysisSection { Title = "💥 Crash Dump Configuration" };
            if (crashControlKey != null)
            {
                // CrashDumpEnabled
                var crashDumpEnabled = crashControlKey.Values.FirstOrDefault(v => v.ValueName == "CrashDumpEnabled")?.ValueData?.ToString() ?? "";
                var dumpTypeText = crashDumpEnabled switch
                {
                    "0" => "0 - None",
                    "1" => "1 - Complete Memory Dump",
                    "2" => "2 - Kernel Memory Dump",
                    "3" => "3 - Small Memory Dump (256KB)",
                    "7" => "7 - Automatic Memory Dump",
                    _ when string.IsNullOrEmpty(crashDumpEnabled) => "Not Configured",
                    _ => crashDumpEnabled
                };
                dumpSection.Items.Add(new AnalysisItem 
                { 
                    Name = "CrashDumpEnabled", 
                    Value = dumpTypeText,
                    RegistryPath = $@"{_currentControlSet}\Control\CrashControl",
                    RegistryValue = $"CrashDumpEnabled = {crashDumpEnabled}"
                });

                // DumpFile path
                var dumpFile = crashControlKey.Values.FirstOrDefault(v => v.ValueName == "DumpFile")?.ValueData?.ToString() ?? "";
                if (!string.IsNullOrEmpty(dumpFile))
                {
                    dumpSection.Items.Add(new AnalysisItem 
                    { 
                        Name = "DumpFile", 
                        Value = dumpFile,
                        RegistryPath = $@"{_currentControlSet}\Control\CrashControl",
                        RegistryValue = $"DumpFile = {dumpFile}"
                    });
                }

                // MinidumpDir
                var minidumpDir = crashControlKey.Values.FirstOrDefault(v => v.ValueName == "MinidumpDir")?.ValueData?.ToString() ?? "";
                if (!string.IsNullOrEmpty(minidumpDir))
                {
                    dumpSection.Items.Add(new AnalysisItem 
                    { 
                        Name = "MinidumpDir", 
                        Value = minidumpDir,
                        RegistryPath = $@"{_currentControlSet}\Control\CrashControl",
                        RegistryValue = $"MinidumpDir = {minidumpDir}"
                    });
                }

                // AutoReboot
                var autoReboot = crashControlKey.Values.FirstOrDefault(v => v.ValueName == "AutoReboot")?.ValueData?.ToString() ?? "";
                dumpSection.Items.Add(new AnalysisItem 
                { 
                    Name = "AutoReboot", 
                    Value = autoReboot == "1" ? "1 - Enabled" : (autoReboot == "0" ? "0 - Disabled" : "Not Configured"),
                    RegistryPath = $@"{_currentControlSet}\Control\CrashControl",
                    RegistryValue = $"AutoReboot = {autoReboot}"
                });

                // Overwrite
                var overwrite = crashControlKey.Values.FirstOrDefault(v => v.ValueName == "Overwrite")?.ValueData?.ToString() ?? "";
                dumpSection.Items.Add(new AnalysisItem 
                { 
                    Name = "Overwrite", 
                    Value = overwrite == "1" ? "1 - Overwrite existing file" : (overwrite == "0" ? "0 - Keep existing file" : "Not Configured"),
                    RegistryPath = $@"{_currentControlSet}\Control\CrashControl",
                    RegistryValue = $"Overwrite = {overwrite}"
                });

                // LogEvent
                var logEvent = crashControlKey.Values.FirstOrDefault(v => v.ValueName == "LogEvent")?.ValueData?.ToString() ?? "";
                dumpSection.Items.Add(new AnalysisItem 
                { 
                    Name = "LogEvent", 
                    Value = logEvent == "1" ? "1 - Log to System Event Log" : (logEvent == "0" ? "0 - Do not log" : "Not Configured"),
                    RegistryPath = $@"{_currentControlSet}\Control\CrashControl",
                    RegistryValue = $"LogEvent = {logEvent}"
                });

                // SendAlert
                var sendAlert = crashControlKey.Values.FirstOrDefault(v => v.ValueName == "SendAlert")?.ValueData?.ToString() ?? "";
                if (!string.IsNullOrEmpty(sendAlert))
                {
                    dumpSection.Items.Add(new AnalysisItem 
                    { 
                        Name = "SendAlert", 
                        Value = sendAlert == "1" ? "1 - Send administrative alert" : "0 - No alert",
                        RegistryPath = $@"{_currentControlSet}\Control\CrashControl",
                        RegistryValue = $"SendAlert = {sendAlert}"
                    });
                }

                // DumpFilters
                var dumpFilters = crashControlKey.Values.FirstOrDefault(v => v.ValueName == "DumpFilters")?.ValueData?.ToString() ?? "";
                if (!string.IsNullOrEmpty(dumpFilters))
                {
                    dumpSection.Items.Add(new AnalysisItem 
                    { 
                        Name = "DumpFilters", 
                        Value = dumpFilters,
                        RegistryPath = $@"{_currentControlSet}\Control\CrashControl",
                        RegistryValue = $"DumpFilters = {dumpFilters}"
                    });
                }
            }

            // Virtual Memory / Paging File Configuration (same tab)
            // Reuse memMgmtPath and memMgmtKey from CPU Hyper-Threading section above

            if (memMgmtKey != null)
            {
                // PagingFiles (REG_MULTI_SZ) - the configured paging files
                // Format: "<path> <initial_MB> <max_MB>" per entry; "0 0" = system managed
                var pagingFilesRaw = memMgmtKey.Values.FirstOrDefault(v => v.ValueName == "PagingFiles")?.ValueData?.ToString() ?? "";
                if (!string.IsNullOrEmpty(pagingFilesRaw))
                {
                    // Multi-string entries may be separated by newlines or stored as one string
                    var entries = pagingFilesRaw.Split(new[] { '\0', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var entry in entries)
                    {
                        var parts = entry.Trim().Split(' ', StringSplitOptions.RemoveEmptyEntries);
                        if (parts.Length >= 1)
                        {
                            var path = parts[0];
                            var initial = parts.Length >= 2 ? parts[1] : "?";
                            var max = parts.Length >= 3 ? parts[2] : "?";

                            // Extract drive letter for display; ?:\ is a system default placeholder
                            var isWildcard = path.Length >= 2 && path[0] == '?' && path[1] == ':';
                            var driveLetter = path.Length >= 2 && path[1] == ':' ? path[0].ToString().ToUpper() + ":" : path;

                            string displayName;
                            string sizeDesc;

                            if (isWildcard)
                            {
                                // ?:\pagefile.sys is ambiguous - could mean auto-managed or no paging file;
                                // show the raw value neutrally without interpretation
                                displayName = "Paging File (Default)";
                                sizeDesc = $"{path}  [Initial: {initial}, Max: {max}]";
                            }
                            else
                            {
                                displayName = $"Paging File ({driveLetter})";
                                if (initial == "0" && max == "0")
                                    sizeDesc = $"System managed  [{path}]";
                                else if (int.TryParse(initial, out var initMb) && int.TryParse(max, out var maxMb) && initMb > 0 && maxMb > 0)
                                    sizeDesc = $"Custom: {initMb} - {maxMb} MB  [{path}]";
                                else
                                    sizeDesc = $"Initial={initial}, Max={max}  [{path}]";
                            }

                            dumpSection.Items.Add(new AnalysisItem
                            {
                                Name = displayName,
                                Value = sizeDesc,
                                RegistryPath = memMgmtPath,
                                RegistryValue = $"PagingFiles = {pagingFilesRaw}"
                            });
                        }
                    }
                }
                else
                {
                    dumpSection.Items.Add(new AnalysisItem
                    {
                        Name = "Paging Files",
                        Value = "No paging files configured",
                        RegistryPath = memMgmtPath,
                        RegistryValue = "PagingFiles = (empty)"
                    });
                }

                // ExistingPageFiles (REG_MULTI_SZ) - page files active at last boot
                var existingPfRaw = memMgmtKey.Values.FirstOrDefault(v => v.ValueName == "ExistingPageFiles")?.ValueData?.ToString() ?? "";
                if (!string.IsNullOrEmpty(existingPfRaw))
                {
                    // Strip \??\ prefix for display
                    var existingEntries = existingPfRaw.Split(new[] { '\0', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    var displayPaths = existingEntries.Select(e => e.Trim().Replace(@"\??\", "")).ToList();
                    dumpSection.Items.Add(new AnalysisItem
                    {
                        Name = "Active Page Files (at boot)",
                        Value = string.Join(", ", displayPaths),
                        RegistryPath = memMgmtPath,
                        RegistryValue = $"ExistingPageFiles = {existingPfRaw}"
                    });
                }

                // ClearPageFileAtShutdown
                var clearPf = memMgmtKey.Values.FirstOrDefault(v => v.ValueName == "ClearPageFileAtShutdown")?.ValueData?.ToString() ?? "";
                dumpSection.Items.Add(new AnalysisItem
                {
                    Name = "Clear Page File at Shutdown",
                    Value = clearPf == "1" ? "1 - Enabled (secure)" : (clearPf == "0" ? "0 - Disabled" : "Not Configured"),
                    RegistryPath = memMgmtPath,
                    RegistryValue = $"ClearPageFileAtShutdown = {clearPf}"
                });

                // DisablePagingExecutive
                var disablePaging = memMgmtKey.Values.FirstOrDefault(v => v.ValueName == "DisablePagingExecutive")?.ValueData?.ToString() ?? "";
                dumpSection.Items.Add(new AnalysisItem
                {
                    Name = "Disable Paging Executive",
                    Value = disablePaging == "1" ? "1 - Kernel kept in RAM" : (disablePaging == "0" ? "0 - Kernel can be paged" : "Not Configured"),
                    RegistryPath = memMgmtPath,
                    RegistryValue = $"DisablePagingExecutive = {disablePaging}"
                });

                // PhysicalAddressExtension (PAE)
                var pae = memMgmtKey.Values.FirstOrDefault(v => v.ValueName == "PhysicalAddressExtension")?.ValueData?.ToString() ?? "";
                if (!string.IsNullOrEmpty(pae))
                {
                    dumpSection.Items.Add(new AnalysisItem
                    {
                        Name = "Physical Address Extension (PAE)",
                        Value = pae == "1" ? "1 - Enabled" : "0 - Disabled",
                        RegistryPath = memMgmtPath,
                        RegistryValue = $"PhysicalAddressExtension = {pae}"
                    });
                }

                // LargeSystemCache
                var largeCache = memMgmtKey.Values.FirstOrDefault(v => v.ValueName == "LargeSystemCache")?.ValueData?.ToString() ?? "";
                dumpSection.Items.Add(new AnalysisItem
                {
                    Name = "Large System Cache",
                    Value = largeCache == "1" ? "1 - Optimize for file sharing" : (largeCache == "0" ? "0 - Optimize for applications" : "Not Configured"),
                    RegistryPath = memMgmtPath,
                    RegistryValue = $"LargeSystemCache = {largeCache}"
                });

                // Pool sizes (non-default values are noteworthy)
                var nonPagedPoolSize = memMgmtKey.Values.FirstOrDefault(v => v.ValueName == "NonPagedPoolSize")?.ValueData?.ToString() ?? "";
                var pagedPoolSize = memMgmtKey.Values.FirstOrDefault(v => v.ValueName == "PagedPoolSize")?.ValueData?.ToString() ?? "";
                if (!string.IsNullOrEmpty(nonPagedPoolSize) && nonPagedPoolSize != "0")
                {
                    dumpSection.Items.Add(new AnalysisItem
                    {
                        Name = "Non-Paged Pool Size",
                        Value = $"{nonPagedPoolSize} bytes (custom)",
                        RegistryPath = memMgmtPath,
                        RegistryValue = $"NonPagedPoolSize = {nonPagedPoolSize}"
                    });
                }
                if (!string.IsNullOrEmpty(pagedPoolSize) && pagedPoolSize != "0")
                {
                    dumpSection.Items.Add(new AnalysisItem
                    {
                        Name = "Paged Pool Size",
                        Value = $"{pagedPoolSize} bytes (custom)",
                        RegistryPath = memMgmtPath,
                        RegistryValue = $"PagedPoolSize = {pagedPoolSize}"
                    });
                }
            }
            else
            {
                dumpSection.Items.Add(new AnalysisItem
                {
                    Name = "Virtual Memory",
                    Value = "Memory Management key not found",
                    RegistryPath = memMgmtPath,
                    RegistryValue = "(key not found)",
                    IsWarning = true
                });
            }

            sections.Add(dumpSection); // Always add for UI consistency

            // Guest Agent section - content depends on hive type
            var guestSection = new AnalysisSection { Title = "☁️ Guest Agent" };
            
            if (_parser.CurrentHiveType == OfflineRegistryParser.HiveType.SOFTWARE)
            {
                // SOFTWARE hive: Show VmId, VmType, and grayed out agent status
                string azurePath = @"Microsoft\Windows Azure";
                var azureKey = _parser.GetKey(azurePath);
                
                // Get VmId
                var vmId = GetValue(azurePath, "VmId");
                guestSection.Items.Add(new AnalysisItem
                {
                    Name = "VmId",
                    Value = !string.IsNullOrEmpty(vmId) ? vmId : "Not Found",
                    RegistryPath = azurePath,
                    RegistryValue = !string.IsNullOrEmpty(vmId) ? $"VmId = {vmId}" : "VmId not present"
                });
                
                // Get VmType if present
                var vmType = GetValue(azurePath, "VmType");
                if (!string.IsNullOrEmpty(vmType))
                {
                    guestSection.Items.Add(new AnalysisItem
                    {
                        Name = "VmType",
                        Value = vmType,
                        RegistryPath = azurePath,
                        RegistryValue = $"VmType = {vmType}"
                    });
                }
                
                // Extensions are now shown via separate sub-tab (↳ Extensions)
            }
            else
            {
                // SYSTEM hive: Show agent service info
                string waGuestAgentPath = $@"{_currentControlSet}\Services\WindowsAzureGuestAgent";
                var waImagePath = GetValue(waGuestAgentPath, "ImagePath") ?? "";
                var waVersion = ExtractAgentVersionInfo(waImagePath);
                guestSection.Items.Add(new AnalysisItem
                {
                    Name = "WindowsAzureGuestAgent",
                    Value = string.IsNullOrEmpty(waImagePath) ? "Not Found" : (string.IsNullOrEmpty(waVersion) ? TruncatePath(waImagePath, 100) : waVersion),
                    RegistryPath = waGuestAgentPath,
                    RegistryValue = string.IsNullOrEmpty(waImagePath) ? "ImagePath missing" : $"ImagePath = {TruncatePath(waImagePath, 140)}"
                });

                string rdAgentPath = $@"{_currentControlSet}\Services\RdAgent";
                var rdImagePath = GetValue(rdAgentPath, "ImagePath") ?? "";
                var rdVersion = ExtractAgentVersionInfo(rdImagePath);
                guestSection.Items.Add(new AnalysisItem
                {
                    Name = "RdAgent",
                    Value = string.IsNullOrEmpty(rdImagePath) ? "Not Found" : (string.IsNullOrEmpty(rdVersion) ? TruncatePath(rdImagePath, 100) : rdVersion),
                    RegistryPath = rdAgentPath,
                    RegistryValue = string.IsNullOrEmpty(rdImagePath) ? "ImagePath missing" : $"ImagePath = {TruncatePath(rdImagePath, 140)}"
                });
            }

            sections.Add(guestSection);

            // System Time Configuration (SYSTEM hive) - Windows Time Service (w32time)
            sections.Add(GetSystemTimeAnalysis());

            // Device Manager (SYSTEM hive)
            sections.AddRange(GetDeviceManagerAnalysis());

            return sections;
        }

        /// <summary>
        /// Get Device Manager analysis - devices grouped by device class from SYSTEM hive
        /// </summary>
        public List<AnalysisSection> GetDeviceManagerAnalysis()
        {
            var sections = new List<AnalysisSection>();
            var section = new AnalysisSection { Title = "\U0001f5a5\ufe0f Device Manager" };

            var enumKey = _parser.GetKey($@"{_currentControlSet}\Enum");
            if (enumKey?.SubKeys == null)
            {
                sections.Add(section);
                return sections;
            }

            // Collect all devices grouped by ClassGUID
            var devicesByClass = new Dictionary<string, List<(string DisplayName, string RegistryPath, List<(string Name, string Value, string ValueName)> Properties, string DriverRegistryPath, List<(string Name, string Value, string ValueName)> DriverProperties)>>(StringComparer.OrdinalIgnoreCase);

            foreach (var busType in enumKey.SubKeys)
            {
                if (busType.SubKeys == null) continue;
                foreach (var device in busType.SubKeys)
                {
                    if (device.SubKeys == null) continue;
                    foreach (var instance in device.SubKeys)
                    {
                        var classGuid = instance.Values.FirstOrDefault(v => v.ValueName == "ClassGUID")?.ValueData?.ToString();
                        var deviceDesc = instance.Values.FirstOrDefault(v => v.ValueName == "DeviceDesc")?.ValueData?.ToString();
                        var driverValue = instance.Values.FirstOrDefault(v => v.ValueName == "Driver")?.ValueData?.ToString();

                        // Determine if this is an "unknown" device (missing ClassGUID or Driver)
                        bool isUnknown = string.IsNullOrEmpty(classGuid) || string.IsNullOrEmpty(driverValue);

                        var friendlyName = instance.Values.FirstOrDefault(v => v.ValueName == "FriendlyName")?.ValueData?.ToString();

                        // Use "__unknown__" as the grouping key for unknown devices
                        var groupKey = isUnknown ? "__unknown__" : classGuid!;

                        // Extract display name: prefer FriendlyName, fall back to DeviceDesc, then instance key name
                        string displayName;
                        if (!string.IsNullOrEmpty(friendlyName))
                        {
                            var lastSemi = friendlyName.LastIndexOf(';');
                            displayName = lastSemi >= 0 ? friendlyName.Substring(lastSemi + 1) : friendlyName;
                        }
                        else if (!string.IsNullOrEmpty(deviceDesc))
                        {
                            var lastSemi = deviceDesc.LastIndexOf(';');
                            displayName = lastSemi >= 0 ? deviceDesc.Substring(lastSemi + 1) : deviceDesc;
                        }
                        else
                        {
                            // Try to find a display name from sibling instances under the same device key
                            string? siblingName = null;
                            if (device.SubKeys != null)
                            {
                                foreach (var sibling in device.SubKeys)
                                {
                                    if (sibling.KeyName == instance.KeyName) continue;
                                    var sibFriendly = sibling.Values?.FirstOrDefault(v => v.ValueName == "FriendlyName")?.ValueData?.ToString();
                                    var sibDesc = sibling.Values?.FirstOrDefault(v => v.ValueName == "DeviceDesc")?.ValueData?.ToString();
                                    var raw = sibFriendly ?? sibDesc;
                                    if (!string.IsNullOrEmpty(raw))
                                    {
                                        var lastSemi = raw.LastIndexOf(';');
                                        siblingName = lastSemi >= 0 ? raw.Substring(lastSemi + 1) : raw;
                                        break;
                                    }
                                }
                            }

                            if (siblingName != null)
                                displayName = $"{siblingName} (not configured)";
                            else
                                displayName = $@"{busType.KeyName}\{device.KeyName}";
                        }

                        var registryPath = $@"{_currentControlSet}\Enum\{busType.KeyName}\{device.KeyName}\{instance.KeyName}";

                        // Collect device properties
                        var properties = new List<(string Name, string Value, string ValueName)>();
                        var propertyNames = new[] { "DeviceDesc", "FriendlyName", "ClassGUID", "Class", "Driver", "Service", "Mfg", "HardwareID", "CompatibleIDs", "ConfigFlags", "ContainerID", "LocationInformation" };
                        foreach (var propName in propertyNames)
                        {
                            var val = instance.Values.FirstOrDefault(v => v.ValueName == propName)?.ValueData?.ToString();
                            if (!string.IsNullOrEmpty(val))
                            {
                                // Clean up localized resource strings for display
                                var displayVal = val;
                                if (propName == "DeviceDesc" || propName == "FriendlyName" || propName == "Mfg")
                                {
                                    var semi = displayVal.LastIndexOf(';');
                                    if (semi >= 0) displayVal = displayVal.Substring(semi + 1);
                                }
                                properties.Add((propName, displayVal, propName));
                            }
                        }

                        // Add enabled/disabled status from ConfigFlags
                        var configFlags = instance.Values.FirstOrDefault(v => v.ValueName == "ConfigFlags")?.ValueData?.ToString();
                        bool isEnabled = true;
                        if (!string.IsNullOrEmpty(configFlags) && int.TryParse(configFlags, out int flags))
                        {
                            isEnabled = (flags & 0x01) == 0;
                        }
                        properties.Add(("Status", isEnabled ? "Enabled" : "Disabled", "ConfigFlags"));
                        properties.Add(("Device Instance Path", $@"{busType.KeyName}\{device.KeyName}\{instance.KeyName}", ""));

                        // Collect driver class details from ControlSet001\Control\Class\{Driver}
                        var driverProperties = new List<(string Name, string Value, string ValueName)>();
                        string driverRegistryPath = "";
                        if (!string.IsNullOrEmpty(driverValue))
                        {
                            var driverKeyPath = $@"{_currentControlSet}\Control\Class\{driverValue}";
                            var driverKey = GetCachedKey(driverKeyPath);
                            if (driverKey?.Values != null)
                            {
                                driverRegistryPath = driverKeyPath;
                                foreach (var val in driverKey.Values.OrderBy(v => v.ValueName))
                                {
                                    var valName = val.ValueName;
                                    if (string.IsNullOrEmpty(valName)) valName = "(Default)";
                                    var valData = val.ValueData?.ToString() ?? "";
                                    driverProperties.Add((valName, valData, valName));
                                }
                            }
                        }

                        if (!devicesByClass.ContainsKey(groupKey))
                            devicesByClass[groupKey] = new();
                        devicesByClass[groupKey].Add((displayName, registryPath, properties, driverRegistryPath, driverProperties));
                    }
                }
            }

            // Look up class names and build hierarchical section items
            var classItems = new List<(string ClassName, AnalysisItem Item)>();
            foreach (var kvp in devicesByClass)
            {
                var groupKey = kvp.Key;
                var devices = kvp.Value;

                string className;
                string classKeyPath;
                string classGuidDisplay;

                if (groupKey == "__unknown__")
                {
                    className = "Unknown Devices";
                    classKeyPath = "";
                    classGuidDisplay = "N/A";
                }
                else
                {
                    // Look up class display name from ControlSet001\Control\Class\{ClassGUID}
                    classKeyPath = $@"{_currentControlSet}\Control\Class\{groupKey}";
                    className = GetValue(classKeyPath, "Class") ?? groupKey;
                    classGuidDisplay = groupKey;
                }

                var classItem = new AnalysisItem
                {
                    Name = $"{className} ({devices.Count})",
                    Value = classGuidDisplay,
                    IsSubSection = true,
                    RegistryPath = classKeyPath,
                    RegistryValue = $"Class = {className}",
                    SubItems = new List<AnalysisItem>()
                };

                foreach (var dev in devices.OrderBy(d => d.DisplayName))
                {
                    var deviceItem = new AnalysisItem
                    {
                        Name = dev.DisplayName,
                        Value = "",
                        RegistryPath = dev.RegistryPath,
                        RegistryValue = string.Join(" | ", dev.Properties.Select(p => $"{p.Name} = {p.Value}")),
                        IsSubSection = true,
                        SubItems = dev.Properties.Select(p => new AnalysisItem
                        {
                            Name = p.Name,
                            Value = p.Value,
                            RegistryPath = dev.RegistryPath,
                            RegistryValue = p.ValueName
                        }).ToList()
                    };

                    // Add driver class details as a special marker sub-item
                    if (dev.DriverProperties.Count > 0)
                    {
                        var driverSection = new AnalysisItem
                        {
                            Name = "__DriverDetails__",
                            Value = "",
                            RegistryPath = dev.DriverRegistryPath,
                            IsSubSection = true,
                            SubItems = dev.DriverProperties.Select(p => new AnalysisItem
                            {
                                Name = p.Name,
                                Value = p.Value,
                                RegistryPath = dev.DriverRegistryPath,
                                RegistryValue = p.ValueName
                            }).ToList()
                        };
                        deviceItem.SubItems!.Add(driverSection);
                    }

                    classItem.SubItems.Add(deviceItem);
                }

                classItems.Add((className, classItem));
            }

            // Sort classes: Unknown Devices first, then alphabetically
            foreach (var ci in classItems.OrderBy(c => c.ClassName == "Unknown Devices" ? 0 : 1).ThenBy(c => c.ClassName))
            {
                section.Items.Add(ci.Item);
            }

            sections.Add(section);
            return sections;
        }

        /// <summary>
        /// Get Windows Time Service (w32time) configuration as a structured section
        /// </summary>
        public AnalysisSection GetSystemTimeAnalysis()
        {
            var timeSection = new AnalysisSection { Title = "🕐 System Time Config" };

            // w32time Parameters - Type (NTP vs NT5DS)
            string w32timeParamsPath = $@"{_currentControlSet}\Services\w32time\Parameters";
            var paramsKey = _parser.GetKey(w32timeParamsPath);
            
            if (paramsKey != null)
            {
                // Type - determines if using NTP servers or domain hierarchy (NT5DS)
                var timeType = paramsKey.Values.FirstOrDefault(v => v.ValueName == "Type")?.ValueData?.ToString() ?? "";
                var typeDescription = timeType.ToUpperInvariant() switch
                {
                    "NTP" => "NTP - Using explicit NTP servers",
                    "NT5DS" => "NT5DS - Domain time sync hierarchy",
                    "NOSYNCH" => "NoSync - Time sync disabled",
                    "ALLSYNC" => "AllSync - All available sync methods",
                    _ when string.IsNullOrEmpty(timeType) => "Not Configured",
                    _ => timeType
                };
                timeSection.Items.Add(new AnalysisItem
                {
                    Name = "Time Sync Type",
                    Value = typeDescription,
                    RegistryPath = w32timeParamsPath,
                    RegistryValue = $"Type = {timeType}"
                });

                // NtpServer - configured NTP servers
                var ntpServer = paramsKey.Values.FirstOrDefault(v => v.ValueName == "NtpServer")?.ValueData?.ToString() ?? "";
                if (!string.IsNullOrEmpty(ntpServer))
                {
                    // Parse NTP server flags (0x1 = SpecialInterval, 0x2 = UseAsFallbackOnly, 0x4 = SymmetricActive, 0x8 = Client)
                    var serverDescription = ParseNtpServerFlags(ntpServer);
                    timeSection.Items.Add(new AnalysisItem
                    {
                        Name = "NTP Server",
                        Value = serverDescription,
                        RegistryPath = w32timeParamsPath,
                        RegistryValue = $"NtpServer = {ntpServer}"
                    });
                }
            }

            // w32time Config - Poll intervals
            string w32timeConfigPath = $@"{_currentControlSet}\Services\w32time\Config";
            var configKey = _parser.GetKey(w32timeConfigPath);
            
            if (configKey != null)
            {
                // MinPollInterval (power of 2 seconds, e.g., 6 = 64 seconds)
                var minPoll = configKey.Values.FirstOrDefault(v => v.ValueName == "MinPollInterval")?.ValueData?.ToString() ?? "";
                if (!string.IsNullOrEmpty(minPoll) && int.TryParse(minPoll, out int minPollVal))
                {
                    var minPollSeconds = (int)Math.Pow(2, minPollVal);
                    timeSection.Items.Add(new AnalysisItem
                    {
                        Name = "MinPollInterval",
                        Value = $"{minPollVal} (2^{minPollVal} = {minPollSeconds} seconds)",
                        RegistryPath = w32timeConfigPath,
                        RegistryValue = $"MinPollInterval = {minPoll}"
                    });
                }

                // MaxPollInterval (power of 2 seconds, e.g., 10 = 1024 seconds)
                var maxPoll = configKey.Values.FirstOrDefault(v => v.ValueName == "MaxPollInterval")?.ValueData?.ToString() ?? "";
                if (!string.IsNullOrEmpty(maxPoll) && int.TryParse(maxPoll, out int maxPollVal))
                {
                    var maxPollSeconds = (int)Math.Pow(2, maxPollVal);
                    timeSection.Items.Add(new AnalysisItem
                    {
                        Name = "MaxPollInterval",
                        Value = $"{maxPollVal} (2^{maxPollVal} = {maxPollSeconds} seconds)",
                        RegistryPath = w32timeConfigPath,
                        RegistryValue = $"MaxPollInterval = {maxPoll}"
                    });
                }

                // UpdateInterval
                var updateInterval = configKey.Values.FirstOrDefault(v => v.ValueName == "UpdateInterval")?.ValueData?.ToString() ?? "";
                if (!string.IsNullOrEmpty(updateInterval))
                {
                    timeSection.Items.Add(new AnalysisItem
                    {
                        Name = "UpdateInterval",
                        Value = updateInterval,
                        RegistryPath = w32timeConfigPath,
                        RegistryValue = $"UpdateInterval = {updateInterval}"
                    });
                }

                // AnnounceFlags
                var announceFlags = configKey.Values.FirstOrDefault(v => v.ValueName == "AnnounceFlags")?.ValueData?.ToString() ?? "";
                if (!string.IsNullOrEmpty(announceFlags))
                {
                    var flagsDescription = ParseAnnounceFlags(announceFlags);
                    timeSection.Items.Add(new AnalysisItem
                    {
                        Name = "AnnounceFlags",
                        Value = flagsDescription,
                        RegistryPath = w32timeConfigPath,
                        RegistryValue = $"AnnounceFlags = {announceFlags}"
                    });
                }
            }

            // Time Providers - VMICTimeProvider (for Azure/Hyper-V VMs)
            string vmicProviderPath = $@"{_currentControlSet}\Services\w32time\TimeProviders\VMICTimeProvider";
            var vmicKey = _parser.GetKey(vmicProviderPath);
            
            if (vmicKey != null)
            {
                var vmicEnabled = vmicKey.Values.FirstOrDefault(v => v.ValueName == "Enabled")?.ValueData?.ToString() ?? "";
                timeSection.Items.Add(new AnalysisItem
                {
                    Name = "VMICTimeProvider",
                    Value = vmicEnabled == "1" ? "Enabled (VM host time sync)" : (vmicEnabled == "0" ? "Disabled" : "Not Configured"),
                    RegistryPath = vmicProviderPath,
                    RegistryValue = $"Enabled = {vmicEnabled}"
                });
            }
            else
            {
                timeSection.Items.Add(new AnalysisItem
                {
                    Name = "VMICTimeProvider",
                    Value = "Not Present (not a VM or provider not installed)",
                    RegistryPath = vmicProviderPath,
                    RegistryValue = "Key not found"
                });
            }

            // Time Providers - NtpClient
            string ntpClientPath = $@"{_currentControlSet}\Services\w32time\TimeProviders\NtpClient";
            var ntpClientKey = _parser.GetKey(ntpClientPath);
            
            if (ntpClientKey != null)
            {
                var ntpClientEnabled = ntpClientKey.Values.FirstOrDefault(v => v.ValueName == "Enabled")?.ValueData?.ToString() ?? "";
                timeSection.Items.Add(new AnalysisItem
                {
                    Name = "NtpClient Provider",
                    Value = ntpClientEnabled == "1" ? "Enabled" : (ntpClientEnabled == "0" ? "Disabled" : "Not Configured"),
                    RegistryPath = ntpClientPath,
                    RegistryValue = $"Enabled = {ntpClientEnabled}"
                });

                // SpecialPollInterval (seconds between polls when using 0x1 flag)
                var specialPoll = ntpClientKey.Values.FirstOrDefault(v => v.ValueName == "SpecialPollInterval")?.ValueData?.ToString() ?? "";
                if (!string.IsNullOrEmpty(specialPoll) && int.TryParse(specialPoll, out int specialPollVal))
                {
                    var timeSpan = TimeSpan.FromSeconds(specialPollVal);
                    var formatted = specialPollVal >= 3600 
                        ? $"{timeSpan.TotalHours:F1} hours ({specialPollVal} seconds)"
                        : $"{timeSpan.TotalMinutes:F0} minutes ({specialPollVal} seconds)";
                    timeSection.Items.Add(new AnalysisItem
                    {
                        Name = "SpecialPollInterval",
                        Value = formatted,
                        RegistryPath = ntpClientPath,
                        RegistryValue = $"SpecialPollInterval = {specialPoll}"
                    });
                }

                // CrossSiteSyncFlags
                var crossSiteSync = ntpClientKey.Values.FirstOrDefault(v => v.ValueName == "CrossSiteSyncFlags")?.ValueData?.ToString() ?? "";
                if (!string.IsNullOrEmpty(crossSiteSync))
                {
                    var crossSiteDescription = crossSiteSync switch
                    {
                        "0" => "0 - None",
                        "1" => "1 - PDC Only",
                        "2" => "2 - All",
                        _ => crossSiteSync
                    };
                    timeSection.Items.Add(new AnalysisItem
                    {
                        Name = "CrossSiteSyncFlags",
                        Value = crossSiteDescription,
                        RegistryPath = ntpClientPath,
                        RegistryValue = $"CrossSiteSyncFlags = {crossSiteSync}"
                    });
                }
            }

            // Time Providers - NtpServer (if this machine serves time)
            string ntpServerPath = $@"{_currentControlSet}\Services\w32time\TimeProviders\NtpServer";
            var ntpServerKey = _parser.GetKey(ntpServerPath);
            
            if (ntpServerKey != null)
            {
                var ntpServerEnabled = ntpServerKey.Values.FirstOrDefault(v => v.ValueName == "Enabled")?.ValueData?.ToString() ?? "";
                timeSection.Items.Add(new AnalysisItem
                {
                    Name = "NtpServer Provider",
                    Value = ntpServerEnabled == "1" ? "Enabled (serving time to clients)" : (ntpServerEnabled == "0" ? "Disabled" : "Not Configured"),
                    RegistryPath = ntpServerPath,
                    RegistryValue = $"Enabled = {ntpServerEnabled}"
                });
            }

            // w32time service start type
            string w32timeServicePath = $@"{_currentControlSet}\Services\w32time";
            var w32timeKey = _parser.GetKey(w32timeServicePath);
            
            if (w32timeKey != null)
            {
                var startType = w32timeKey.Values.FirstOrDefault(v => v.ValueName == "Start")?.ValueData?.ToString() ?? "";
                var startTypeDescription = startType switch
                {
                    "0" => "0 - Boot",
                    "1" => "1 - System",
                    "2" => "2 - Automatic",
                    "3" => "3 - Manual",
                    "4" => "4 - Disabled",
                    _ when string.IsNullOrEmpty(startType) => "Not Configured",
                    _ => startType
                };
                timeSection.Items.Add(new AnalysisItem
                {
                    Name = "w32time Service Start",
                    Value = startTypeDescription,
                    RegistryPath = w32timeServicePath,
                    RegistryValue = $"Start = {startType}"
                });
            }

            // Add summary/recommendation if no items found
            if (timeSection.Items.Count == 0)
            {
                timeSection.Items.Add(new AnalysisItem
                {
                    Name = "Status",
                    Value = "Windows Time configuration not found",
                    RegistryPath = w32timeParamsPath,
                    RegistryValue = "w32time service keys not present"
                });
            }

            return timeSection;
        }

        /// <summary>
        /// Get Azure VM Extensions from HandlerState (SOFTWARE hive)
        /// </summary>
        public AnalysisSection GetAzureExtensionsAnalysis()
        {
            var extensionsSection = new AnalysisSection { Title = "🔌 Extensions" };
            
            string handlerStatePath = @"Microsoft\Windows Azure\HandlerState";
            var handlerStateKey = _parser.GetKey(handlerStatePath);
            
            if (handlerStateKey?.SubKeys == null || !handlerStateKey.SubKeys.Any())
            {
                extensionsSection.Items.Add(new AnalysisItem
                {
                    Name = "Status",
                    Value = "No extensions found",
                    RegistryPath = handlerStatePath,
                    RegistryValue = "HandlerState key is empty or not present"
                });
                return extensionsSection;
            }
            
            // Each subkey under HandlerState is an extension (e.g., Microsoft.Compute.JsonADDomainExtension_1.3.14)
            foreach (var extensionKey in handlerStateKey.SubKeys.OrderBy(k => k.KeyName))
            {
                var extensionName = extensionKey.KeyName;
                var extensionPath = $@"{handlerStatePath}\{extensionName}";
                
                // Get InstallState value
                var installState = extensionKey.Values?.FirstOrDefault(v => v.ValueName == "InstallState")?.ValueData?.ToString() ?? "";
                
                // Determine status
                string status;
                if (string.IsNullOrEmpty(installState))
                {
                    status = "Unknown";
                }
                else
                {
                    status = installState.ToLowerInvariant() switch
                    {
                        "enabled" => "✓ Enabled",
                        "disabled" => "✗ Disabled",
                        "notinstalled" => "Not Installed",
                        "installed" => "Installed",
                        "failed" => "⚠ Failed",
                        _ => installState
                    };
                }
                
                // Parse extension name to get friendly name and version
                // Format: Microsoft.Compute.JsonADDomainExtension_1.3.14
                var friendlyName = extensionName;
                var version = "";
                var underscoreIndex = extensionName.LastIndexOf('_');
                if (underscoreIndex > 0)
                {
                    friendlyName = extensionName.Substring(0, underscoreIndex);
                    version = extensionName.Substring(underscoreIndex + 1);
                }
                
                // Build detail string with additional info if available
                var detailBuilder = new System.Text.StringBuilder();
                detailBuilder.AppendLine($"InstallState: {installState}");
                if (!string.IsNullOrEmpty(version))
                    detailBuilder.AppendLine($"Version: {version}");
                
                // Get other useful values
                var enableCommandState = extensionKey.Values?.FirstOrDefault(v => v.ValueName == "EnableCommandState")?.ValueData?.ToString();
                if (!string.IsNullOrEmpty(enableCommandState))
                    detailBuilder.AppendLine($"EnableCommandState: {enableCommandState}");
                
                var guestAgentMessage = extensionKey.Values?.FirstOrDefault(v => v.ValueName == "GuestAgentMessage")?.ValueData?.ToString();
                if (!string.IsNullOrEmpty(guestAgentMessage))
                    detailBuilder.AppendLine($"Message: {guestAgentMessage}");
                
                extensionsSection.Items.Add(new AnalysisItem
                {
                    Name = friendlyName,
                    Value = status,
                    RegistryPath = extensionPath,
                    RegistryValue = detailBuilder.ToString().TrimEnd()
                });
            }
            
            return extensionsSection;
        }

        /// <summary>
        /// Parse NTP server string with flags (e.g., "time.windows.com,0x8")
        /// </summary>
        private string ParseNtpServerFlags(string ntpServerValue)
        {
            if (string.IsNullOrEmpty(ntpServerValue))
                return ntpServerValue;

            var parts = ntpServerValue.Split(' ');
            var result = new List<string>();

            foreach (var part in parts)
            {
                var serverParts = part.Split(',');
                var server = serverParts[0];
                var flagInfo = "";

                if (serverParts.Length > 1)
                {
                    var flagStr = serverParts[1].Trim();
                    if (flagStr.StartsWith("0x", StringComparison.OrdinalIgnoreCase) && 
                        int.TryParse(flagStr.Substring(2), System.Globalization.NumberStyles.HexNumber, null, out int flag))
                    {
                        var flags = new List<string>();
                        if ((flag & 0x1) != 0) flags.Add("SpecialInterval");
                        if ((flag & 0x2) != 0) flags.Add("UseAsFallback");
                        if ((flag & 0x4) != 0) flags.Add("SymmetricActive");
                        if ((flag & 0x8) != 0) flags.Add("NtpClient");
                        flagInfo = flags.Count > 0 ? $" ({string.Join(", ", flags)})" : $" ({flagStr})";
                    }
                    else
                    {
                        flagInfo = $" ({flagStr})";
                    }
                }

                result.Add(server + flagInfo);
            }

            return string.Join("; ", result);
        }

        /// <summary>
        /// Parse AnnounceFlags value
        /// </summary>
        private string ParseAnnounceFlags(string announceFlagsValue)
        {
            if (!int.TryParse(announceFlagsValue, out int flags))
                return announceFlagsValue;

            var flagList = new List<string>();
            if ((flags & 0x1) != 0) flagList.Add("Timeserv");
            if ((flags & 0x2) != 0) flagList.Add("Timeserv_Announce_Auto");
            if ((flags & 0x4) != 0) flagList.Add("Reliable_Timeserv");
            if ((flags & 0x8) != 0) flagList.Add("Timeserv_Announce_Yes");

            return flagList.Count > 0 
                ? $"{announceFlagsValue} ({string.Join(", ", flagList)})"
                : announceFlagsValue;
        }

        /// <summary>
        /// Get user information as structured sections
        /// </summary>
        public List<AnalysisSection> GetUserAnalysis()
        {
            var sections = new List<AnalysisSection>();

            // User Profiles (SOFTWARE hive)
            var profilesPath = @"Microsoft\Windows NT\CurrentVersion\ProfileList";
            var profilesKey = _parser.GetKey(profilesPath);
            if (profilesKey?.SubKeys != null && profilesKey.SubKeys.Count > 0)
            {
                var profilesSection = new AnalysisSection { Title = "📂 User Profiles" };
                foreach (var profile in profilesKey.SubKeys)
                {
                    var profilePath = profile.Values.FirstOrDefault(v => v.ValueName == "ProfileImagePath")?.ValueData?.ToString();
                    if (!string.IsNullOrEmpty(profilePath))
                    {
                        var username = Path.GetFileName(profilePath);
                        var profileRegPath = $"{profilesPath}\\{profile.KeyName}";
                        
                        // Get Last Logon Time (LocalProfileLoadTime)
                        var loadTimeHigh = profile.Values.FirstOrDefault(v => v.ValueName == "LocalProfileLoadTimeHigh");
                        var loadTimeLow = profile.Values.FirstOrDefault(v => v.ValueName == "LocalProfileLoadTimeLow");
                        string lastLogonStr = "N/A";
                        string lastLogonRaw = "LocalProfileLoadTimeHigh/Low not present";
                        if (loadTimeHigh != null && loadTimeLow != null)
                        {
                            var highVal = ParseDwordValue(loadTimeHigh.ValueData?.ToString());
                            var lowVal = ParseDwordValue(loadTimeLow.ValueData?.ToString());
                            var logonTime = FileTimeToDateTime(highVal, lowVal);
                            if (logonTime.HasValue)
                                lastLogonStr = logonTime.Value.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss");
                            lastLogonRaw = $"LocalProfileLoadTimeHigh = {loadTimeHigh.ValueData}, LocalProfileLoadTimeLow = {loadTimeLow.ValueData}";
                        }

                        // Get Last Logoff Time (LocalProfileUnloadTime)
                        var unloadTimeHigh = profile.Values.FirstOrDefault(v => v.ValueName == "LocalProfileUnloadTimeHigh");
                        var unloadTimeLow = profile.Values.FirstOrDefault(v => v.ValueName == "LocalProfileUnloadTimeLow");
                        string lastLogoffStr = "N/A";
                        string lastLogoffRaw = "LocalProfileUnloadTimeHigh/Low not present";
                        if (unloadTimeHigh != null && unloadTimeLow != null)
                        {
                            var highVal = ParseDwordValue(unloadTimeHigh.ValueData?.ToString());
                            var lowVal = ParseDwordValue(unloadTimeLow.ValueData?.ToString());
                            var logoffTime = FileTimeToDateTime(highVal, lowVal);
                            if (logoffTime.HasValue)
                                lastLogoffStr = logoffTime.Value.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss");
                            lastLogoffRaw = $"LocalProfileUnloadTimeHigh = {unloadTimeHigh.ValueData}, LocalProfileUnloadTimeLow = {unloadTimeLow.ValueData}";
                        }

                        var subItem = new AnalysisItem
                        {
                            Name = username,
                            IsSubSection = true,
                            RegistryPath = profileRegPath,
                            RegistryValue = $"SID = {profile.KeyName}\nProfileImagePath = {profilePath}",
                            SubItems = new List<AnalysisItem>
                            {
                                new AnalysisItem { Name = "SID", Value = profile.KeyName, RegistryPath = profileRegPath, RegistryValue = profile.KeyName },
                                new AnalysisItem { Name = "Path", Value = profilePath, RegistryPath = profileRegPath, RegistryValue = $"ProfileImagePath = {profilePath}" },
                                new AnalysisItem { Name = "Last Logon", Value = lastLogonStr, RegistryPath = profileRegPath, RegistryValue = lastLogonRaw },
                                new AnalysisItem { Name = "Last Logoff", Value = lastLogoffStr, RegistryPath = profileRegPath, RegistryValue = lastLogoffRaw }
                            }
                        };
                        profilesSection.Items.Add(subItem);
                    }
                }
                // Sort profiles alphabetically by username
                profilesSection.Items = profilesSection.Items.OrderBy(i => i.Name, StringComparer.OrdinalIgnoreCase).ToList();
                if (profilesSection.Items.Count > 0)
                    sections.Add(profilesSection);
            }

            if (sections.Count == 0)
            {
                var noDataSection = new AnalysisSection { Title = "ℹ️ Notice" };
                noDataSection.Items.Add(new AnalysisItem { Name = "Info", Value = "No user information found in this hive type", RegistryPath = "", RegistryValue = "User data not available in this hive" });
                noDataSection.Items.Add(new AnalysisItem { Name = "Tip", Value = "Load SAM, NTUSER.DAT, or SOFTWARE hive for user info", RegistryPath = "", RegistryValue = "Try loading a different registry hive" });
                sections.Add(noDataSection);
            }

            return sections;
        }

        /// <summary>
        /// Get all services as raw data for filtering in UI
        /// </summary>
        public List<ServiceInfo> GetAllServices()
        {
            var services = new List<ServiceInfo>();
            var servicesKey = _parser.GetKey($@"{_currentControlSet}\Services");
            
            if (servicesKey?.SubKeys != null)
            {
                foreach (var k in servicesKey.SubKeys)
                {
                    var startValue = k.Values.FirstOrDefault(v => v.ValueName == "Start")?.ValueData?.ToString() ?? "";
                    var imagePath = k.Values.FirstOrDefault(v => v.ValueName == "ImagePath")?.ValueData?.ToString() ?? "";
                    var displayName = k.Values.FirstOrDefault(v => v.ValueName == "DisplayName")?.ValueData?.ToString() ?? "";
                    var description = k.Values.FirstOrDefault(v => v.ValueName == "Description")?.ValueData?.ToString() ?? "";
                    var delayedAutoStart = k.Values.FirstOrDefault(v => v.ValueName == "DelayedAutostart")?.ValueData?.ToString() ?? "";
                    
                    bool isDelayedAuto = startValue == "2" && delayedAutoStart == "1";
                    
                    services.Add(new ServiceInfo
                    {
                        ServiceName = k.KeyName,
                        DisplayName = ServiceDisplayNameLookup.Resolve(k.KeyName, displayName),
                        Description = description,
                        StartType = startValue,
                        StartTypeName = isDelayedAuto ? "Automatic (Delayed)" : GetStartTypeName(startValue),
                        ImagePath = imagePath,
                        RegistryPath = $@"{_currentControlSet}\Services\{k.KeyName}",
                        IsDisabled = startValue == "4",
                        IsAutoStart = startValue == "2",
                        IsDelayedAutoStart = isDelayedAuto,
                        IsBoot = startValue == "0",
                        IsSystem = startValue == "1",
                        IsManual = startValue == "3"
                    });
                }
            }

            return services;
        }

        /// <summary>
        /// Get RDP (Remote Desktop) configuration as structured sections
        /// </summary>
        public List<AnalysisSection> GetRdpAnalysis()
        {
            var sections = new List<AnalysisSection>();
            string tsPath = $@"{_currentControlSet}\Control\Terminal Server";
            string rdpTcpPath = $@"{_currentControlSet}\Control\Terminal Server\WinStations\RDP-Tcp";
            string termServicePath = $@"{_currentControlSet}\Services\TermService";

            // Terminal Server Configuration (SYSTEM hive)
            var terminalServerKey = _parser.GetKey(tsPath);
            if (terminalServerKey != null)
            {
                var rdpSection = new AnalysisSection { Title = "🖥️ RDP Configuration" };

                // fDenyTSConnections
                var fDenyTSConnections = terminalServerKey.Values.FirstOrDefault(v => v.ValueName == "fDenyTSConnections")?.ValueData?.ToString() ?? "";
                rdpSection.Items.Add(new AnalysisItem
                {
                    Name = "fDenyTSConnections",
                    Value = fDenyTSConnections == "0" ? "0 (Enabled)" : $"{fDenyTSConnections} (Disabled)",
                    RegistryPath = tsPath,
                    RegistryValue = $"fDenyTSConnections = {fDenyTSConnections}"
                });

                // fSingleSessionPerUser
                var fSingleSessionPerUser = terminalServerKey.Values.FirstOrDefault(v => v.ValueName == "fSingleSessionPerUser")?.ValueData?.ToString() ?? "";
                rdpSection.Items.Add(new AnalysisItem
                {
                    Name = "fSingleSessionPerUser",
                    Value = fSingleSessionPerUser == "1" ? "1 (Yes)" : $"{fSingleSessionPerUser} (No)",
                    RegistryPath = tsPath,
                    RegistryValue = $"fSingleSessionPerUser = {fSingleSessionPerUser}"
                });

                // WinStations SelfSignedCertificate (binary thumbprint/identifier)
                string winStationsPathForCert = $@"{_currentControlSet}\Control\Terminal Server\WinStations";
                var selfSignedCertBytes = GetBinaryValue(winStationsPathForCert, "SelfSignedCertificate");
                var selfSignedCertDisplay = (selfSignedCertBytes != null && selfSignedCertBytes.Length > 0)
                    ? BitConverter.ToString(selfSignedCertBytes).Replace("-", "")
                    : "Not Set";
                rdpSection.Items.Add(new AnalysisItem
                {
                    Name = "SelfSignedCertificate",
                    Value = selfSignedCertDisplay,
                    RegistryPath = winStationsPathForCert,
                    RegistryValue = "SelfSignedCertificate"
                });

                // Check WinStations for RDP-Tcp settings
                var rdpTcpKey = _parser.GetKey(rdpTcpPath);
                if (rdpTcpKey != null)
                {
                    // PortNumber
                    var portNumber = rdpTcpKey.Values.FirstOrDefault(v => v.ValueName == "PortNumber")?.ValueData?.ToString() ?? "3389";
                    rdpSection.Items.Add(new AnalysisItem
                    {
                        Name = "PortNumber",
                        Value = portNumber,
                        RegistryPath = rdpTcpPath,
                        RegistryValue = $"PortNumber = {portNumber}"
                    });

                    // LanAdapter
                    var lanAdapter = rdpTcpKey.Values.FirstOrDefault(v => v.ValueName == "LanAdapter")?.ValueData?.ToString() ?? "";
                    rdpSection.Items.Add(new AnalysisItem
                    {
                        Name = "LanAdapter",
                        Value = lanAdapter,
                        RegistryPath = rdpTcpPath,
                        RegistryValue = $"LanAdapter = {lanAdapter}"
                    });

                    // SecurityLayer
                    var securityLayer = rdpTcpKey.Values.FirstOrDefault(v => v.ValueName == "SecurityLayer")?.ValueData?.ToString() ?? "";
                    var secLayerText = securityLayer switch
                    {
                        "0" => "0 (RDP Security Layer)",
                        "1" => "1 (Negotiate)",
                        "2" => "2 (TLS/SSL)",
                        _ => securityLayer
                    };
                    rdpSection.Items.Add(new AnalysisItem
                    {
                        Name = "SecurityLayer",
                        Value = secLayerText,
                        RegistryPath = rdpTcpPath,
                        RegistryValue = $"SecurityLayer = {securityLayer}"
                    });

                    // MinEncryptionLevel
                    var minEncryptionLevel = rdpTcpKey.Values.FirstOrDefault(v => v.ValueName == "MinEncryptionLevel")?.ValueData?.ToString() ?? "";
                    var encLevelText = minEncryptionLevel switch
                    {
                        "1" => "1 (Low)",
                        "2" => "2 (Client Compatible)",
                        "3" => "3 (High)",
                        "4" => "4 (FIPS Compliant)",
                        _ => minEncryptionLevel
                    };
                    rdpSection.Items.Add(new AnalysisItem
                    {
                        Name = "MinEncryptionLevel",
                        Value = encLevelText,
                        RegistryPath = rdpTcpPath,
                        RegistryValue = $"MinEncryptionLevel = {minEncryptionLevel}"
                    });

                    // UserAuthentication (NLA)
                    var userAuthentication = rdpTcpKey.Values.FirstOrDefault(v => v.ValueName == "UserAuthentication")?.ValueData?.ToString() ?? "";
                    rdpSection.Items.Add(new AnalysisItem
                    {
                        Name = "UserAuthentication",
                        Value = userAuthentication == "1" ? "1 (NLA Required)" : $"{userAuthentication} (NLA Not Required)",
                        RegistryPath = rdpTcpPath,
                        RegistryValue = $"UserAuthentication = {userAuthentication}"
                    });

                    // fAllowSecProtocolNegotiation
                    var fAllowSecProtocolNegotiation = rdpTcpKey.Values.FirstOrDefault(v => v.ValueName == "fAllowSecProtocolNegotiation")?.ValueData?.ToString() ?? "";
                    rdpSection.Items.Add(new AnalysisItem
                    {
                        Name = "fAllowSecProtocolNegotiation",
                        Value = fAllowSecProtocolNegotiation,
                        RegistryPath = rdpTcpPath,
                        RegistryValue = $"fAllowSecProtocolNegotiation = {fAllowSecProtocolNegotiation}"
                    });

                    // LoadableProtocol_Object - Citrix detection
                    var loadableProtocolObject = rdpTcpKey.Values.FirstOrDefault(v => v.ValueName == "LoadableProtocol_Object")?.ValueData?.ToString() ?? "";
                    var isCitrixProtocol = loadableProtocolObject.Contains("CitrixBackupRdpTcpLoadableProtocolObject", StringComparison.OrdinalIgnoreCase);
                    var protocolStatus = loadableProtocolObject switch
                    {
                        "18b726bb-6fe6-4fb9-9276-ed57ce7c7cb2" => "Default Windows",
                        "5828227c-20cf-4408-b73f-73ab70b8849f" => "Default Windows",
                        _ when isCitrixProtocol => "⚠️ Citrix (Modified)",
                        _ when string.IsNullOrEmpty(loadableProtocolObject) => "Not Set",
                        _ => loadableProtocolObject
                    };
                    rdpSection.Items.Add(new AnalysisItem
                    {
                        Name = "LoadableProtocol_Object",
                        Value = protocolStatus,
                        RegistryPath = rdpTcpPath,
                        RegistryValue = $"LoadableProtocol_Object = {loadableProtocolObject}"
                    });

                    // fQueryUserConfigFromLocalMachine
                    var fQueryUserConfig = rdpTcpKey.Values.FirstOrDefault(v => v.ValueName == "fQueryUserConfigFromLocalMachine")?.ValueData?.ToString() ?? "";
                    rdpSection.Items.Add(new AnalysisItem
                    {
                        Name = "fQueryUserConfigFromLocalMachine",
                        Value = fQueryUserConfig == "1" ? "1 (Enabled)" : $"{(string.IsNullOrEmpty(fQueryUserConfig) ? "Not Set" : fQueryUserConfig)} (Disabled)",
                        RegistryPath = rdpTcpPath,
                        RegistryValue = $"fQueryUserConfigFromLocalMachine = {fQueryUserConfig}"
                    });

                    // fEnableWinStation
                    var fEnableWinStation = rdpTcpKey.Values.FirstOrDefault(v => v.ValueName == "fEnableWinStation")?.ValueData?.ToString() ?? "";
                    rdpSection.Items.Add(new AnalysisItem
                    {
                        Name = "fEnableWinStation",
                        Value = fEnableWinStation == "1" ? "1 (Enabled)" : $"{(string.IsNullOrEmpty(fEnableWinStation) ? "Not Set" : fEnableWinStation)} (Disabled)",
                        RegistryPath = rdpTcpPath,
                        RegistryValue = $"fEnableWinStation = {fEnableWinStation}"
                    });

                    // MaxInstanceCount
                    var maxInstanceCount = rdpTcpKey.Values.FirstOrDefault(v => v.ValueName == "MaxInstanceCount")?.ValueData?.ToString() ?? "";
                    rdpSection.Items.Add(new AnalysisItem
                    {
                        Name = "MaxInstanceCount",
                        Value = string.IsNullOrEmpty(maxInstanceCount) ? "Not Set (Unlimited)" : maxInstanceCount,
                        RegistryPath = rdpTcpPath,
                        RegistryValue = $"MaxInstanceCount = {maxInstanceCount}"
                    });

                    // fInheritReconnectSame
                    var fInheritReconnectSame = rdpTcpKey.Values.FirstOrDefault(v => v.ValueName == "fInheritReconnectSame")?.ValueData?.ToString() ?? "";
                    rdpSection.Items.Add(new AnalysisItem
                    {
                        Name = "fInheritReconnectSame",
                        Value = fInheritReconnectSame == "1" ? "1 (Inherit from user)" : $"{(string.IsNullOrEmpty(fInheritReconnectSame) ? "Not Set" : fInheritReconnectSame)} (Use local)",
                        RegistryPath = rdpTcpPath,
                        RegistryValue = $"fInheritReconnectSame = {fInheritReconnectSame}"
                    });

                    // fReconnectSame
                    var fReconnectSame = rdpTcpKey.Values.FirstOrDefault(v => v.ValueName == "fReconnectSame")?.ValueData?.ToString() ?? "";
                    rdpSection.Items.Add(new AnalysisItem
                    {
                        Name = "fReconnectSame",
                        Value = fReconnectSame == "1" ? "1 (Reconnect to same session)" : $"{(string.IsNullOrEmpty(fReconnectSame) ? "Not Set" : fReconnectSame)} (Allow any)",
                        RegistryPath = rdpTcpPath,
                        RegistryValue = $"fReconnectSame = {fReconnectSame}"
                    });

                    // fDisableAutoReconnect (Policy)
                    string tsPolicyPath = @"SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services";
                    var tsPolicyKey = _parser.GetKey(tsPolicyPath);
                    if (tsPolicyKey != null)
                    {
                        var fDisableAutoReconnect = tsPolicyKey.Values.FirstOrDefault(v => v.ValueName == "fDisableAutoReconnect")?.ValueData?.ToString() ?? "";
                        rdpSection.Items.Add(new AnalysisItem
                        {
                            Name = "fDisableAutoReconnect (Policy)",
                            Value = fDisableAutoReconnect == "1" ? "1 (Auto-reconnect disabled)" : $"{(string.IsNullOrEmpty(fDisableAutoReconnect) ? "Not Set" : fDisableAutoReconnect)} (Auto-reconnect enabled)",
                            RegistryPath = tsPolicyPath,
                            RegistryValue = $"fDisableAutoReconnect = {fDisableAutoReconnect}"
                        });

                        // KeepAlive Policy settings
                        var keepAliveEnable = tsPolicyKey.Values.FirstOrDefault(v => v.ValueName == "KeepAliveEnable")?.ValueData?.ToString() ?? "";
                        var keepAliveInterval = tsPolicyKey.Values.FirstOrDefault(v => v.ValueName == "KeepAliveInterval")?.ValueData?.ToString() ?? "";
                        if (!string.IsNullOrEmpty(keepAliveEnable) || !string.IsNullOrEmpty(keepAliveInterval))
                        {
                            rdpSection.Items.Add(new AnalysisItem
                            {
                                Name = "KeepAliveEnable (Policy)",
                                Value = keepAliveEnable == "1" ? "1 (Enabled)" : $"{(string.IsNullOrEmpty(keepAliveEnable) ? "Not Set" : keepAliveEnable)}",
                                RegistryPath = tsPolicyPath,
                                RegistryValue = $"KeepAliveEnable = {keepAliveEnable}"
                            });
                            rdpSection.Items.Add(new AnalysisItem
                            {
                                Name = "KeepAliveInterval (Policy)",
                                Value = string.IsNullOrEmpty(keepAliveInterval) ? "Not Set" : $"{keepAliveInterval} minutes",
                                RegistryPath = tsPolicyPath,
                                RegistryValue = $"KeepAliveInterval = {keepAliveInterval}"
                            });
                        }
                    }
                }

                // TSServerDrainMode (Local and Policy)
                var localDrainMode = terminalServerKey.Values.FirstOrDefault(v => v.ValueName == "TSServerDrainMode")?.ValueData?.ToString() ?? "";
                var drainModeText = localDrainMode switch
                {
                    "0" => "0 (Allow all connections)",
                    "1" => "1 (Allow reconnections only until reboot)",
                    "2" => "2 (Allow reconnections only)",
                    _ when string.IsNullOrEmpty(localDrainMode) => "Not Set (Default: Allow all)",
                    _ => localDrainMode
                };
                rdpSection.Items.Add(new AnalysisItem
                {
                    Name = "TSServerDrainMode",
                    Value = drainModeText,
                    RegistryPath = tsPath,
                    RegistryValue = $"TSServerDrainMode = {localDrainMode}"
                });

                sections.Add(rdpSection);

                // Session Limits Section
                var rdpTcpKeyForLimits = _parser.GetKey(rdpTcpPath);
                if (rdpTcpKeyForLimits != null)
                {
                    var limitsSection = new AnalysisSection { Title = "⏱️ Session Limits" };

                    // KeepAliveTimeout
                    var keepAliveTimeout = rdpTcpKeyForLimits.Values.FirstOrDefault(v => v.ValueName == "KeepAliveTimeout")?.ValueData?.ToString() ?? "";
                    limitsSection.Items.Add(new AnalysisItem
                    {
                        Name = "KeepAliveTimeout",
                        Value = string.IsNullOrEmpty(keepAliveTimeout) ? "Not Set" : $"{keepAliveTimeout} ms",
                        RegistryPath = rdpTcpPath,
                        RegistryValue = $"KeepAliveTimeout = {keepAliveTimeout}"
                    });

                    // MaxConnectionTime
                    var maxConnectionTime = rdpTcpKeyForLimits.Values.FirstOrDefault(v => v.ValueName == "MaxConnectionTime")?.ValueData?.ToString() ?? "";
                    var maxConnTimeDisplay = string.IsNullOrEmpty(maxConnectionTime) || maxConnectionTime == "0" 
                        ? "Not Set (No limit)" 
                        : FormatMilliseconds(maxConnectionTime);
                    limitsSection.Items.Add(new AnalysisItem
                    {
                        Name = "MaxConnectionTime",
                        Value = maxConnTimeDisplay,
                        RegistryPath = rdpTcpPath,
                        RegistryValue = $"MaxConnectionTime = {maxConnectionTime}"
                    });

                    // fInheritMaxSessionTime
                    var fInheritMaxSessionTime = rdpTcpKeyForLimits.Values.FirstOrDefault(v => v.ValueName == "fInheritMaxSessionTime")?.ValueData?.ToString() ?? "";
                    limitsSection.Items.Add(new AnalysisItem
                    {
                        Name = "fInheritMaxSessionTime",
                        Value = fInheritMaxSessionTime == "1" ? "1 (Inherit from user)" : $"{(string.IsNullOrEmpty(fInheritMaxSessionTime) ? "Not Set" : fInheritMaxSessionTime)} (Use local)",
                        RegistryPath = rdpTcpPath,
                        RegistryValue = $"fInheritMaxSessionTime = {fInheritMaxSessionTime}"
                    });

                    // MaxDisconnectionTime
                    var maxDisconnectionTime = rdpTcpKeyForLimits.Values.FirstOrDefault(v => v.ValueName == "MaxDisconnectionTime")?.ValueData?.ToString() ?? "";
                    var maxDisconnTimeDisplay = string.IsNullOrEmpty(maxDisconnectionTime) || maxDisconnectionTime == "0" 
                        ? "Not Set (No limit)" 
                        : FormatMilliseconds(maxDisconnectionTime);
                    limitsSection.Items.Add(new AnalysisItem
                    {
                        Name = "MaxDisconnectionTime",
                        Value = maxDisconnTimeDisplay,
                        RegistryPath = rdpTcpPath,
                        RegistryValue = $"MaxDisconnectionTime = {maxDisconnectionTime}"
                    });

                    // fInheritMaxDisconnectionTime
                    var fInheritMaxDisconnectionTime = rdpTcpKeyForLimits.Values.FirstOrDefault(v => v.ValueName == "fInheritMaxDisconnectionTime")?.ValueData?.ToString() ?? "";
                    limitsSection.Items.Add(new AnalysisItem
                    {
                        Name = "fInheritMaxDisconnectionTime",
                        Value = fInheritMaxDisconnectionTime == "1" ? "1 (Inherit from user)" : $"{(string.IsNullOrEmpty(fInheritMaxDisconnectionTime) ? "Not Set" : fInheritMaxDisconnectionTime)} (Use local)",
                        RegistryPath = rdpTcpPath,
                        RegistryValue = $"fInheritMaxDisconnectionTime = {fInheritMaxDisconnectionTime}"
                    });

                    // MaxIdleTime
                    var maxIdleTime = rdpTcpKeyForLimits.Values.FirstOrDefault(v => v.ValueName == "MaxIdleTime")?.ValueData?.ToString() ?? "";
                    var maxIdleTimeDisplay = string.IsNullOrEmpty(maxIdleTime) || maxIdleTime == "0" 
                        ? "Not Set (No limit)" 
                        : FormatMilliseconds(maxIdleTime);
                    limitsSection.Items.Add(new AnalysisItem
                    {
                        Name = "MaxIdleTime",
                        Value = maxIdleTimeDisplay,
                        RegistryPath = rdpTcpPath,
                        RegistryValue = $"MaxIdleTime = {maxIdleTime}"
                    });

                    // fInheritMaxIdleTime
                    var fInheritMaxIdleTime = rdpTcpKeyForLimits.Values.FirstOrDefault(v => v.ValueName == "fInheritMaxIdleTime")?.ValueData?.ToString() ?? "";
                    limitsSection.Items.Add(new AnalysisItem
                    {
                        Name = "fInheritMaxIdleTime",
                        Value = fInheritMaxIdleTime == "1" ? "1 (Inherit from user)" : $"{(string.IsNullOrEmpty(fInheritMaxIdleTime) ? "Not Set" : fInheritMaxIdleTime)} (Use local)",
                        RegistryPath = rdpTcpPath,
                        RegistryValue = $"fInheritMaxIdleTime = {fInheritMaxIdleTime}"
                    });

                    sections.Add(limitsSection);
                }
            }

            // Citrix Detection Section
            string winStationsPath = $@"{_currentControlSet}\Control\Terminal Server\WinStations";
            var winStationsKey = _parser.GetKey(winStationsPath);
            bool hasCitrixICA = false;
            var icaStations = new List<string>();

            if (winStationsKey?.SubKeys != null)
            {
                foreach (var subKey in winStationsKey.SubKeys)
                {
                    if (subKey.KeyName.Contains("ICA", StringComparison.OrdinalIgnoreCase))
                    {
                        hasCitrixICA = true;
                        icaStations.Add(subKey.KeyName);
                    }
                }
            }

            // Check LoadableProtocol_Object for Citrix
            var rdpTcpKeyForCitrix = _parser.GetKey(rdpTcpPath);
            var citrixProtocolDetected = false;
            var loadableProtocolValue = "";
            if (rdpTcpKeyForCitrix != null)
            {
                loadableProtocolValue = rdpTcpKeyForCitrix.Values.FirstOrDefault(v => v.ValueName == "LoadableProtocol_Object")?.ValueData?.ToString() ?? "";
                citrixProtocolDetected = loadableProtocolValue.Contains("CitrixBackupRdpTcpLoadableProtocolObject", StringComparison.OrdinalIgnoreCase);
            }

            // Always show Citrix Detection section
            {
                var citrixSection = new AnalysisSection { Title = "🍊 Citrix Detection" };

                // Overall status
                if (hasCitrixICA || citrixProtocolDetected)
                {
                    citrixSection.Items.Add(new AnalysisItem
                    {
                        Name = "Citrix Status",
                        Value = "⚠️ CITRIX DETECTED - VM is using Citrix for RDP",
                        RegistryPath = winStationsPath,
                        RegistryValue = "Citrix components found in WinStations"
                    });
                }
                else
                {
                    citrixSection.Items.Add(new AnalysisItem
                    {
                        Name = "Citrix Status",
                        Value = "✅ Not Detected - Standard Windows RDP",
                        RegistryPath = winStationsPath,
                        RegistryValue = "No Citrix components found"
                    });
                }

                // Protocol Override status
                if (citrixProtocolDetected)
                {
                    citrixSection.Items.Add(new AnalysisItem
                    {
                        Name = "Protocol Override",
                        Value = "⚠️ CitrixBackupRdpTcpLoadableProtocolObject",
                        RegistryPath = rdpTcpPath,
                        RegistryValue = $"LoadableProtocol_Object = {loadableProtocolValue}"
                    });
                }
                else
                {
                    citrixSection.Items.Add(new AnalysisItem
                    {
                        Name = "Protocol Override",
                        Value = string.IsNullOrEmpty(loadableProtocolValue) ? "Not Set (Default)" : "Standard Windows",
                        RegistryPath = rdpTcpPath,
                        RegistryValue = $"LoadableProtocol_Object = {(string.IsNullOrEmpty(loadableProtocolValue) ? "(empty)" : loadableProtocolValue)}"
                    });
                }

                // ICA WinStations status
                if (hasCitrixICA)
                {
                    citrixSection.Items.Add(new AnalysisItem
                    {
                        Name = "ICA WinStations",
                        Value = $"⚠️ Found: {string.Join(", ", icaStations)}",
                        RegistryPath = winStationsPath,
                        RegistryValue = $"ICA Subkeys: {string.Join(", ", icaStations)}"
                    });

                    foreach (var icaStation in icaStations.Take(5)) // Limit to first 5
                    {
                        var icaPath = $@"{winStationsPath}\{icaStation}";
                        var icaKey = _parser.GetKey(icaPath);
                        if (icaKey != null)
                        {
                            var comment = icaKey.Values.FirstOrDefault(v => v.ValueName == "Comment")?.ValueData?.ToString() ?? "";
                            if (!string.IsNullOrEmpty(comment))
                            {
                                citrixSection.Items.Add(new AnalysisItem
                                {
                                    Name = icaStation,
                                    Value = comment,
                                    RegistryPath = icaPath,
                                    RegistryValue = $"Comment = {comment}"
                                });
                            }
                        }
                    }
                }
                else
                {
                    citrixSection.Items.Add(new AnalysisItem
                    {
                        Name = "ICA WinStations",
                        Value = "✅ None Found",
                        RegistryPath = winStationsPath,
                        RegistryValue = "No ICA subkeys in WinStations"
                    });
                }

                sections.Insert(0, citrixSection); // Add at the top for visibility
            }

            // Terminal Service Settings
            var rdpServiceKey = _parser.GetKey(termServicePath);
            if (rdpServiceKey != null)
            {
                var serviceSection = new AnalysisSection { Title = "⚙️ Terminal Service" };

                // TermServiceStart (Start value)
                var startValue = rdpServiceKey.Values.FirstOrDefault(v => v.ValueName == "Start")?.ValueData?.ToString() ?? "";
                var startTypeText = startValue switch
                {
                    "0" => "0 (Boot)",
                    "1" => "1 (System)",
                    "2" => "2 (Automatic)",
                    "3" => "3 (Manual)",
                    "4" => "4 (Disabled)",
                    _ => startValue
                };
                serviceSection.Items.Add(new AnalysisItem
                {
                    Name = "TermServiceStart",
                    Value = startTypeText,
                    RegistryPath = termServicePath,
                    RegistryValue = $"Start = {startValue}"
                });

                // TermServiceAccount (ObjectName)
                var objectName = rdpServiceKey.Values.FirstOrDefault(v => v.ValueName == "ObjectName")?.ValueData?.ToString() ?? "";
                serviceSection.Items.Add(new AnalysisItem
                {
                    Name = "TermServiceAccount",
                    Value = objectName,
                    RegistryPath = termServicePath,
                    RegistryValue = $"ObjectName = {objectName}"
                });

                // ImagePath
                var imagePath = rdpServiceKey.Values.FirstOrDefault(v => v.ValueName == "ImagePath")?.ValueData?.ToString() ?? "";
                if (!string.IsNullOrEmpty(imagePath))
                {
                    serviceSection.Items.Add(new AnalysisItem
                    {
                        Name = "ImagePath",
                        Value = imagePath,
                        RegistryPath = termServicePath,
                        RegistryValue = $"ImagePath = {imagePath}"
                    });
                }

                sections.Add(serviceSection);
            }

            // Terminal Services Licensing
            string licensingPath = $@"{_currentControlSet}\Control\Terminal Server\Licensing Core";
            var licensingKey = _parser.GetKey(licensingPath);
            if (licensingKey != null)
            {
                var licenseSection = new AnalysisSection { Title = "📜 RDP Licensing" };

                // LicensingMode
                var licensingMode = licensingKey.Values.FirstOrDefault(v => v.ValueName == "LicensingMode")?.ValueData?.ToString() ?? "";
                var licensingModeText = licensingMode switch
                {
                    "2" => "2 (Per Device)",
                    "4" => "4 (Per User)",
                    "5" => "5 (Not Configured)",
                    _ => licensingMode
                };
                licenseSection.Items.Add(new AnalysisItem
                {
                    Name = "LicensingMode",
                    Value = licensingModeText,
                    RegistryPath = licensingPath,
                    RegistryValue = $"LicensingMode = {licensingMode}"
                });

                // EnableConcurrentSessions
                var enableConcurrentSessions = licensingKey.Values.FirstOrDefault(v => v.ValueName == "EnableConcurrentSessions")?.ValueData?.ToString() ?? "";
                licenseSection.Items.Add(new AnalysisItem
                {
                    Name = "EnableConcurrentSessions",
                    Value = enableConcurrentSessions == "1" ? "1 (Enabled)" : $"{enableConcurrentSessions} (Disabled)",
                    RegistryPath = licensingPath,
                    RegistryValue = $"EnableConcurrentSessions = {enableConcurrentSessions}"
                });

                sections.Add(licenseSection);
            }

            // RDS (Remote Desktop Services) Licensing Section
            var rdsSection = new AnalysisSection { Title = "🏢 Remote Desktop Services (RDS)" };
            
            // SpecifiedLicenseServers from TermService\Parameters
            string termServiceParamsPath = $@"{_currentControlSet}\Services\TermService\Parameters";
            var termServiceParamsKey = _parser.GetKey(termServiceParamsPath);
            if (termServiceParamsKey != null)
            {
                // SpecifiedLicenseServers can be a multi-string value
                var specifiedLicenseServers = termServiceParamsKey.Values.FirstOrDefault(v => v.ValueName == "SpecifiedLicenseServers");
                if (specifiedLicenseServers != null)
                {
                    var serversValue = specifiedLicenseServers.ValueData?.ToString() ?? "";
                    // Multi-string values may be separated by various delimiters
                    var displayValue = string.IsNullOrEmpty(serversValue) ? "Not Configured" : serversValue.Replace("\0", ", ").TrimEnd(',', ' ');
                    rdsSection.Items.Add(new AnalysisItem
                    {
                        Name = "SpecifiedLicenseServers",
                        Value = displayValue,
                        RegistryPath = termServiceParamsPath,
                        RegistryValue = $"SpecifiedLicenseServers = {serversValue}"
                    });
                }
                else
                {
                    rdsSection.Items.Add(new AnalysisItem
                    {
                        Name = "SpecifiedLicenseServers",
                        Value = "Not Configured",
                        RegistryPath = termServiceParamsPath,
                        RegistryValue = "SpecifiedLicenseServers not set"
                    });
                }
            }
            else
            {
                rdsSection.Items.Add(new AnalysisItem
                {
                    Name = "SpecifiedLicenseServers",
                    Value = "Not Configured",
                    RegistryPath = termServiceParamsPath,
                    RegistryValue = "TermService\\Parameters key not found"
                });
            }

            // LicensingMode from RCM\Licensing Core
            string rcmLicensingPath = $@"{_currentControlSet}\Control\Terminal Server\RCM\Licensing Core";
            var rcmLicensingKey = _parser.GetKey(rcmLicensingPath);
            if (rcmLicensingKey != null)
            {
                var rcmLicensingMode = rcmLicensingKey.Values.FirstOrDefault(v => v.ValueName == "LicensingMode")?.ValueData?.ToString() ?? "";
                var rcmLicensingModeText = rcmLicensingMode switch
                {
                    "2" => "2 (Per Device CAL)",
                    "4" => "4 (Per User CAL)",
                    "5" => "5 (Not Configured)",
                    _ when string.IsNullOrEmpty(rcmLicensingMode) => "Not Configured",
                    _ => rcmLicensingMode
                };
                rdsSection.Items.Add(new AnalysisItem
                {
                    Name = "LicensingMode (RCM)",
                    Value = rcmLicensingModeText,
                    RegistryPath = rcmLicensingPath,
                    RegistryValue = $"LicensingMode = {rcmLicensingMode}"
                });
            }
            else
            {
                rdsSection.Items.Add(new AnalysisItem
                {
                    Name = "LicensingMode (RCM)",
                    Value = "Not Configured",
                    RegistryPath = rcmLicensingPath,
                    RegistryValue = "RCM\\Licensing Core key not found"
                });
            }

            // Add additional RDS relevant settings if available
            // Check for RDS policy settings
            string rdsPolicyPath = $@"{_currentControlSet}\Control\Terminal Server\RCM";
            var rcmKey = _parser.GetKey(rdsPolicyPath);
            if (rcmKey != null)
            {
                var gracePeriodRemaining = rcmKey.Values.FirstOrDefault(v => v.ValueName == "GracePeriod")?.ValueData?.ToString() ?? "";
                if (!string.IsNullOrEmpty(gracePeriodRemaining))
                {
                    rdsSection.Items.Add(new AnalysisItem
                    {
                        Name = "GracePeriod",
                        Value = gracePeriodRemaining,
                        RegistryPath = rdsPolicyPath,
                        RegistryValue = $"GracePeriod = {gracePeriodRemaining}"
                    });
                }
            }

            sections.Add(rdsSection);

            if (sections.Count == 0)
            {
                var noDataSection = new AnalysisSection { Title = "ℹ️ Notice" };
                noDataSection.Items.Add(new AnalysisItem { Name = "Info", Value = "No RDP configuration found in this hive" });
                noDataSection.Items.Add(new AnalysisItem { Name = "Tip", Value = "Load SYSTEM hive for RDP configuration" });
                sections.Add(noDataSection);
            }

            return sections;
        }

        /// <summary>
        /// Get network information as structured sections
        /// </summary>
        public List<AnalysisSection> GetNetworkAnalysis()
        {
            var sections = new List<AnalysisSection>();

            // Build a lookup from NetCfgInstanceId to adapter info from Class key
            var adapterClassInfo = new Dictionary<string, (string DriverDesc, string DriverVersion, string ProviderName, string ClassPath, string DeviceInstanceId, string InstallTimestamp)>(StringComparer.OrdinalIgnoreCase);
            var classKey = _parser.GetKey($@"{_currentControlSet}\Control\Class\{{4d36e972-e325-11ce-bfc1-08002be10318}}");
            if (classKey?.SubKeys != null)
            {
                foreach (var adapter in classKey.SubKeys)
                {
                    var netCfgInstanceId = adapter.Values.FirstOrDefault(v => v.ValueName == "NetCfgInstanceId")?.ValueData?.ToString();
                    if (!string.IsNullOrEmpty(netCfgInstanceId))
                    {
                        var driverDesc = adapter.Values.FirstOrDefault(v => v.ValueName == "DriverDesc")?.ValueData?.ToString() ?? "";
                        var driverVersion = adapter.Values.FirstOrDefault(v => v.ValueName == "DriverVersion")?.ValueData?.ToString() ?? "";
                        var providerName = adapter.Values.FirstOrDefault(v => v.ValueName == "ProviderName")?.ValueData?.ToString() ?? "";
                        var deviceInstanceId = adapter.Values.FirstOrDefault(v => v.ValueName == "DeviceInstanceID")?.ValueData?.ToString() ?? "";
                        var classPath = $@"{_currentControlSet}\Control\Class\{{4d36e972-e325-11ce-bfc1-08002be10318}}\{adapter.KeyName}";
                        
                        // Get NetworkInterfaceInstallTimestamp (FILETIME format - 64-bit value)
                        string installTimestamp = "";
                        var timestampValue = adapter.Values.FirstOrDefault(v => v.ValueName == "NetworkInterfaceInstallTimestamp");
                        if (timestampValue?.ValueData != null)
                        {
                            try
                            {
                                var dataStr = timestampValue.ValueData.ToString() ?? "";
                                
                                // Try parsing as hex string (common format for binary data)
                                // The value might be in format like "01 D9 8A 7E 3B 84 DA 01" or "01D98A7E3B84DA01"
                                var hexStr = dataStr.Replace(" ", "").Replace("-", "");
                                if (hexStr.Length == 16 && System.Text.RegularExpressions.Regex.IsMatch(hexStr, @"^[0-9A-Fa-f]+$"))
                                {
                                    // Parse as hex and convert to bytes (little-endian)
                                    byte[] bytes = new byte[8];
                                    for (int i = 0; i < 8; i++)
                                    {
                                        bytes[i] = Convert.ToByte(hexStr.Substring(i * 2, 2), 16);
                                    }
                                    long fileTime = BitConverter.ToInt64(bytes, 0);
                                    if (fileTime > 0 && fileTime < long.MaxValue)
                                    {
                                        var dateTime = DateTime.FromFileTimeUtc(fileTime);
                                        installTimestamp = dateTime.ToString("yyyy-MM-dd HH:mm:ss") + " UTC";
                                    }
                                }
                                // Try parsing as decimal number
                                else if (long.TryParse(dataStr, out long ft) && ft > 0 && ft < long.MaxValue)
                                {
                                    var dateTime = DateTime.FromFileTimeUtc(ft);
                                    installTimestamp = dateTime.ToString("yyyy-MM-dd HH:mm:ss") + " UTC";
                                }
                            }
                            catch { /* Ignore conversion errors */ }
                        }
                        
                        adapterClassInfo[netCfgInstanceId] = (driverDesc, driverVersion, providerName, classPath, deviceInstanceId, installTimestamp);
                    }
                }
            }

            // Build a lookup for device enabled/disabled state from Enum key
            // ConfigFlags: 0 = Enabled, 1 (bit 0 set) = Disabled
            var deviceEnabledState = new Dictionary<string, (bool IsEnabled, string EnumPath)>(StringComparer.OrdinalIgnoreCase);
            var enumKey = _parser.GetKey($@"{_currentControlSet}\Enum");
            if (enumKey?.SubKeys != null)
            {
                foreach (var busType in enumKey.SubKeys) // PCI, USB, ROOT, etc.
                {
                    if (busType.SubKeys != null)
                    {
                        foreach (var device in busType.SubKeys)
                        {
                            if (device.SubKeys != null)
                            {
                                foreach (var instance in device.SubKeys)
                                {
                                    var configFlags = instance.Values.FirstOrDefault(v => v.ValueName == "ConfigFlags")?.ValueData?.ToString();
                                    var deviceInstancePath = $@"{busType.KeyName}\{device.KeyName}\{instance.KeyName}";
                                    var enumPath = $@"{_currentControlSet}\Enum\{deviceInstancePath}";
                                    
                                    // ConfigFlags bit 0 = CONFIGFLAG_DISABLED
                                    bool isEnabled = true;
                                    if (!string.IsNullOrEmpty(configFlags) && int.TryParse(configFlags, out int flags))
                                    {
                                        isEnabled = (flags & 0x01) == 0; // Bit 0 = disabled
                                    }
                                    
                                    deviceEnabledState[deviceInstancePath] = (isEnabled, enumPath);
                                }
                            }
                        }
                    }
                }
            }

            // Network Interfaces (SYSTEM hive)
            var interfacesKey = _parser.GetKey($@"{_currentControlSet}\Services\Tcpip\Parameters\Interfaces");
            if (interfacesKey?.SubKeys != null && interfacesKey.SubKeys.Count > 0)
            {
                var ifaceSection = new AnalysisSection { Title = "🔌 Network Interfaces" };
                foreach (var iface in interfacesKey.SubKeys)
                {
                    var dhcpEnabled = iface.Values.FirstOrDefault(v => v.ValueName == "EnableDHCP")?.ValueData?.ToString();
                    var ipAddress = iface.Values.FirstOrDefault(v => v.ValueName == "DhcpIPAddress")?.ValueData?.ToString()
                        ?? iface.Values.FirstOrDefault(v => v.ValueName == "IPAddress")?.ValueData?.ToString();
                    var subnetMask = iface.Values.FirstOrDefault(v => v.ValueName == "DhcpSubnetMask")?.ValueData?.ToString()
                        ?? iface.Values.FirstOrDefault(v => v.ValueName == "SubnetMask")?.ValueData?.ToString();
                    var gateway = iface.Values.FirstOrDefault(v => v.ValueName == "DhcpDefaultGateway")?.ValueData?.ToString()
                        ?? iface.Values.FirstOrDefault(v => v.ValueName == "DefaultGateway")?.ValueData?.ToString();
                    var dns = iface.Values.FirstOrDefault(v => v.ValueName == "DhcpNameServer")?.ValueData?.ToString()
                        ?? iface.Values.FirstOrDefault(v => v.ValueName == "NameServer")?.ValueData?.ToString();
                    var domain = iface.Values.FirstOrDefault(v => v.ValueName == "DhcpDomain")?.ValueData?.ToString()
                        ?? iface.Values.FirstOrDefault(v => v.ValueName == "Domain")?.ValueData?.ToString();
                    var dhcpServer = iface.Values.FirstOrDefault(v => v.ValueName == "DhcpServer")?.ValueData?.ToString();

                    var registryPath = $@"{_currentControlSet}\Services\Tcpip\Parameters\Interfaces\{iface.KeyName}";

                    // Try to get adapter class info
                    string adapterName = iface.KeyName;
                    string driverDesc = "";
                    string driverVersion = "";
                    string providerName = "";
                    string classPath = "";
                    string deviceInstanceId = "";
                    string installTimestamp = "";
                    bool isEnabled = true;
                    string enumPath = "";
                    
                    if (adapterClassInfo.TryGetValue(iface.KeyName, out var classInfo))
                    {
                        driverDesc = classInfo.DriverDesc;
                        driverVersion = classInfo.DriverVersion;
                        providerName = classInfo.ProviderName;
                        classPath = classInfo.ClassPath;
                        deviceInstanceId = classInfo.DeviceInstanceId;
                        installTimestamp = classInfo.InstallTimestamp;
                        if (!string.IsNullOrEmpty(driverDesc))
                            adapterName = driverDesc;
                        
                        // Look up enabled state from Enum key
                        if (!string.IsNullOrEmpty(deviceInstanceId) && deviceEnabledState.TryGetValue(deviceInstanceId, out var enabledInfo))
                        {
                            isEnabled = enabledInfo.IsEnabled;
                            enumPath = enabledInfo.EnumPath;
                        }
                    }

                    if (!string.IsNullOrEmpty(ipAddress) || !string.IsNullOrEmpty(gateway))
                    {
                        var subItems = new List<AnalysisItem>();
                        
                        // Add adapter/driver info first if available (from Class key)
                        if (!string.IsNullOrEmpty(driverDesc))
                            subItems.Add(new AnalysisItem { Name = "Adapter Name", Value = driverDesc, RegistryPath = classPath, RegistryValue = "DriverDesc" });
                        
                        // Add enabled/disabled status (from Enum key ConfigFlags)
                        var statusValue = isEnabled ? "✅ Enabled" : "❌ Disabled";
                        var statusPath = !string.IsNullOrEmpty(enumPath) ? enumPath : classPath;
                        subItems.Add(new AnalysisItem { Name = "Status", Value = statusValue, RegistryPath = statusPath, RegistryValue = "ConfigFlags" });
                        
                        // Add install timestamp if available
                        if (!string.IsNullOrEmpty(installTimestamp))
                            subItems.Add(new AnalysisItem { Name = "Install Date", Value = installTimestamp, RegistryPath = classPath, RegistryValue = "NetworkInterfaceInstallTimestamp" });
                        
                        if (!string.IsNullOrEmpty(providerName))
                            subItems.Add(new AnalysisItem { Name = "Provider", Value = providerName, RegistryPath = classPath, RegistryValue = "ProviderName" });
                        if (!string.IsNullOrEmpty(driverVersion))
                            subItems.Add(new AnalysisItem { Name = "Driver Version", Value = driverVersion, RegistryPath = classPath, RegistryValue = "DriverVersion" });
                        
                        // Add clear separator for TCP/IP section
                        subItems.Add(new AnalysisItem { Name = "══════════════════", Value = "══════════════════════════" });
                        subItems.Add(new AnalysisItem { Name = "► TCP/IP Configuration", Value = "", RegistryPath = registryPath, RegistryValue = "" });
                        
                        // Add TCP/IP config (from Interfaces key)
                        subItems.Add(new AnalysisItem { Name = "DHCP", Value = dhcpEnabled == "1" ? "Yes" : "No", RegistryPath = registryPath, RegistryValue = "EnableDHCP" });
                        subItems.Add(new AnalysisItem { Name = "IP Address", Value = ipAddress ?? "", RegistryPath = registryPath, RegistryValue = dhcpEnabled == "1" ? "DhcpIPAddress" : "IPAddress" });
                        subItems.Add(new AnalysisItem { Name = "Subnet", Value = subnetMask ?? "", RegistryPath = registryPath, RegistryValue = dhcpEnabled == "1" ? "DhcpSubnetMask" : "SubnetMask" });
                        subItems.Add(new AnalysisItem { Name = "Gateway", Value = gateway ?? "", RegistryPath = registryPath, RegistryValue = dhcpEnabled == "1" ? "DhcpDefaultGateway" : "DefaultGateway" });
                        subItems.Add(new AnalysisItem { Name = "DNS", Value = dns ?? "", RegistryPath = registryPath, RegistryValue = dhcpEnabled == "1" ? "DhcpNameServer" : "NameServer" });
                        if (!string.IsNullOrEmpty(dhcpServer))
                            subItems.Add(new AnalysisItem { Name = "DHCP Server", Value = dhcpServer, RegistryPath = registryPath, RegistryValue = "DhcpServer" });
                        if (!string.IsNullOrEmpty(domain))
                            subItems.Add(new AnalysisItem { Name = "Domain", Value = domain, RegistryPath = registryPath, RegistryValue = dhcpEnabled == "1" ? "DhcpDomain" : "Domain" });

                        // DNS Registered Adapters data for this interface GUID
                        var dnsRegAdapterPath = $@"{_currentControlSet}\Services\Tcpip\Parameters\DNSRegisteredAdapters\{iface.KeyName}";
                        var dnsRegKey = _parser.GetKey(dnsRegAdapterPath);
                        if (dnsRegKey != null)
                        {
                            var primaryDomainNameVal = dnsRegKey.Values.FirstOrDefault(v => v.ValueName == "PrimaryDomainName")?.ValueData?.ToString() ?? "";
                            if (!string.IsNullOrEmpty(primaryDomainNameVal))
                            {
                                subItems.Add(new AnalysisItem
                                {
                                    Name = "PrimaryDomainName",
                                    Value = primaryDomainNameVal,
                                    RegistryPath = dnsRegAdapterPath,
                                    RegistryValue = $"PrimaryDomainName = {primaryDomainNameVal}"
                                });
                            }

                            var staleAdapterVal = dnsRegKey.Values.FirstOrDefault(v => v.ValueName == "StaleAdapter")?.ValueData?.ToString() ?? "";
                            if (!string.IsNullOrEmpty(staleAdapterVal))
                            {
                                var staleText = staleAdapterVal == "1" ? "Yes (Ghosted)" : "No";
                                subItems.Add(new AnalysisItem
                                {
                                    Name = "StaleAdapter",
                                    Value = staleText,
                                    RegistryPath = dnsRegAdapterPath,
                                    RegistryValue = $"StaleAdapter = {staleAdapterVal}"
                                });
                            }
                        }

                        // Build detailed registry value for bottom pane
                        var details = new List<string>();
                        details.Add($"Interface GUID: {iface.KeyName}");
                        if (!string.IsNullOrEmpty(classPath)) details.Add($"Class Path: {classPath}");

                        // Use full friendly name if available
                        var displayName = !string.IsNullOrEmpty(driverDesc) 
                            ? driverDesc
                            : iface.KeyName;

                        ifaceSection.Items.Add(new AnalysisItem
                        {
                            Name = displayName,
                            IsSubSection = true,
                            SubItems = subItems,
                            RegistryPath = registryPath,
                            RegistryValue = string.Join(" | ", details)
                        });
                    }
                }

                if (ifaceSection.Items.Count > 0)
                    sections.Add(ifaceSection);
            }

            // DNS Registered Adapters (PrimaryDomainName and Ghosted NICs)
            string dnsRegBasePath = $@"{_currentControlSet}\Services\Tcpip\Parameters\DNSRegisteredAdapters";
            var primaryDomainKey = _parser.GetKey($"{dnsRegBasePath}\\PrimaryDomainName");
            if (primaryDomainKey != null)
            {
                var dnsRegSection = new AnalysisSection { Title = "🧭 DNS Registered Adapters" };

                // AdapterGuidList (multi-string) under PrimaryDomainName
                var adapterGuidListVal = primaryDomainKey.Values.FirstOrDefault(v => v.ValueName.Equals("AdapterGuidList", StringComparison.OrdinalIgnoreCase));
                if (adapterGuidListVal != null)
                {
                    var val = adapterGuidListVal.ValueData?.ToString() ?? string.Empty;
                    var display = string.IsNullOrWhiteSpace(val) ? "(empty)" : val.Replace("\0", ", ").TrimEnd(',', ' ');
                    dnsRegSection.Items.Add(new AnalysisItem
                    {
                        Name = "PrimaryDomainName.AdapterGuidList",
                        Value = display,
                        RegistryPath = $"{dnsRegBasePath}\\PrimaryDomainName",
                        RegistryValue = $"AdapterGuidList = {val}"
                    });
                }
                else if (primaryDomainKey.Values.Count > 0)
                {
                    // Fallback: list any values present under PrimaryDomainName
                    foreach (var val in primaryDomainKey.Values)
                    {
                        dnsRegSection.Items.Add(new AnalysisItem
                        {
                            Name = $"PrimaryDomainName.{val.ValueName}",
                            Value = val.ValueData?.ToString() ?? "",
                            RegistryPath = $"{dnsRegBasePath}\\PrimaryDomainName",
                            RegistryValue = $"{val.ValueName} = {val.ValueData}"
                        });
                    }
                }

                // StaleAdapter (ghosted NIC GUIDs) under PrimaryDomainName
                var staleKey = _parser.GetKey($"{dnsRegBasePath}\\PrimaryDomainName\\StaleAdapter");
                if (staleKey != null)
                {
                    if (staleKey.Values.Count > 0)
                    {
                        foreach (var v in staleKey.Values)
                        {
                            var data = v.ValueData?.ToString() ?? string.Empty;
                            dnsRegSection.Items.Add(new AnalysisItem
                            {
                                Name = $"StaleAdapter.{v.ValueName}",
                                Value = string.IsNullOrEmpty(data) ? "(present)" : data,
                                RegistryPath = $"{dnsRegBasePath}\\PrimaryDomainName\\StaleAdapter",
                                RegistryValue = $"{v.ValueName} = {data}"
                            });
                        }
                    }
                    else if (staleKey.SubKeys != null && staleKey.SubKeys.Count > 0)
                    {
                        // Some systems may create subkeys for ghosted NIC references
                        foreach (var sub in staleKey.SubKeys)
                        {
                            dnsRegSection.Items.Add(new AnalysisItem
                            {
                                Name = "StaleAdapter.GUID",
                                Value = sub.KeyName,
                                RegistryPath = $"{dnsRegBasePath}\\PrimaryDomainName\\StaleAdapter\\{sub.KeyName}",
                                RegistryValue = "(ghosted NIC)"
                            });
                        }
                    }
                    else
                    {
                        dnsRegSection.Items.Add(new AnalysisItem
                        {
                            Name = "StaleAdapter",
                            Value = "None",
                            RegistryPath = $"{dnsRegBasePath}\\PrimaryDomainName\\StaleAdapter",
                            RegistryValue = "No ghosted NICs recorded"
                        });
                    }
                }

                if (dnsRegSection.Items.Count > 0)
                {
                    sections.Add(dnsRegSection);
                }
            }

            // Shares (SYSTEM hive)
            var sharesKey = _parser.GetKey($@"{_currentControlSet}\Services\LanmanServer\Shares");
            var sharesSection = new AnalysisSection { Title = "📁 Network Shares" };
            
            if (sharesKey != null && sharesKey.Values.Count > 0)
            {
                foreach (var share in sharesKey.Values)
                {
                    var shareData = share.ValueData?.ToString() ?? "";
                    var path = "";
                    var remark = "";
                    var permissions = "";
                    foreach (var part in shareData.Split('\0'))
                    {
                        if (part.StartsWith("Path="))
                            path = part.Substring(5);
                        if (part.StartsWith("Remark="))
                            remark = part.Substring(7);
                        if (part.StartsWith("Permissions="))
                            permissions = part.Substring(12);
                    }

                    var details = new List<string>();
                    details.Add($"Path: {path}");
                    if (!string.IsNullOrEmpty(remark)) details.Add($"Remark: {remark}");
                    if (!string.IsNullOrEmpty(permissions)) details.Add($"Permissions: {permissions}");

                    sharesSection.Items.Add(new AnalysisItem
                    {
                        Name = share.ValueName,
                        Value = path,
                        RegistryPath = $@"{_currentControlSet}\Services\LanmanServer\Shares",
                        RegistryValue = string.Join(" | ", details)
                    });
                }
            }
            else
            {
                sharesSection.Items.Add(new AnalysisItem
                {
                    Name = "No network shares configured",
                    Value = "",
                    RegistryPath = $@"{_currentControlSet}\Services\LanmanServer\Shares",
                    RegistryValue = "No shares found in the registry"
                });
            }
            sections.Add(sharesSection);

            // NTLM Authentication Settings (SYSTEM hive)
            var ntlmSection = new AnalysisSection { Title = "🔑 NTLM Authentication" };
            string lsaPath = $@"{_currentControlSet}\Control\Lsa";
            var lsaKey = _parser.GetKey(lsaPath);
            
            if (lsaKey != null)
            {
                // LmCompatibilityLevel
                var lmCompatLevel = lsaKey.Values.FirstOrDefault(v => v.ValueName == "LmCompatibilityLevel")?.ValueData?.ToString() ?? "";
                var lmCompatText = lmCompatLevel switch
                {
                    "0" => "0 - Send LM & NTLM responses",
                    "1" => "1 - Send LM & NTLM, use NTLMv2 if negotiated",
                    "2" => "2 - Send NTLM response only",
                    "3" => "3 - Send NTLMv2 response only",
                    "4" => "4 - Send NTLMv2 only, refuse LM",
                    "5" => "5 - Send NTLMv2 only, refuse LM & NTLM",
                    _ when string.IsNullOrEmpty(lmCompatLevel) => "Not Configured (Default: 3)",
                    _ => lmCompatLevel
                };
                ntlmSection.Items.Add(new AnalysisItem
                {
                    Name = "LmCompatibilityLevel",
                    Value = lmCompatText,
                    RegistryPath = lsaPath,
                    RegistryValue = $"LmCompatibilityLevel = {lmCompatLevel}"
                });

                // Check MSV1_0 subkey for RestrictSendingNTLMTraffic
                string msv1Path = $@"{_currentControlSet}\Control\Lsa\MSV1_0";
                var msv1Key = _parser.GetKey(msv1Path);
                if (msv1Key != null)
                {
                    var restrictNtlm = msv1Key.Values.FirstOrDefault(v => v.ValueName == "RestrictSendingNTLMTraffic")?.ValueData?.ToString() ?? "";
                    var restrictNtlmText = restrictNtlm switch
                    {
                        "0" => "0 - Allow all",
                        "1" => "1 - Audit all",
                        "2" => "2 - Deny all",
                        _ when string.IsNullOrEmpty(restrictNtlm) => "Not Configured (Allow all)",
                        _ => restrictNtlm
                    };
                    ntlmSection.Items.Add(new AnalysisItem
                    {
                        Name = "RestrictSendingNTLMTraffic",
                        Value = restrictNtlmText,
                        RegistryPath = msv1Path,
                        RegistryValue = $"RestrictSendingNTLMTraffic = {restrictNtlm}"
                    });

                    // Also check for AuditReceivingNTLMTraffic
                    var auditNtlm = msv1Key.Values.FirstOrDefault(v => v.ValueName == "AuditReceivingNTLMTraffic")?.ValueData?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(auditNtlm))
                    {
                        var auditNtlmText = auditNtlm switch
                        {
                            "0" => "0 - Disabled",
                            "1" => "1 - Audit for domain accounts",
                            "2" => "2 - Audit all accounts",
                            _ => auditNtlm
                        };
                        ntlmSection.Items.Add(new AnalysisItem
                        {
                            Name = "AuditReceivingNTLMTraffic",
                            Value = auditNtlmText,
                            RegistryPath = msv1Path,
                            RegistryValue = $"AuditReceivingNTLMTraffic = {auditNtlm}"
                        });
                    }

                    // RestrictReceivingNTLMTraffic
                    var restrictReceiving = msv1Key.Values.FirstOrDefault(v => v.ValueName == "RestrictReceivingNTLMTraffic")?.ValueData?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(restrictReceiving))
                    {
                        var restrictReceivingText = restrictReceiving switch
                        {
                            "0" => "0 - Allow all",
                            "1" => "1 - Deny for domain accounts",
                            "2" => "2 - Deny all",
                            _ => restrictReceiving
                        };
                        ntlmSection.Items.Add(new AnalysisItem
                        {
                            Name = "RestrictReceivingNTLMTraffic",
                            Value = restrictReceivingText,
                            RegistryPath = msv1Path,
                            RegistryValue = $"RestrictReceivingNTLMTraffic = {restrictReceiving}"
                        });
                    }
                }
                else
                {
                    ntlmSection.Items.Add(new AnalysisItem
                    {
                        Name = "RestrictSendingNTLMTraffic",
                        Value = "Not Configured (Allow all)",
                        RegistryPath = msv1Path,
                        RegistryValue = "MSV1_0 key not found"
                    });
                }
            }
            else
            {
                ntlmSection.Items.Add(new AnalysisItem
                {
                    Name = "LmCompatibilityLevel",
                    Value = "Not Configured",
                    RegistryPath = lsaPath,
                    RegistryValue = "Lsa key not found"
                });
            }
            sections.Add(ntlmSection);

            // TLS/SSL Protocol Settings (SYSTEM hive)
            var schannelPath = $@"{_currentControlSet}\Control\SecurityProviders\SCHANNEL\Protocols";
            var schannelKey = _parser.GetKey(schannelPath);
            if (schannelKey?.SubKeys != null && schannelKey.SubKeys.Count > 0)
            {
                var tlsSection = new AnalysisSection { Title = "🔐 TLS/SSL Protocols" };
                
                // Common protocols to check
                var protocolOrder = new[] { "SSL 2.0", "SSL 3.0", "TLS 1.0", "TLS 1.1", "TLS 1.2", "TLS 1.3" };
                
                foreach (var protocolName in protocolOrder)
                {
                    var protocolKey = schannelKey.SubKeys.FirstOrDefault(s => s.KeyName.Equals(protocolName, StringComparison.OrdinalIgnoreCase));
                    if (protocolKey?.SubKeys != null)
                    {
                        var subItems = new List<AnalysisItem>();
                        var protocolPath = $@"{schannelPath}\{protocolKey.KeyName}";
                        
                        // Check Client settings
                        var clientKey = protocolKey.SubKeys.FirstOrDefault(s => s.KeyName.Equals("Client", StringComparison.OrdinalIgnoreCase));
                        if (clientKey != null)
                        {
                            var clientEnabled = clientKey.Values.FirstOrDefault(v => v.ValueName == "Enabled")?.ValueData?.ToString();
                            var clientDisabledByDefault = clientKey.Values.FirstOrDefault(v => v.ValueName == "DisabledByDefault")?.ValueData?.ToString();
                            
                            var clientStatus = GetProtocolStatus(clientEnabled, clientDisabledByDefault);
                            subItems.Add(new AnalysisItem
                            {
                                Name = "Client",
                                Value = clientStatus,
                                RegistryPath = $@"{protocolPath}\Client",
                                RegistryValue = $"Enabled={clientEnabled ?? "Not Set"}, DisabledByDefault={clientDisabledByDefault ?? "Not Set"}"
                            });
                        }
                        else
                        {
                            subItems.Add(new AnalysisItem
                            {
                                Name = "Client",
                                Value = "Default (OS Controlled)",
                                RegistryPath = $@"{protocolPath}\Client",
                                RegistryValue = "Key not present - using OS defaults"
                            });
                        }
                        
                        // Check Server settings
                        var serverKey = protocolKey.SubKeys.FirstOrDefault(s => s.KeyName.Equals("Server", StringComparison.OrdinalIgnoreCase));
                        if (serverKey != null)
                        {
                            var serverEnabled = serverKey.Values.FirstOrDefault(v => v.ValueName == "Enabled")?.ValueData?.ToString();
                            var serverDisabledByDefault = serverKey.Values.FirstOrDefault(v => v.ValueName == "DisabledByDefault")?.ValueData?.ToString();
                            
                            var serverStatus = GetProtocolStatus(serverEnabled, serverDisabledByDefault);
                            subItems.Add(new AnalysisItem
                            {
                                Name = "Server",
                                Value = serverStatus,
                                RegistryPath = $@"{protocolPath}\Server",
                                RegistryValue = $"Enabled={serverEnabled ?? "Not Set"}, DisabledByDefault={serverDisabledByDefault ?? "Not Set"}"
                            });
                        }
                        else
                        {
                            subItems.Add(new AnalysisItem
                            {
                                Name = "Server",
                                Value = "Default (OS Controlled)",
                                RegistryPath = $@"{protocolPath}\Server",
                                RegistryValue = "Key not present - using OS defaults"
                            });
                        }
                        
                        // Determine overall protocol status
                        var overallStatus = DetermineOverallProtocolStatus(protocolKey);
                        
                        tlsSection.Items.Add(new AnalysisItem
                        {
                            Name = protocolKey.KeyName,
                            Value = overallStatus,
                            IsSubSection = true,
                            SubItems = subItems,
                            RegistryPath = protocolPath,
                            RegistryValue = $"Protocol: {protocolKey.KeyName}"
                        });
                    }
                }
                
                // Check for any other protocols not in our standard list
                foreach (var protocolKey in schannelKey.SubKeys)
                {
                    if (!protocolOrder.Any(p => p.Equals(protocolKey.KeyName, StringComparison.OrdinalIgnoreCase)))
                    {
                        var subItems = new List<AnalysisItem>();
                        var protocolPath = $@"{schannelPath}\{protocolKey.KeyName}";
                        
                        if (protocolKey.SubKeys != null)
                        {
                            foreach (var subKey in protocolKey.SubKeys)
                            {
                                var enabled = subKey.Values.FirstOrDefault(v => v.ValueName == "Enabled")?.ValueData?.ToString();
                                var disabledByDefault = subKey.Values.FirstOrDefault(v => v.ValueName == "DisabledByDefault")?.ValueData?.ToString();
                                
                                subItems.Add(new AnalysisItem
                                {
                                    Name = subKey.KeyName,
                                    Value = GetProtocolStatus(enabled, disabledByDefault),
                                    RegistryPath = $@"{protocolPath}\{subKey.KeyName}",
                                    RegistryValue = $"Enabled={enabled ?? "Not Set"}, DisabledByDefault={disabledByDefault ?? "Not Set"}"
                                });
                            }
                        }
                        
                        var overallStatus = DetermineOverallProtocolStatus(protocolKey);
                        
                        tlsSection.Items.Add(new AnalysisItem
                        {
                            Name = protocolKey.KeyName,
                            Value = overallStatus,
                            IsSubSection = subItems.Count > 0,
                            SubItems = subItems,
                            RegistryPath = protocolPath,
                            RegistryValue = $"Protocol: {protocolKey.KeyName}"
                        });
                    }
                }
                
                if (tlsSection.Items.Count > 0)
                    sections.Add(tlsSection);
            }

            // Windows Firewall Settings (SYSTEM hive)
            string firewallBasePath = $@"{_currentControlSet}\Services\SharedAccess\Parameters\FirewallPolicy";
            var firewallBaseKey = _parser.GetKey(firewallBasePath);
            if (firewallBaseKey != null)
            {
                var firewallSection = new AnalysisSection { Title = "🔥 Windows Firewall" };
                
                // Check each profile: DomainProfile, StandardProfile, PublicProfile
                var profiles = new[] { ("DomainProfile", "Domain Profile"), ("StandardProfile", "Private Profile"), ("PublicProfile", "Public Profile") };
                foreach (var (profileName, displayName) in profiles)
                {
                    var profileKey = _parser.GetKey($@"{firewallBasePath}\{profileName}");
                    var profilePath = $@"{firewallBasePath}\{profileName}";
                    
                    var enableFirewall = profileKey?.Values.FirstOrDefault(v => v.ValueName == "EnableFirewall")?.ValueData?.ToString() ?? "";
                    var firewallStatus = enableFirewall switch
                    {
                        "1" => "Enabled",
                        "0" => "Disabled",
                        _ => "Not Configured"
                    };
                    
                    // Get rule counts for this profile
                    var allRules = GetFirewallRulesForProfile(profileName);
                    int activeCount = 0, blockCount = 0;
                    foreach (var rule in allRules)
                    {
                        if (rule.IsActive)
                        {
                            activeCount++;
                            if (rule.Action == "Block")
                                blockCount++;
                        }
                    }
                    
                    var statusWithCounts = $"{(enableFirewall == "1" ? "✅" : "❌")} {firewallStatus} | Rules: {allRules.Count} total, {activeCount} active, {blockCount} blocking";
                    
                    firewallSection.Items.Add(new AnalysisItem
                    {
                        Name = displayName,
                        Value = statusWithCounts,
                        RegistryPath = profilePath,
                        RegistryValue = $"EnableFirewall = {enableFirewall}"
                    });
                }
                
                sections.Add(firewallSection);
            }

            if (sections.Count == 0)
            {
                var noDataSection = new AnalysisSection { Title = "ℹ️ Notice" };
                noDataSection.Items.Add(new AnalysisItem { Name = "Info", Value = "No network information found in this hive type" });
                noDataSection.Items.Add(new AnalysisItem { Name = "Tip", Value = "Load SYSTEM or SOFTWARE hive for network info" });
                sections.Add(noDataSection);
            }

            return sections;
        }

        #endregion

        #region Windows Update Analysis (SOFTWARE hive)

        /// <summary>
        /// Get Windows Update registry settings (requires SOFTWARE hive)
        /// </summary>
        public List<AnalysisSection> GetUpdateAnalysis()
        {
            var sections = new List<AnalysisSection>();

            // Check if this is SOFTWARE hive
            if (_parser.CurrentHiveType != OfflineRegistryParser.HiveType.SOFTWARE)
            {
                var noticeSection = new AnalysisSection { Title = "ℹ️ Notice" };
                noticeSection.Items.Add(new AnalysisItem
                {
                    Name = "Info",
                    Value = "This information requires the SOFTWARE hive to be loaded",
                    RegistryPath = "",
                    RegistryValue = "Load SOFTWARE hive for update policies and settings"
                });
                sections.Add(noticeSection);
                return sections;
            }

            // Windows Update Policy Settings
            var policySection = new AnalysisSection { Title = "📋 Update Policy" };
            
            // Group Policy Windows Update settings
            var wuPolicyPath = @"Policies\Microsoft\Windows\WindowsUpdate";
            var wuPolicyKey = _parser.GetKey(wuPolicyPath);
            
            // WSUS/WUfB Settings
            var wuServer = GetValue(wuPolicyPath, "WUServer");
            if (!string.IsNullOrEmpty(wuServer))
            {
                policySection.Items.Add(new AnalysisItem
                {
                    Name = "WUServer",
                    Value = wuServer,
                    RegistryPath = wuPolicyPath,
                    RegistryValue = $"WUServer = {wuServer}"
                });
            }

            var wuStatusServer = GetValue(wuPolicyPath, "WUStatusServer");
            if (!string.IsNullOrEmpty(wuStatusServer))
            {
                policySection.Items.Add(new AnalysisItem
                {
                    Name = "WUStatusServer",
                    Value = wuStatusServer,
                    RegistryPath = wuPolicyPath,
                    RegistryValue = $"WUStatusServer = {wuStatusServer}"
                });
            }

            var useWUServer = GetValue(wuPolicyPath, "UseWUServer");
            if (!string.IsNullOrEmpty(useWUServer))
            {
                policySection.Items.Add(new AnalysisItem
                {
                    Name = "UseWUServer",
                    Value = useWUServer == "1" ? "Yes (WSUS enabled)" : "No",
                    RegistryPath = wuPolicyPath,
                    RegistryValue = $"UseWUServer = {useWUServer}"
                });
            }

            // Target Group for WSUS
            var targetGroup = GetValue(wuPolicyPath, "TargetGroup");
            if (!string.IsNullOrEmpty(targetGroup))
            {
                policySection.Items.Add(new AnalysisItem
                {
                    Name = "TargetGroup",
                    Value = targetGroup,
                    RegistryPath = wuPolicyPath,
                    RegistryValue = $"TargetGroup = {targetGroup}"
                });
            }

            var targetGroupEnabled = GetValue(wuPolicyPath, "TargetGroupEnabled");
            if (!string.IsNullOrEmpty(targetGroupEnabled))
            {
                policySection.Items.Add(new AnalysisItem
                {
                    Name = "TargetGroupEnabled",
                    Value = targetGroupEnabled == "1" ? "Yes" : "No",
                    RegistryPath = wuPolicyPath,
                    RegistryValue = $"TargetGroupEnabled = {targetGroupEnabled}"
                });
            }

            // Automatic Update settings (AU subkey)
            var auPath = @"Policies\Microsoft\Windows\WindowsUpdate\AU";
            var auKey = _parser.GetKey(auPath);

            var noAutoUpdate = GetValue(auPath, "NoAutoUpdate");
            if (!string.IsNullOrEmpty(noAutoUpdate))
            {
                policySection.Items.Add(new AnalysisItem
                {
                    Name = "NoAutoUpdate",
                    Value = noAutoUpdate == "1" ? "Yes (Auto Update Disabled)" : "No",
                    RegistryPath = auPath,
                    RegistryValue = $"NoAutoUpdate = {noAutoUpdate}"
                });
            }

            var auOptions = GetValue(auPath, "AUOptions");
            if (!string.IsNullOrEmpty(auOptions))
            {
                var optionText = auOptions switch
                {
                    "2" => "2 - Notify before download",
                    "3" => "3 - Auto download, notify before install",
                    "4" => "4 - Auto download and install",
                    "5" => "5 - Allow local admin to choose",
                    "7" => "7 - Auto download, notify, schedule install",
                    _ => auOptions
                };
                policySection.Items.Add(new AnalysisItem
                {
                    Name = "AUOptions",
                    Value = optionText,
                    RegistryPath = auPath,
                    RegistryValue = $"AUOptions = {auOptions}"
                });
            }

            var scheduledInstallDay = GetValue(auPath, "ScheduledInstallDay");
            if (!string.IsNullOrEmpty(scheduledInstallDay))
            {
                var dayText = scheduledInstallDay switch
                {
                    "0" => "0 - Every day",
                    "1" => "1 - Sunday",
                    "2" => "2 - Monday",
                    "3" => "3 - Tuesday",
                    "4" => "4 - Wednesday",
                    "5" => "5 - Thursday",
                    "6" => "6 - Friday",
                    "7" => "7 - Saturday",
                    _ => scheduledInstallDay
                };
                policySection.Items.Add(new AnalysisItem
                {
                    Name = "ScheduledInstallDay",
                    Value = dayText,
                    RegistryPath = auPath,
                    RegistryValue = $"ScheduledInstallDay = {scheduledInstallDay}"
                });
            }

            var scheduledInstallTime = GetValue(auPath, "ScheduledInstallTime");
            if (!string.IsNullOrEmpty(scheduledInstallTime))
            {
                policySection.Items.Add(new AnalysisItem
                {
                    Name = "ScheduledInstallTime",
                    Value = $"{scheduledInstallTime}:00",
                    RegistryPath = auPath,
                    RegistryValue = $"ScheduledInstallTime = {scheduledInstallTime}"
                });
            }

            // Include recommended updates
            var includeRecommended = GetValue(auPath, "IncludeRecommendedUpdates");
            if (!string.IsNullOrEmpty(includeRecommended))
            {
                policySection.Items.Add(new AnalysisItem
                {
                    Name = "IncludeRecommendedUpdates",
                    Value = includeRecommended == "1" ? "Yes" : "No",
                    RegistryPath = auPath,
                    RegistryValue = $"IncludeRecommendedUpdates = {includeRecommended}"
                });
            }

            // Auto reboot settings
            var noAutoRebootWithLoggedOnUsers = GetValue(auPath, "NoAutoRebootWithLoggedOnUsers");
            if (!string.IsNullOrEmpty(noAutoRebootWithLoggedOnUsers))
            {
                policySection.Items.Add(new AnalysisItem
                {
                    Name = "NoAutoRebootWithLoggedOnUsers",
                    Value = noAutoRebootWithLoggedOnUsers == "1" ? "Yes (No reboot with users)" : "No",
                    RegistryPath = auPath,
                    RegistryValue = $"NoAutoRebootWithLoggedOnUsers = {noAutoRebootWithLoggedOnUsers}"
                });
            }

            // UseWUServer in AU path (overrides parent setting)
            var useWUServerAU = GetValue(auPath, "UseWUServer");
            if (!string.IsNullOrEmpty(useWUServerAU))
            {
                policySection.Items.Add(new AnalysisItem
                {
                    Name = "UseWUServer (AU)",
                    Value = useWUServerAU == "1" ? "Yes (WSUS enabled)" : "No",
                    RegistryPath = auPath,
                    RegistryValue = $"UseWUServer = {useWUServerAU}"
                });
            }

            // NoAUShutdownOption - Disable shutdown option in Windows Update UI
            var noAUShutdownOption = GetValue(auPath, "NoAUShutdownOption");
            if (!string.IsNullOrEmpty(noAUShutdownOption))
            {
                policySection.Items.Add(new AnalysisItem
                {
                    Name = "NoAUShutdownOption",
                    Value = noAUShutdownOption == "1" ? "Yes (Shutdown option hidden)" : "No",
                    RegistryPath = auPath,
                    RegistryValue = $"NoAUShutdownOption = {noAUShutdownOption}"
                });
            }

            // AlwaysAutoRebootAtScheduledTime - Force reboot at scheduled time
            var alwaysAutoReboot = GetValue(auPath, "AlwaysAutoRebootAtScheduledTime");
            if (!string.IsNullOrEmpty(alwaysAutoReboot))
            {
                policySection.Items.Add(new AnalysisItem
                {
                    Name = "AlwaysAutoRebootAtScheduledTime",
                    Value = alwaysAutoReboot == "1" ? "Yes (Force reboot)" : "No",
                    RegistryPath = auPath,
                    RegistryValue = $"AlwaysAutoRebootAtScheduledTime = {alwaysAutoReboot}"
                });
            }

            // AlwaysAutoRebootAtScheduledTimeMinutes - Minutes to wait before reboot
            var alwaysAutoRebootMinutes = GetValue(auPath, "AlwaysAutoRebootAtScheduledTimeMinutes");
            if (!string.IsNullOrEmpty(alwaysAutoRebootMinutes))
            {
                policySection.Items.Add(new AnalysisItem
                {
                    Name = "AlwaysAutoRebootAtScheduledTimeMinutes",
                    Value = $"{alwaysAutoRebootMinutes} minutes",
                    RegistryPath = auPath,
                    RegistryValue = $"AlwaysAutoRebootAtScheduledTimeMinutes = {alwaysAutoRebootMinutes}"
                });
            }

            sections.Add(policySection);

            // Windows Update for Business (WUfB) settings
            var wufbSection = new AnalysisSection { Title = "🏢 Windows Update for Business" };

            // Defer feature updates
            var deferFeatureUpdates = GetValue(wuPolicyPath, "DeferFeatureUpdates");
            if (!string.IsNullOrEmpty(deferFeatureUpdates))
            {
                wufbSection.Items.Add(new AnalysisItem
                {
                    Name = "DeferFeatureUpdates",
                    Value = deferFeatureUpdates == "1" ? "Yes" : "No",
                    RegistryPath = wuPolicyPath,
                    RegistryValue = $"DeferFeatureUpdates = {deferFeatureUpdates}"
                });
            }

            var deferFeatureUpdatesPeriod = GetValue(wuPolicyPath, "DeferFeatureUpdatesPeriodInDays");
            if (!string.IsNullOrEmpty(deferFeatureUpdatesPeriod))
            {
                wufbSection.Items.Add(new AnalysisItem
                {
                    Name = "DeferFeatureUpdatesPeriodInDays",
                    Value = $"{deferFeatureUpdatesPeriod} days",
                    RegistryPath = wuPolicyPath,
                    RegistryValue = $"DeferFeatureUpdatesPeriodInDays = {deferFeatureUpdatesPeriod}"
                });
            }

            // Defer quality updates
            var deferQualityUpdates = GetValue(wuPolicyPath, "DeferQualityUpdates");
            if (!string.IsNullOrEmpty(deferQualityUpdates))
            {
                wufbSection.Items.Add(new AnalysisItem
                {
                    Name = "DeferQualityUpdates",
                    Value = deferQualityUpdates == "1" ? "Yes" : "No",
                    RegistryPath = wuPolicyPath,
                    RegistryValue = $"DeferQualityUpdates = {deferQualityUpdates}"
                });
            }

            var deferQualityUpdatesPeriod = GetValue(wuPolicyPath, "DeferQualityUpdatesPeriodInDays");
            if (!string.IsNullOrEmpty(deferQualityUpdatesPeriod))
            {
                wufbSection.Items.Add(new AnalysisItem
                {
                    Name = "DeferQualityUpdatesPeriodInDays",
                    Value = $"{deferQualityUpdatesPeriod} days",
                    RegistryPath = wuPolicyPath,
                    RegistryValue = $"DeferQualityUpdatesPeriodInDays = {deferQualityUpdatesPeriod}"
                });
            }

            // Pause updates
            var pauseFeatureUpdates = GetValue(wuPolicyPath, "PauseFeatureUpdates");
            if (!string.IsNullOrEmpty(pauseFeatureUpdates))
            {
                wufbSection.Items.Add(new AnalysisItem
                {
                    Name = "PauseFeatureUpdates",
                    Value = pauseFeatureUpdates == "1" ? "Yes (Paused)" : "No",
                    RegistryPath = wuPolicyPath,
                    RegistryValue = $"PauseFeatureUpdates = {pauseFeatureUpdates}"
                });
            }

            var pauseQualityUpdates = GetValue(wuPolicyPath, "PauseQualityUpdates");
            if (!string.IsNullOrEmpty(pauseQualityUpdates))
            {
                wufbSection.Items.Add(new AnalysisItem
                {
                    Name = "PauseQualityUpdates",
                    Value = pauseQualityUpdates == "1" ? "Yes (Paused)" : "No",
                    RegistryPath = wuPolicyPath,
                    RegistryValue = $"PauseQualityUpdates = {pauseQualityUpdates}"
                });
            }

            // Target release version
            var targetReleaseVersion = GetValue(wuPolicyPath, "TargetReleaseVersion");
            if (!string.IsNullOrEmpty(targetReleaseVersion))
            {
                wufbSection.Items.Add(new AnalysisItem
                {
                    Name = "TargetReleaseVersion",
                    Value = targetReleaseVersion == "1" ? "Enabled" : "Disabled",
                    RegistryPath = wuPolicyPath,
                    RegistryValue = $"TargetReleaseVersion = {targetReleaseVersion}"
                });
            }

            var targetReleaseVersionInfo = GetValue(wuPolicyPath, "TargetReleaseVersionInfo");
            if (!string.IsNullOrEmpty(targetReleaseVersionInfo))
            {
                wufbSection.Items.Add(new AnalysisItem
                {
                    Name = "TargetReleaseVersionInfo",
                    Value = targetReleaseVersionInfo,
                    RegistryPath = wuPolicyPath,
                    RegistryValue = $"TargetReleaseVersionInfo = {targetReleaseVersionInfo}"
                });
            }

            var productVersion = GetValue(wuPolicyPath, "ProductVersion");
            if (!string.IsNullOrEmpty(productVersion))
            {
                wufbSection.Items.Add(new AnalysisItem
                {
                    Name = "ProductVersion",
                    Value = productVersion,
                    RegistryPath = wuPolicyPath,
                    RegistryValue = $"ProductVersion = {productVersion}"
                });
            }

            // Branch readiness level
            var branchReadinessLevel = GetValue(wuPolicyPath, "BranchReadinessLevel");
            if (!string.IsNullOrEmpty(branchReadinessLevel))
            {
                var levelText = branchReadinessLevel switch
                {
                    "2" => "2 - Semi-Annual Channel (Targeted)",
                    "4" => "4 - Semi-Annual Channel",
                    "8" => "8 - Release Preview",
                    "16" => "16 - Beta Channel",
                    "32" => "32 - Dev Channel",
                    _ => branchReadinessLevel
                };
                wufbSection.Items.Add(new AnalysisItem
                {
                    Name = "BranchReadinessLevel",
                    Value = levelText,
                    RegistryPath = wuPolicyPath,
                    RegistryValue = $"BranchReadinessLevel = {branchReadinessLevel}"
                });
            }

            if (wufbSection.Items.Count > 0)
                sections.Add(wufbSection);

            // Delivery Optimization settings
            var doSection = new AnalysisSection { Title = "📦 Delivery Optimization" };
            var doPath = @"Policies\Microsoft\Windows\DeliveryOptimization";
            var doKey = _parser.GetKey(doPath);

            var downloadMode = GetValue(doPath, "DODownloadMode");
            if (!string.IsNullOrEmpty(downloadMode))
            {
                var modeText = downloadMode switch
                {
                    "0" => "0 - HTTP only (no peering)",
                    "1" => "1 - LAN only",
                    "2" => "2 - Group (LAN + same domain)",
                    "3" => "3 - Internet (LAN + Internet peers)",
                    "99" => "99 - Simple mode (fallback)",
                    "100" => "100 - Bypass mode (BITS only)",
                    _ => downloadMode
                };
                doSection.Items.Add(new AnalysisItem
                {
                    Name = "DODownloadMode",
                    Value = modeText,
                    RegistryPath = doPath,
                    RegistryValue = $"DODownloadMode = {downloadMode}"
                });
            }

            var groupId = GetValue(doPath, "DOGroupId");
            if (!string.IsNullOrEmpty(groupId))
            {
                doSection.Items.Add(new AnalysisItem
                {
                    Name = "DOGroupId",
                    Value = groupId,
                    RegistryPath = doPath,
                    RegistryValue = $"DOGroupId = {groupId}"
                });
            }

            var maxCacheSize = GetValue(doPath, "DOMaxCacheSize");
            if (!string.IsNullOrEmpty(maxCacheSize))
            {
                doSection.Items.Add(new AnalysisItem
                {
                    Name = "DOMaxCacheSize",
                    Value = $"{maxCacheSize}%",
                    RegistryPath = doPath,
                    RegistryValue = $"DOMaxCacheSize = {maxCacheSize}"
                });
            }

            var maxCacheAge = GetValue(doPath, "DOMaxCacheAge");
            if (!string.IsNullOrEmpty(maxCacheAge))
            {
                var days = int.TryParse(maxCacheAge, out int seconds) ? seconds / 86400 : 0;
                doSection.Items.Add(new AnalysisItem
                {
                    Name = "DOMaxCacheAge",
                    Value = $"{days} days ({maxCacheAge} seconds)",
                    RegistryPath = doPath,
                    RegistryValue = $"DOMaxCacheAge = {maxCacheAge}"
                });
            }

            if (doSection.Items.Count > 0)
                sections.Add(doSection);

            // Update Configuration - UX Settings
            var configSection = new AnalysisSection { Title = "📜 Update Configuration" };
            
            // Check UX settings - show ALL values from this key
            var uxPath = @"Microsoft\WindowsUpdate\UX\Settings";
            var uxKey = _parser.GetKey(uxPath);
            
            if (uxKey != null && uxKey.Values != null && uxKey.Values.Any())
            {
                foreach (var val in uxKey.Values.OrderBy(v => v.ValueName))
                {
                    var valName = val.ValueName ?? "(Default)";
                    var valData = val.ValueData?.ToString() ?? "";
                    
                    // Format specific values for better readability
                    string displayValue = valData;
                    
                    if (valName == "ActiveHoursStart" || valName == "ActiveHoursEnd")
                    {
                        displayValue = $"{valData}:00";
                    }
                    else if (valName == "BranchReadinessLevel")
                    {
                        displayValue = valData switch
                        {
                            "2" => "2 - Semi-Annual Channel (Targeted)",
                            "4" => "4 - Semi-Annual Channel",
                            "8" => "8 - Release Preview",
                            "16" => "16 - Beta Channel",
                            "32" => "32 - Dev Channel",
                            _ => valData
                        };
                    }
                    else if (valName == "UxOption")
                    {
                        displayValue = valData switch
                        {
                            "0" => "0 - Automatic",
                            "1" => "1 - Notify to schedule restart",
                            "2" => "2 - Automatic with scheduled restart",
                            "3" => "3 - Automatic, let user decide restart time",
                            "4" => "4 - Automatic, allow user to reschedule",
                            _ => valData
                        };
                    }
                    else if (valName.Contains("Enabled") || valName.Contains("Allow") || valName.StartsWith("Exclude") || valName == "FlightCommitted")
                    {
                        displayValue = valData == "1" ? "Yes" : "No";
                    }
                    else if (valName.EndsWith("PeriodInDays"))
                    {
                        displayValue = $"{valData} days";
                    }
                    
                    configSection.Items.Add(new AnalysisItem
                    {
                        Name = valName,
                        Value = displayValue,
                        RegistryPath = uxPath,
                        RegistryValue = $"{valName} = {valData}"
                    });
                }
            }

            // Check Windows Update Agent info
            var wuaPath = @"Microsoft\Windows\CurrentVersion\WindowsUpdate";
            var wuaKey = _parser.GetKey(wuaPath);
            if (wuaKey != null && wuaKey.Values != null)
            {
                foreach (var val in wuaKey.Values.OrderBy(v => v.ValueName))
                {
                    var valName = val.ValueName ?? "(Default)";
                    var valData = val.ValueData?.ToString() ?? "";
                    
                    configSection.Items.Add(new AnalysisItem
                    {
                        Name = valName,
                        Value = valData,
                        RegistryPath = wuaPath,
                        RegistryValue = $"{valName} = {valData}"
                    });
                }
            }
            
            // Check Auto Update Results
            var autoUpdateResultsPath = @"Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\Results";
            var installResultsPath = $@"{autoUpdateResultsPath}\Install";
            var downloadResultsPath = $@"{autoUpdateResultsPath}\Download";
            
            var installKey = _parser.GetKey(installResultsPath);
            if (installKey != null && installKey.Values != null)
            {
                foreach (var val in installKey.Values)
                {
                    var valName = val.ValueName ?? "(Default)";
                    var valData = val.ValueData?.ToString() ?? "";
                    
                    configSection.Items.Add(new AnalysisItem
                    {
                        Name = $"Install {valName}",
                        Value = valData,
                        RegistryPath = installResultsPath,
                        RegistryValue = $"{valName} = {valData}"
                    });
                }
            }
            
            var downloadKey = _parser.GetKey(downloadResultsPath);
            if (downloadKey != null && downloadKey.Values != null)
            {
                foreach (var val in downloadKey.Values)
                {
                    var valName = val.ValueName ?? "(Default)";
                    var valData = val.ValueData?.ToString() ?? "";
                    
                    configSection.Items.Add(new AnalysisItem
                    {
                        Name = $"Download {valName}",
                        Value = valData,
                        RegistryPath = downloadResultsPath,
                        RegistryValue = $"{valName} = {valData}"
                    });
                }
            }
            
            // Check for Reboot Required
            var rebootRequiredPath = @"Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired";
            var rebootKey = _parser.GetKey(rebootRequiredPath);
            if (rebootKey != null)
            {
                configSection.Items.Add(new AnalysisItem
                {
                    Name = "RebootRequired",
                    Value = "Yes - Pending reboot for updates",
                    RegistryPath = rebootRequiredPath,
                    RegistryValue = "RebootRequired key exists"
                });
            }

            if (configSection.Items.Count > 0)
                sections.Add(configSection);

            // SSU (Servicing Stack Update) Version
            var ssuSection = new AnalysisSection { Title = "🔧 Servicing Stack Update (SSU)" };
            var cbsPath = @"Microsoft\Windows\CurrentVersion\Component Based Servicing";
            var cbsVersionPath = @"Microsoft\Windows\CurrentVersion\Component Based Servicing\Version";
            var cbsKey = _parser.GetKey(cbsPath);
            var versionKey = _parser.GetKey(cbsVersionPath);
            
            // The SSU version is stored as a VALUE NAME under the Version subkey
            // e.g., the value name "10.0.17763.8020" contains the path to servicing stack
            if (versionKey != null && versionKey.Values.Any())
            {
                // Get all version entries (value names are the versions)
                var versions = versionKey.Values
                    .Where(v => !string.IsNullOrEmpty(v.ValueName) && v.ValueName.StartsWith("10."))
                    .OrderByDescending(v => v.ValueName)
                    .ToList();

                if (versions.Count > 0)
                {
                    // Show the latest (highest) version
                    var latestVersion = versions.First();
                    ssuSection.Items.Add(new AnalysisItem
                    {
                        Name = "SSU Version",
                        Value = latestVersion.ValueName,
                        RegistryPath = cbsVersionPath,
                        RegistryValue = $"{latestVersion.ValueName} = {latestVersion.ValueData}"
                    });

                    // If there are multiple versions, show count
                    if (versions.Count > 1)
                    {
                        ssuSection.Items.Add(new AnalysisItem
                        {
                            Name = "Installed SSU Count",
                            Value = versions.Count.ToString(),
                            RegistryPath = cbsVersionPath,
                            RegistryValue = $"Total SSU versions installed: {versions.Count}"
                        });

                        // Show oldest version too for reference
                        var oldestVersion = versions.Last();
                        ssuSection.Items.Add(new AnalysisItem
                        {
                            Name = "Oldest SSU",
                            Value = oldestVersion.ValueName,
                            RegistryPath = cbsVersionPath,
                            RegistryValue = $"{oldestVersion.ValueName} = {oldestVersion.ValueData}"
                        });
                    }
                }
            }
            
            if (cbsKey != null)
            {
                // Also check for additional CBS info if available
                var lastSuccessfulBoot = GetValue(cbsPath, "LastSuccessfulBoot");
                if (!string.IsNullOrEmpty(lastSuccessfulBoot))
                {
                    ssuSection.Items.Add(new AnalysisItem
                    {
                        Name = "LastSuccessfulBoot",
                        Value = lastSuccessfulBoot,
                        RegistryPath = cbsPath,
                        RegistryValue = $"LastSuccessfulBoot = {lastSuccessfulBoot}"
                    });
                }

                var rebootPending = GetValue(cbsPath, "RebootPending");
                if (!string.IsNullOrEmpty(rebootPending))
                {
                    ssuSection.Items.Add(new AnalysisItem
                    {
                        Name = "RebootPending",
                        Value = rebootPending == "1" ? "Yes" : "No",
                        RegistryPath = cbsPath,
                        RegistryValue = $"RebootPending = {rebootPending}"
                    });
                }
            }
            
            if (ssuSection.Items.Count == 0)
            {
                ssuSection.Items.Add(new AnalysisItem
                {
                    Name = "Status",
                    Value = "SSU information not found",
                    RegistryPath = cbsVersionPath,
                    RegistryValue = "Component Based Servicing\\Version not present"
                });
            }
            
            sections.Add(ssuSection);

            // If no update settings found at all
            if (sections.All(s => s.Items.Count == 0))
            {
                var defaultSection = new AnalysisSection { Title = "📋 Update Policy" };
                defaultSection.Items.Add(new AnalysisItem
                {
                    Name = "Status",
                    Value = "No custom update policies configured",
                    RegistryPath = wuPolicyPath,
                    RegistryValue = "Using default Windows Update settings"
                });
                sections.Clear();
                sections.Add(defaultSection);
            }

            // Add CBS Packages section (will be specially handled in UI)
            var packagesSection = new AnalysisSection { Title = "📦 CBS Packages" };
            packagesSection.Items.Add(new AnalysisItem
            {
                Name = "Click to view",
                Value = "Component Based Servicing packages installed on this system",
                RegistryPath = @"Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages",
                RegistryValue = "This section shows all CBS packages with their install state, time, and user"
            });
            sections.Add(packagesSection);

            return sections;
        }

        /// <summary>
        /// Get CBS Packages information (requires SOFTWARE hive)
        /// Shows installed packages from Component Based Servicing
        /// </summary>
        public List<AnalysisSection> GetPackagesAnalysis()
        {
            var sections = new List<AnalysisSection>();

            // Check if this is SOFTWARE hive
            if (_parser.CurrentHiveType != OfflineRegistryParser.HiveType.SOFTWARE)
            {
                var noticeSection = new AnalysisSection { Title = "ℹ️ Notice" };
                noticeSection.Items.Add(new AnalysisItem
                {
                    Name = "Info",
                    Value = "This information requires the SOFTWARE hive to be loaded",
                    RegistryPath = "",
                    RegistryValue = "Load SOFTWARE hive for package information"
                });
                sections.Add(noticeSection);
                return sections;
            }

            var packagesPath = @"Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages";
            var packagesKey = _parser.GetKey(packagesPath);

            if (packagesKey == null || packagesKey.SubKeys == null || !packagesKey.SubKeys.Any())
            {
                var emptySection = new AnalysisSection { Title = "📦 Packages" };
                emptySection.Items.Add(new AnalysisItem
                {
                    Name = "Status",
                    Value = "No packages found",
                    RegistryPath = packagesPath,
                    RegistryValue = "Component Based Servicing\\Packages not present or empty"
                });
                sections.Add(emptySection);
                return sections;
            }

            // Group packages by their base name (everything before the second tilde)
            var packageGroups = new Dictionary<string, List<PackageInfo>>(StringComparer.OrdinalIgnoreCase);

            foreach (var pkgKey in packagesKey.SubKeys)
            {
                var packageName = pkgKey.KeyName;
                var baseName = GetPackageBaseName(packageName);
                
                // Parse package info
                var pkgInfo = new PackageInfo
                {
                    FullName = packageName,
                    BaseName = baseName,
                    KeyPath = $"{packagesPath}\\{packageName}"
                };

                // Get CurrentState
                var currentStateValue = pkgKey.Values?.FirstOrDefault(v => 
                    v.ValueName?.Equals("CurrentState", StringComparison.OrdinalIgnoreCase) == true);
                if (currentStateValue != null)
                {
                    pkgInfo.CurrentStateRaw = ParseDwordValue(currentStateValue.ValueData);
                    pkgInfo.CurrentState = TranslatePackageState(pkgInfo.CurrentStateRaw);
                }

                // Get InstallUser SID
                var installUserValue = pkgKey.Values?.FirstOrDefault(v => 
                    v.ValueName?.Equals("InstallUser", StringComparison.OrdinalIgnoreCase) == true);
                if (installUserValue != null && !string.IsNullOrEmpty(installUserValue.ValueData))
                {
                    pkgInfo.InstallUserSid = installUserValue.ValueData;
                    pkgInfo.InstallUserName = TranslateSidToName(installUserValue.ValueData);
                }

                // Get InstallTimeHigh and InstallTimeLow (FILETIME)
                var installTimeHighValue = pkgKey.Values?.FirstOrDefault(v => 
                    v.ValueName?.Equals("InstallTimeHigh", StringComparison.OrdinalIgnoreCase) == true);
                var installTimeLowValue = pkgKey.Values?.FirstOrDefault(v => 
                    v.ValueName?.Equals("InstallTimeLow", StringComparison.OrdinalIgnoreCase) == true);
                
                if (installTimeHighValue != null && installTimeLowValue != null)
                {
                    var highPart = ParseDwordValue(installTimeHighValue.ValueData);
                    var lowPart = ParseDwordValue(installTimeLowValue.ValueData);
                    pkgInfo.InstallTime = FileTimeToDateTime(highPart, lowPart);
                }

                // Get InstallName and InstallLocation for additional info
                var installNameValue = pkgKey.Values?.FirstOrDefault(v => 
                    v.ValueName?.Equals("InstallName", StringComparison.OrdinalIgnoreCase) == true);
                if (installNameValue != null)
                {
                    pkgInfo.InstallName = installNameValue.ValueData;
                }

                var installLocationValue = pkgKey.Values?.FirstOrDefault(v => 
                    v.ValueName?.Equals("InstallLocation", StringComparison.OrdinalIgnoreCase) == true);
                if (installLocationValue != null)
                {
                    pkgInfo.InstallLocation = installLocationValue.ValueData;
                }

                // Get Visibility
                var visibilityValue = pkgKey.Values?.FirstOrDefault(v => 
                    v.ValueName?.Equals("Visibility", StringComparison.OrdinalIgnoreCase) == true);
                if (visibilityValue != null)
                {
                    pkgInfo.Visibility = ParseDwordValue(visibilityValue.ValueData);
                }

                // Add to group
                if (!packageGroups.ContainsKey(baseName))
                {
                    packageGroups[baseName] = new List<PackageInfo>();
                }
                packageGroups[baseName].Add(pkgInfo);
            }

            // Create sections for each package group
            // Sort groups alphabetically
            var sortedGroups = packageGroups.OrderBy(g => g.Key).ToList();

            // Summary section first
            var summarySection = new AnalysisSection { Title = "📊 Package Summary" };
            summarySection.Items.Add(new AnalysisItem
            {
                Name = "Total Package Groups",
                Value = packageGroups.Count.ToString("N0"),
                RegistryPath = packagesPath,
                RegistryValue = $"Unique package families: {packageGroups.Count}"
            });
            summarySection.Items.Add(new AnalysisItem
            {
                Name = "Total Packages",
                Value = packagesKey.SubKeys.Count.ToString("N0"),
                RegistryPath = packagesPath,
                RegistryValue = $"Total package entries: {packagesKey.SubKeys.Count}"
            });

            // Count by state
            var stateGroups = packageGroups.SelectMany(g => g.Value)
                .GroupBy(p => p.CurrentState)
                .OrderByDescending(g => g.Count())
                .ToList();
            
            foreach (var stateGroup in stateGroups)
            {
                summarySection.Items.Add(new AnalysisItem
                {
                    Name = stateGroup.Key,
                    Value = stateGroup.Count().ToString("N0"),
                    RegistryPath = packagesPath,
                    RegistryValue = $"Packages in {stateGroup.Key} state: {stateGroup.Count()}"
                });
            }
            sections.Add(summarySection);

            // Create a section for each package group
            foreach (var group in sortedGroups)
            {
                var groupSection = new AnalysisSection 
                { 
                    Title = $"📦 {group.Key} ({group.Value.Count})"
                };

                // Sort packages within group by install time (newest first), then by name
                var sortedPackages = group.Value
                    .OrderByDescending(p => p.InstallTime ?? DateTime.MinValue)
                    .ThenBy(p => p.FullName)
                    .ToList();

                foreach (var pkg in sortedPackages)
                {
                    // Build value string with key info
                    var valueBuilder = new System.Text.StringBuilder();
                    valueBuilder.Append($"State: {pkg.CurrentState}");
                    
                    if (pkg.InstallTime.HasValue)
                    {
                        valueBuilder.Append($" | Installed: {pkg.InstallTime.Value:yyyy-MM-dd HH:mm:ss}");
                    }
                    
                    if (!string.IsNullOrEmpty(pkg.InstallUserName))
                    {
                        valueBuilder.Append($" | User: {pkg.InstallUserName}");
                    }

                    valueBuilder.Append($" | Visibility: {pkg.Visibility}");

                    // Build registry value string with more details
                    var regValueBuilder = new System.Text.StringBuilder();
                    regValueBuilder.AppendLine($"Full Package Name: {pkg.FullName}");
                    regValueBuilder.AppendLine($"CurrentState: {pkg.CurrentStateRaw} (0x{pkg.CurrentStateRaw:X}) = {pkg.CurrentState}");
                    
                    if (pkg.InstallTime.HasValue)
                    {
                        regValueBuilder.AppendLine($"Install Time: {pkg.InstallTime.Value:yyyy-MM-dd HH:mm:ss}");
                    }
                    
                    if (!string.IsNullOrEmpty(pkg.InstallUserSid))
                    {
                        regValueBuilder.AppendLine($"InstallUser SID: {pkg.InstallUserSid}");
                        if (!string.IsNullOrEmpty(pkg.InstallUserName) && pkg.InstallUserName != pkg.InstallUserSid)
                        {
                            regValueBuilder.AppendLine($"InstallUser Name: {pkg.InstallUserName}");
                        }
                    }
                    
                    if (!string.IsNullOrEmpty(pkg.InstallName))
                    {
                        regValueBuilder.AppendLine($"InstallName: {pkg.InstallName}");
                    }
                    
                    if (!string.IsNullOrEmpty(pkg.InstallLocation))
                    {
                        regValueBuilder.AppendLine($"InstallLocation: {pkg.InstallLocation}");
                    }

                    if (pkg.Visibility > 0)
                    {
                        regValueBuilder.AppendLine($"Visibility: {pkg.Visibility}");
                    }

                    // Use shortened name for display (after base name)
                    var displayName = GetPackageDisplayName(pkg.FullName, pkg.BaseName);

                    groupSection.Items.Add(new AnalysisItem
                    {
                        Name = displayName,
                        Value = valueBuilder.ToString(),
                        RegistryPath = pkg.KeyPath,
                        RegistryValue = regValueBuilder.ToString().TrimEnd()
                    });
                }

                sections.Add(groupSection);
            }

            return sections;
        }

        /// <summary>
        /// Extract base name from package name (everything before the second tilde)
        /// Example: "Windows-Defender-Server-Service-WOW64-Package~31bf3856ad364e35~amd64~~10.0.17763.1"
        /// Returns: "Windows-Defender-Server-Service-WOW64-Package"
        /// </summary>
        private string GetPackageBaseName(string packageName)
        {
            if (string.IsNullOrEmpty(packageName))
                return packageName;

            var tildeIndex = packageName.IndexOf('~');
            if (tildeIndex > 0)
            {
                return packageName.Substring(0, tildeIndex);
            }
            return packageName;
        }

        /// <summary>
        /// Get display name for package (the version/variant part after the base name)
        /// </summary>
        private string GetPackageDisplayName(string fullName, string baseName)
        {
            if (string.IsNullOrEmpty(fullName))
                return fullName;

            if (!string.IsNullOrEmpty(baseName) && fullName.StartsWith(baseName, StringComparison.OrdinalIgnoreCase))
            {
                var remainder = fullName.Substring(baseName.Length);
                // Clean up leading tildes
                remainder = remainder.TrimStart('~');
                if (!string.IsNullOrEmpty(remainder))
                {
                    return remainder;
                }
            }
            return fullName;
        }

        /// <summary>
        /// Translate CurrentState DWORD value to human-readable string
        /// </summary>
        private string TranslatePackageState(int stateValue)
        {
            return stateValue switch
            {
                0 => "Absent",
                5 => "Uninstall Pending",
                16 => "Resolving",
                32 => "Resolved",
                48 => "Staging",
                64 => "Staged",
                80 => "Superseded",
                96 => "Install Pending",
                101 => "Partially Installed",
                112 => "Installed",
                128 => "Permanent",
                _ => $"Unknown ({stateValue})"
            };
        }

        /// <summary>
        /// Parse DWORD value from string (handles hex and decimal)
        /// </summary>
        private int ParseDwordValue(string? value)
        {
            if (string.IsNullOrEmpty(value))
                return 0;

            // Handle format like "112 (0x70)" 
            var spaceIndex = value.IndexOf(' ');
            if (spaceIndex > 0)
            {
                value = value.Substring(0, spaceIndex);
            }

            // Handle hex format
            if (value.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
            {
                if (int.TryParse(value.Substring(2), System.Globalization.NumberStyles.HexNumber, null, out int hexResult))
                    return hexResult;
            }

            // Try decimal
            if (int.TryParse(value, out int result))
                return result;

            return 0;
        }


        /// <summary>
        /// Convert FILETIME (high/low parts) to DateTime
        /// </summary>
        private DateTime? FileTimeToDateTime(int highPart, int lowPart)
        {
            try
            {
                // Combine high and low parts into a 64-bit value
                long fileTime = ((long)(uint)highPart << 32) | (uint)lowPart;
                
                if (fileTime <= 0)
                    return null;

                // Convert FILETIME to DateTime
                return DateTime.FromFileTimeUtc(fileTime);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Translate well-known SIDs to names
        /// </summary>
        private string TranslateSidToName(string sid)
        {
            if (string.IsNullOrEmpty(sid))
                return sid;

            // Well-known SIDs
            var wellKnownSids = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "S-1-5-18", "SYSTEM (Local System)" },
                { "S-1-5-19", "LOCAL SERVICE" },
                { "S-1-5-20", "NETWORK SERVICE" },
                { "S-1-5-32-544", "Administrators" },
                { "S-1-5-32-545", "Users" },
                { "S-1-5-32-546", "Guests" },
                { "S-1-5-32-547", "Power Users" },
                { "S-1-1-0", "Everyone" },
                { "S-1-5-11", "Authenticated Users" },
                { "S-1-5-4", "Interactive" },
                { "S-1-5-6", "Service" },
                { "S-1-5-7", "Anonymous" },
                { "S-1-5-9", "Enterprise Domain Controllers" },
                { "S-1-5-10", "Principal Self" },
            };

            if (wellKnownSids.TryGetValue(sid, out var name))
            {
                return name;
            }

            // Check if it's a user SID (S-1-5-21-...)
            if (sid.StartsWith("S-1-5-21-", StringComparison.OrdinalIgnoreCase))
            {
                // Try to get RID (last part)
                var parts = sid.Split('-');
                if (parts.Length >= 8)
                {
                    var rid = parts[^1];
                    
                    // Well-known RIDs
                    var ridName = rid switch
                    {
                        "500" => "Administrator",
                        "501" => "Guest",
                        "502" => "KRBTGT",
                        "512" => "Domain Admins",
                        "513" => "Domain Users",
                        "514" => "Domain Guests",
                        "515" => "Domain Computers",
                        "516" => "Domain Controllers",
                        "517" => "Cert Publishers",
                        "518" => "Schema Admins",
                        "519" => "Enterprise Admins",
                        "520" => "Group Policy Creator Owners",
                        _ => null
                    };

                    if (ridName != null)
                    {
                        return $"{ridName} ({sid})";
                    }
                }

                // Return SID with hint that it's a domain/local user
                return $"User ({sid})";
            }

            return sid;
        }

        /// <summary>
        /// Package information container
        /// </summary>
        private class PackageInfo
        {
            public string FullName { get; set; } = "";
            public string BaseName { get; set; } = "";
            public string KeyPath { get; set; } = "";
            public int CurrentStateRaw { get; set; }
            public string CurrentState { get; set; } = "Unknown";
            public string? InstallUserSid { get; set; }
            public string? InstallUserName { get; set; }
            public DateTime? InstallTime { get; set; }
            public string? InstallName { get; set; }
            public string? InstallLocation { get; set; }
            public int Visibility { get; set; }
        }

        /// <summary>
        /// Get Component Store information from COMPONENTS hive
        /// </summary>
        public List<AnalysisSection> GetComponentStoreAnalysis()
        {
            var sections = new List<AnalysisSection>();

            // Check if this is COMPONENTS hive
            if (_parser.CurrentHiveType != OfflineRegistryParser.HiveType.COMPONENTS)
            {
                var noticeSection = new AnalysisSection { Title = "Notice" };
                noticeSection.Items.Add(new AnalysisItem
                {
                    Name = "Info",
                    Value = "This information requires the COMPONENTS hive to be loaded",
                    RegistryPath = "",
                    RegistryValue = "Load COMPONENTS hive (C:\\Windows\\System32\\config\\COMPONENTS) for component store information"
                });
                sections.Add(noticeSection);
                return sections;
            }

            // Summary section
            var summarySection = new AnalysisSection { Title = "Component Store Summary" };
            
            // Try to get components from DerivedData\Components
            var componentsPath = @"DerivedData\Components";
            var componentsKey = _parser.GetKey(componentsPath);

            int totalComponents = 0;
            var componentCategories = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            if (componentsKey?.SubKeys != null)
            {
                totalComponents = componentsKey.SubKeys.Count;

                // Categorize components by their prefix
                foreach (var compKey in componentsKey.SubKeys)
                {
                    var compName = compKey.KeyName;
                    var category = GetComponentCategory(compName);
                    
                    if (!componentCategories.ContainsKey(category))
                        componentCategories[category] = 0;
                    componentCategories[category]++;
                }
            }

            summarySection.Items.Add(new AnalysisItem
            {
                Name = "Total Components",
                Value = totalComponents.ToString("N0"),
                RegistryPath = componentsPath,
                RegistryValue = $"Total component entries in Component Store: {totalComponents}"
            });

            // Add categories breakdown
            foreach (var cat in componentCategories.OrderByDescending(c => c.Value).Take(10))
            {
                summarySection.Items.Add(new AnalysisItem
                {
                    Name = cat.Key,
                    Value = cat.Value.ToString("N0"),
                    RegistryPath = componentsPath,
                    RegistryValue = $"{cat.Key}: {cat.Value} components"
                });
            }

            sections.Add(summarySection);

            // Try to get deployments
            var deploymentsPath = @"DerivedData\CanonicalData\Deployments";
            var deploymentsKey = _parser.GetKey(deploymentsPath);

            if (deploymentsKey?.SubKeys != null && deploymentsKey.SubKeys.Any())
            {
                var deploymentsSection = new AnalysisSection { Title = "Deployments" };
                
                int deploymentCount = 0;
                foreach (var deployKey in deploymentsKey.SubKeys.Take(100)) // Limit to first 100
                {
                    var deployName = deployKey.KeyName;
                    var appIdValue = deployKey.Values?.FirstOrDefault(v => 
                        v.ValueName?.Equals("appid", StringComparison.OrdinalIgnoreCase) == true);
                    var appId = appIdValue?.ValueData ?? "";

                    deploymentsSection.Items.Add(new AnalysisItem
                    {
                        Name = TruncateString(deployName, 80),
                        Value = !string.IsNullOrEmpty(appId) ? appId : "Active",
                        RegistryPath = $"{deploymentsPath}\\{deployName}",
                        RegistryValue = $"Deployment: {deployName}\nAppId: {appId}"
                    });
                    deploymentCount++;
                }

                if (deploymentCount > 0)
                {
                    sections.Add(deploymentsSection);
                }
            }

            // Try to get package index from Servicing\PackageIndex
            var packageIndexPath = @"Servicing\PackageIndex";
            var packageIndexKey = _parser.GetKey(packageIndexPath);

            if (packageIndexKey?.SubKeys != null && packageIndexKey.SubKeys.Any())
            {
                var packageIndexSection = new AnalysisSection { Title = "Package Index" };
                
                int pkgCount = 0;
                foreach (var pkgKey in packageIndexKey.SubKeys.Take(100)) // Limit to first 100
                {
                    var pkgName = pkgKey.KeyName;
                    var subKeyCount = pkgKey.SubKeys?.Count ?? 0;

                    packageIndexSection.Items.Add(new AnalysisItem
                    {
                        Name = TruncateString(pkgName, 80),
                        Value = subKeyCount > 0 ? $"{subKeyCount} version(s)" : "Indexed",
                        RegistryPath = $"{packageIndexPath}\\{pkgName}",
                        RegistryValue = $"Package: {pkgName}\nVersions: {subKeyCount}"
                    });
                    pkgCount++;
                }

                if (pkgCount > 0)
                {
                    // Add summary at the beginning
                    packageIndexSection.Items.Insert(0, new AnalysisItem
                    {
                        Name = "Total Indexed Packages",
                        Value = packageIndexKey.SubKeys.Count.ToString("N0"),
                        RegistryPath = packageIndexPath,
                        RegistryValue = $"Total packages in index: {packageIndexKey.SubKeys.Count}"
                    });
                    sections.Add(packageIndexSection);
                }
            }

            // If no data found
            if (sections.Count == 1 && summarySection.Items.Count == 1 && totalComponents == 0)
            {
                summarySection.Items.Add(new AnalysisItem
                {
                    Name = "Status",
                    Value = "Component Store appears to be empty or has an unexpected structure",
                    RegistryPath = "",
                    RegistryValue = "The COMPONENTS hive was loaded but no component data was found in expected locations"
                });
            }

            return sections;
        }

        /// <summary>
        /// Get all components from COMPONENTS hive as a flat list
        /// Returns each individual component with full key name
        /// </summary>
        public List<AnalysisItem> GetAllComponentsList()
        {
            var items = new List<AnalysisItem>();

            // Check if this is COMPONENTS hive
            if (_parser.CurrentHiveType != OfflineRegistryParser.HiveType.COMPONENTS)
            {
                return items;
            }

            var componentsPath = @"DerivedData\Components";
            var componentsKey = _parser.GetKey(componentsPath);

            if (componentsKey?.SubKeys == null)
            {
                return items;
            }

            // Return each component with its full key name
            foreach (var compKey in componentsKey.SubKeys)
            {
                var compName = compKey.KeyName;
                var fullPath = $@"{componentsPath}\{compName}";
                
                // Build simple list of values: Name; Type; Data
                var detailBuilder = new System.Text.StringBuilder();
                
                // Show registry values - simple format: Name; Type; Data
                if (compKey.Values != null && compKey.Values.Any())
                {
                    foreach (var val in compKey.Values)
                    {
                        var valName = val.ValueName ?? "(Default)";
                        var valType = val.ValueType?.ToString() ?? "Unknown";
                        var valData = val.ValueData?.ToString() ?? "";
                        
                        detailBuilder.AppendLine($"{valName}; {valType}; {valData}");
                    }
                }
                else
                {
                    detailBuilder.AppendLine("(No values)");
                }
                
                // Show subkeys count at the end if any
                if (compKey.SubKeys != null && compKey.SubKeys.Any())
                {
                    detailBuilder.AppendLine();
                    detailBuilder.AppendLine($"Subkeys: {compKey.SubKeys.Count}");
                }

                items.Add(new AnalysisItem
                {
                    Name = compName,
                    Value = "", // Not used for single column display
                    RegistryPath = fullPath,
                    RegistryValue = detailBuilder.ToString()
                });
            }

            return items;
        }

        /// <summary>
        /// Scan component identities for invalid (non-ASCII) characters
        /// Based on PowerShell script logic: checks if any byte in "identity" value is outside 0x00-0x7F range
        /// </summary>
        public List<AnalysisItem> GetComponentIdentitiesAnalysis()
        {
            var items = new List<AnalysisItem>();

            // Check if this is COMPONENTS hive
            if (_parser.CurrentHiveType != OfflineRegistryParser.HiveType.COMPONENTS)
            {
                return items;
            }

            var componentsPath = @"DerivedData\Components";
            var componentsKey = _parser.GetKey(componentsPath);

            if (componentsKey?.SubKeys == null)
            {
                return items;
            }

            int totalScanned = 0;
            int invalidFound = 0;

            foreach (var compKey in componentsKey.SubKeys)
            {
                var compName = compKey.KeyName;
                var fullPath = $@"{componentsPath}\{compName}";
                totalScanned++;

                // Look for "identity" value (case-insensitive)
                var identityValue = compKey.Values?.FirstOrDefault(v => 
                    v.ValueName?.Equals("identity", StringComparison.OrdinalIgnoreCase) == true);

                if (identityValue == null)
                    continue;

                // Get the raw binary data using ValueDataRaw
                byte[]? identityData = identityValue.ValueDataRaw;
                
                if (identityData == null || identityData.Length == 0)
                {
                    // Fallback: try to convert string representation to bytes
                    var strData = identityValue.ValueData?.ToString();
                    if (!string.IsNullOrEmpty(strData))
                    {
                        identityData = System.Text.Encoding.Unicode.GetBytes(strData);
                    }
                }

                if (identityData == null || identityData.Length == 0)
                    continue;

                // Check for invalid characters (outside ASCII range 0x00-0x7F)
                var invalidChars = new List<(int position, byte value, char character)>();
                for (int i = 0; i < identityData.Length; i++)
                {
                    byte b = identityData[i];
                    if (b > 0x7F) // Outside ASCII range
                    {
                        char c = (char)b;
                        invalidChars.Add((i, b, c));
                    }
                }

                if (invalidChars.Count > 0)
                {
                    invalidFound++;
                    
                    var detailBuilder = new System.Text.StringBuilder();
                    detailBuilder.AppendLine($"Component: {compName}");
                    detailBuilder.AppendLine($"Registry Path: {fullPath}");
                    detailBuilder.AppendLine();
                    detailBuilder.AppendLine($"INVALID CHARACTERS FOUND: {invalidChars.Count}");
                    detailBuilder.AppendLine();
                    
                    foreach (var (pos, val, chr) in invalidChars.Take(20)) // Limit to first 20
                    {
                        detailBuilder.AppendLine($"  Position {pos}: 0x{val:X2} ('{chr}')");
                    }
                    
                    if (invalidChars.Count > 20)
                    {
                        detailBuilder.AppendLine($"  ... and {invalidChars.Count - 20} more invalid characters");
                    }

                    detailBuilder.AppendLine();
                    detailBuilder.AppendLine("Identity data may be corrupted. Non-ASCII characters in component");
                    detailBuilder.AppendLine("identity values can cause Windows servicing issues.");

                    items.Add(new AnalysisItem
                    {
                        Name = compName,
                        Value = $"{invalidChars.Count} invalid char(s)",
                        RegistryPath = fullPath,
                        RegistryValue = detailBuilder.ToString()
                    });
                }
            }

            // Add summary at the beginning
            var summaryBuilder = new System.Text.StringBuilder();
            summaryBuilder.AppendLine($"Identity Scan Results");
            summaryBuilder.AppendLine();
            summaryBuilder.AppendLine($"Total components scanned: {totalScanned:N0}");
            summaryBuilder.AppendLine($"Components with invalid characters: {invalidFound:N0}");
            summaryBuilder.AppendLine();
            if (invalidFound == 0)
            {
                summaryBuilder.AppendLine("No invalid characters found in any component identity values.");
                summaryBuilder.AppendLine("The Component Store appears healthy.");
            }
            else
            {
                summaryBuilder.AppendLine($"WARNING: {invalidFound} component(s) have invalid characters in their identity values.");
                summaryBuilder.AppendLine("This may indicate corruption in the Component Store.");
                summaryBuilder.AppendLine("Invalid characters are bytes outside the ASCII range (0x00-0x7F).");
            }

            items.Insert(0, new AnalysisItem
            {
                Name = "Scan Summary",
                Value = invalidFound == 0 ? "No issues found" : $"{invalidFound} issue(s) found",
                RegistryPath = componentsPath,
                RegistryValue = summaryBuilder.ToString()
            });

            return items;
        }

        /// <summary>
        /// Extract component category from component name
        /// </summary>
        private string GetComponentCategory(string componentName)
        {
            if (string.IsNullOrEmpty(componentName))
                return "Other";

            // Try to extract the main category from component name
            // Component names often follow patterns like:
            // amd64_microsoft-windows-something_31bf3856ad364e35_...
            // x86_microsoft-windows-something_31bf3856ad364e35_...
            // wow64_microsoft-windows-something_31bf3856ad364e35_...
            
            var parts = componentName.Split('_');
            if (parts.Length >= 2)
            {
                var category = parts[1];
                
                // Clean up common prefixes
                if (category.StartsWith("microsoft-windows-", StringComparison.OrdinalIgnoreCase))
                {
                    category = category.Substring(18);
                    var dashIndex = category.IndexOf('-');
                    if (dashIndex > 0)
                        category = category.Substring(0, dashIndex);
                    return $"Windows {category}";
                }
                if (category.StartsWith("microsoft-", StringComparison.OrdinalIgnoreCase))
                {
                    return category.Substring(10);
                }
                
                return category;
            }

            return "Other";
        }

        /// <summary>
        /// Truncate string with ellipsis
        /// </summary>
        private string TruncateString(string str, int maxLength)
        {
            return TruncatePath(str, maxLength);
        }

        /// <summary>
        /// Get CBS Pending Sessions information (requires SOFTWARE hive)
        /// Shows sessions pending reboot
        /// </summary>
        public List<AnalysisSection> GetCbsPendingSessionsAnalysis()
        {
            var sections = new List<AnalysisSection>();

            // Check if this is SOFTWARE hive
            if (_parser.CurrentHiveType != OfflineRegistryParser.HiveType.SOFTWARE)
            {
                var noticeSection = new AnalysisSection { Title = "Notice" };
                noticeSection.Items.Add(new AnalysisItem
                {
                    Name = "Info",
                    Value = "This information requires the SOFTWARE hive to be loaded",
                    RegistryPath = "",
                    RegistryValue = "Load SOFTWARE hive for pending sessions information"
                });
                sections.Add(noticeSection);
                return sections;
            }

            var sessionsPath = @"Microsoft\Windows\CurrentVersion\Component Based Servicing\SessionsPending";
            var sessionsKey = _parser.GetKey(sessionsPath);

            var sessionsSection = new AnalysisSection { Title = "Pending Sessions" };

            if (sessionsKey?.SubKeys != null && sessionsKey.SubKeys.Any())
            {
                sessionsSection.Items.Add(new AnalysisItem
                {
                    Name = "Total Pending Sessions",
                    Value = sessionsKey.SubKeys.Count.ToString(),
                    RegistryPath = sessionsPath,
                    RegistryValue = $"Number of sessions awaiting reboot: {sessionsKey.SubKeys.Count}"
                });

                foreach (var sessionKey in sessionsKey.SubKeys.Take(50)) // Limit to 50
                {
                    var sessionId = sessionKey.KeyName;
                    var clientValue = sessionKey.Values?.FirstOrDefault(v => 
                        v.ValueName?.Equals("Client", StringComparison.OrdinalIgnoreCase) == true);
                    var statusValue = sessionKey.Values?.FirstOrDefault(v => 
                        v.ValueName?.Equals("Status", StringComparison.OrdinalIgnoreCase) == true);
                    
                    var client = clientValue?.ValueData ?? "Unknown";
                    var status = statusValue?.ValueData ?? "Pending";

                    // Build detailed info
                    var detailBuilder = new System.Text.StringBuilder();
                    detailBuilder.AppendLine($"Session ID: {sessionId}");
                    detailBuilder.AppendLine($"Client: {client}");
                    detailBuilder.AppendLine($"Status: {status}");
                    
                    // Check for operations in the session
                    if (sessionKey.SubKeys != null)
                    {
                        detailBuilder.AppendLine($"Operations: {sessionKey.SubKeys.Count}");
                    }

                    sessionsSection.Items.Add(new AnalysisItem
                    {
                        Name = sessionId,
                        Value = $"Client: {client}",
                        RegistryPath = $"{sessionsPath}\\{sessionId}",
                        RegistryValue = detailBuilder.ToString().TrimEnd()
                    });
                }
            }
            else
            {
                sessionsSection.Items.Add(new AnalysisItem
                {
                    Name = "Status",
                    Value = "No pending sessions",
                    RegistryPath = sessionsPath,
                    RegistryValue = "No sessions are pending - no reboot required for CBS operations"
                });
            }

            sections.Add(sessionsSection);
            return sections;
        }

        /// <summary>
        /// Get CBS Pending Packages information (requires SOFTWARE hive)
        /// Shows packages pending installation/removal
        /// </summary>
        public List<AnalysisSection> GetCbsPendingPackagesAnalysis()
        {
            var sections = new List<AnalysisSection>();

            // Check if this is SOFTWARE hive
            if (_parser.CurrentHiveType != OfflineRegistryParser.HiveType.SOFTWARE)
            {
                var noticeSection = new AnalysisSection { Title = "Notice" };
                noticeSection.Items.Add(new AnalysisItem
                {
                    Name = "Info",
                    Value = "This information requires the SOFTWARE hive to be loaded",
                    RegistryPath = "",
                    RegistryValue = "Load SOFTWARE hive for pending packages information"
                });
                sections.Add(noticeSection);
                return sections;
            }

            var packagesPath = @"Microsoft\Windows\CurrentVersion\Component Based Servicing\PackagesPending";
            var packagesKey = _parser.GetKey(packagesPath);

            var packagesSection = new AnalysisSection { Title = "Pending Packages" };

            if (packagesKey?.SubKeys != null && packagesKey.SubKeys.Any())
            {
                packagesSection.Items.Add(new AnalysisItem
                {
                    Name = "Total Pending Packages",
                    Value = packagesKey.SubKeys.Count.ToString(),
                    RegistryPath = packagesPath,
                    RegistryValue = $"Number of packages pending: {packagesKey.SubKeys.Count}"
                });

                foreach (var pkgKey in packagesKey.SubKeys.Take(100)) // Limit to 100
                {
                    var pkgName = pkgKey.KeyName;
                    
                    // Get pending state
                    var pendingStateValue = pkgKey.Values?.FirstOrDefault(v => 
                        v.ValueName?.Equals("PendingState", StringComparison.OrdinalIgnoreCase) == true);
                    var pendingState = pendingStateValue != null 
                        ? TranslatePackageState(ParseDwordValue(pendingStateValue.ValueData))
                        : "Pending";

                    // Build detailed info
                    var detailBuilder = new System.Text.StringBuilder();
                    detailBuilder.AppendLine($"Package: {pkgName}");
                    detailBuilder.AppendLine($"Pending State: {pendingState}");
                    
                    // Add other values
                    if (pkgKey.Values != null)
                    {
                        foreach (var val in pkgKey.Values.Take(10))
                        {
                            if (val.ValueName != "PendingState")
                            {
                                detailBuilder.AppendLine($"{val.ValueName}: {val.ValueData}");
                            }
                        }
                    }

                    packagesSection.Items.Add(new AnalysisItem
                    {
                        Name = TruncateString(pkgName, 60),
                        Value = pendingState,
                        RegistryPath = $"{packagesPath}\\{pkgName}",
                        RegistryValue = detailBuilder.ToString().TrimEnd()
                    });
                }
            }
            else
            {
                packagesSection.Items.Add(new AnalysisItem
                {
                    Name = "Status",
                    Value = "No pending packages",
                    RegistryPath = packagesPath,
                    RegistryValue = "No packages are pending installation or removal"
                });
            }

            sections.Add(packagesSection);
            return sections;
        }

        /// <summary>
        /// Get CBS Reboot Status information (requires SOFTWARE hive)
        /// Shows comprehensive reboot requirements including ServicingInProgress state
        /// </summary>
        public List<AnalysisSection> GetCbsRebootStatusAnalysis()
        {
            var sections = new List<AnalysisSection>();

            // Check if this is SOFTWARE hive
            if (_parser.CurrentHiveType != OfflineRegistryParser.HiveType.SOFTWARE)
            {
                var noticeSection = new AnalysisSection { Title = "Notice" };
                noticeSection.Items.Add(new AnalysisItem
                {
                    Name = "Info",
                    Value = "This information requires the SOFTWARE hive to be loaded",
                    RegistryPath = "",
                    RegistryValue = "Load SOFTWARE hive for reboot status information"
                });
                sections.Add(noticeSection);
                return sections;
            }

            var cbsPath = @"Microsoft\Windows\CurrentVersion\Component Based Servicing";
            var interfacePath = $@"{cbsPath}\Interface";
            var cbsKey = _parser.GetKey(cbsPath);

            var statusSection = new AnalysisSection { Title = "Servicing Status" };

            // Check ServicingInProgress from Interface subkey (most important indicator)
            var servicingInProgressValue = GetValue(interfacePath, "ServicingInProgress");
            var servicingState = 0;
            if (!string.IsNullOrEmpty(servicingInProgressValue))
            {
                servicingState = ParseDwordValue(servicingInProgressValue);
            }

            var (servicingStatus, servicingDescription) = servicingState switch
            {
                0 => ("Ready", "No servicing operation has required a reboot. System is ready."),
                1 => ("Pending Reboot", "A servicing operation is pending, waiting for a reboot. TrustedInstaller will not initiate the reboot. The user or client is expected to reboot."),
                2 => ("Executing", "A servicing operation is currently executing. TrustedInstaller can and will reboot the computer as needed."),
                _ => ($"Unknown ({servicingState})", $"Unknown servicing state value: {servicingState}")
            };

            statusSection.Items.Add(new AnalysisItem
            {
                Name = "Servicing State",
                Value = servicingStatus,
                RegistryPath = interfacePath,
                RegistryValue = $"ServicingInProgress = {servicingState}\n{servicingDescription}"
            });

            // Check RebootPending value
            var rebootPending = GetValue(cbsPath, "RebootPending");
            var rebootRequired = !string.IsNullOrEmpty(rebootPending) && rebootPending != "0";
            
            statusSection.Items.Add(new AnalysisItem
            {
                Name = "Reboot Required (CBS)",
                Value = rebootRequired ? "Yes" : "No",
                RegistryPath = cbsPath,
                RegistryValue = rebootRequired 
                    ? $"RebootPending = {rebootPending}\nA reboot is required to complete pending CBS operations" 
                    : "RebootPending flag is not set or is 0"
            });

            // Check RebootInProgress
            var rebootInProgress = GetValue(cbsPath, "RebootInProgress");
            if (!string.IsNullOrEmpty(rebootInProgress))
            {
                statusSection.Items.Add(new AnalysisItem
                {
                    Name = "Reboot In Progress",
                    Value = rebootInProgress == "1" ? "Yes" : "No",
                    RegistryPath = cbsPath,
                    RegistryValue = $"RebootInProgress = {rebootInProgress}\n" +
                        (rebootInProgress == "1" ? "A reboot was initiated and may be in progress" : "No reboot is currently in progress")
                });
            }

            // Count pending sessions
            var sessionsPath = $@"{cbsPath}\SessionsPending";
            var sessionsKey = _parser.GetKey(sessionsPath);
            var pendingSessionsCount = sessionsKey?.SubKeys?.Count ?? 0;
            
            statusSection.Items.Add(new AnalysisItem
            {
                Name = "Pending Sessions",
                Value = pendingSessionsCount.ToString(),
                RegistryPath = sessionsPath,
                RegistryValue = pendingSessionsCount > 0 
                    ? $"{pendingSessionsCount} session(s) waiting to complete after reboot.\nThese are CBS operations that require a system restart to finalize."
                    : "No pending sessions. All CBS sessions have completed."
            });

            // Count pending packages
            var packagesPath = $@"{cbsPath}\PackagesPending";
            var packagesKey = _parser.GetKey(packagesPath);
            var pendingPackagesCount = packagesKey?.SubKeys?.Count ?? 0;
            
            statusSection.Items.Add(new AnalysisItem
            {
                Name = "Pending Packages",
                Value = pendingPackagesCount.ToString(),
                RegistryPath = packagesPath,
                RegistryValue = pendingPackagesCount > 0 
                    ? $"{pendingPackagesCount} package(s) pending installation or removal.\nThese packages are queued and will be processed on next reboot."
                    : "No pending packages. All package operations have completed."
            });

            // Check for other reboot indicators
            var sessionsPendingExclusive = GetValue(cbsPath, "SessionsPendingExclusive");
            if (!string.IsNullOrEmpty(sessionsPendingExclusive))
            {
                statusSection.Items.Add(new AnalysisItem
                {
                    Name = "Exclusive Sessions Pending",
                    Value = sessionsPendingExclusive,
                    RegistryPath = cbsPath,
                    RegistryValue = $"SessionsPendingExclusive = {sessionsPendingExclusive}\nExclusive sessions require sole access to the servicing stack during reboot."
                });
            }

            // Check LastSuccessfulTrust scan
            var lastTrustTime = GetValue(cbsPath, "LastTrustTime");
            if (!string.IsNullOrEmpty(lastTrustTime))
            {
                statusSection.Items.Add(new AnalysisItem
                {
                    Name = "Last Trust Verification",
                    Value = lastTrustTime,
                    RegistryPath = cbsPath,
                    RegistryValue = $"LastTrustTime = {lastTrustTime}\nTimestamp of the last component trust verification."
                });
            }

            sections.Add(statusSection);
            return sections;
        }

        #endregion

        #region Activation Analysis (SOFTWARE hive)

        /// <summary>
        /// Get Windows activation and KMS information (requires SOFTWARE hive)
        /// </summary>
        public AnalysisSection GetActivationAnalysis()
        {
            var section = new AnalysisSection { Title = "🔑 Windows Activation" };

            // Check if this is SOFTWARE hive
            if (_parser.CurrentHiveType != OfflineRegistryParser.HiveType.SOFTWARE)
            {
                section.Items.Add(new AnalysisItem
                {
                    Name = "Notice",
                    Value = "This information requires the SOFTWARE hive to be loaded",
                    RegistryPath = "",
                    RegistryValue = "Load SOFTWARE hive to view activation information"
                });
                return section;
            }

            // Windows Licensing paths
            var sppPath = @"Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform";
            var sppKey = _parser.GetKey(sppPath);

            // KMS Server settings
            var kmsServer = GetValue(sppPath, "KeyManagementServiceName");
            if (!string.IsNullOrEmpty(kmsServer))
            {
                section.Items.Add(new AnalysisItem
                {
                    Name = "KeyManagementServiceName",
                    Value = kmsServer,
                    RegistryPath = sppPath,
                    RegistryValue = $"KeyManagementServiceName = {kmsServer}"
                });
            }

            var kmsPort = GetValue(sppPath, "KeyManagementServicePort");
            if (!string.IsNullOrEmpty(kmsPort))
            {
                section.Items.Add(new AnalysisItem
                {
                    Name = "KeyManagementServicePort",
                    Value = kmsPort,
                    RegistryPath = sppPath,
                    RegistryValue = $"KeyManagementServicePort = {kmsPort}"
                });
            }
            else if (!string.IsNullOrEmpty(kmsServer))
            {
                section.Items.Add(new AnalysisItem
                {
                    Name = "KeyManagementServicePort",
                    Value = "1688 (default)",
                    RegistryPath = sppPath,
                    RegistryValue = "KeyManagementServicePort not set, using default 1688"
                });
            }

            // Disable KMS DNS Auto-Discovery
            var disableAuto = GetValue(sppPath, "DisableDnsPublishing");
            if (!string.IsNullOrEmpty(disableAuto))
            {
                section.Items.Add(new AnalysisItem
                {
                    Name = "DisableDnsPublishing",
                    Value = disableAuto == "1" ? "Yes (1)" : "No (0)",
                    RegistryPath = sppPath,
                    RegistryValue = $"DisableDnsPublishing = {disableAuto}"
                });
            }

            // Activation Override
            var activationDisabled = GetValue(sppPath, "ActivationDisabled");
            if (!string.IsNullOrEmpty(activationDisabled))
            {
                section.Items.Add(new AnalysisItem
                {
                    Name = "ActivationDisabled",
                    Value = activationDisabled == "1" ? "Yes (1)" : "No (0)",
                    RegistryPath = sppPath,
                    RegistryValue = $"ActivationDisabled = {activationDisabled}"
                });
            }

            // Look for Office KMS settings
            var officeKmsPath = @"Microsoft\Office\ClickToRun\Configuration";
            var officeKmsServer = GetValue(officeKmsPath, "KeyManagementServiceName");
            if (!string.IsNullOrEmpty(officeKmsServer))
            {
                section.Items.Add(new AnalysisItem
                {
                    Name = "Office KeyManagementServiceName",
                    Value = officeKmsServer,
                    RegistryPath = officeKmsPath,
                    RegistryValue = $"KeyManagementServiceName = {officeKmsServer}"
                });
            }

            // Product Key backup path (Windows 8+)
            var backupProductKeyPath = @"Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform";
            var backupKey = GetValue(backupProductKeyPath, "BackupProductKeyDefault");
            if (!string.IsNullOrEmpty(backupKey))
            {
                section.Items.Add(new AnalysisItem
                {
                    Name = "BackupProductKeyDefault",
                    Value = backupKey,
                    RegistryPath = backupProductKeyPath,
                    RegistryValue = $"BackupProductKeyDefault = {backupKey}"
                });
            }

            // Digital Product ID
            var productIdPath = @"Microsoft\Windows NT\CurrentVersion";
            var digitalProductId = GetBinaryValue(productIdPath, "DigitalProductId");
            if (digitalProductId != null && digitalProductId.Length > 0)
            {
                section.Items.Add(new AnalysisItem
                {
                    Name = "DigitalProductId",
                    Value = $"Present ({digitalProductId.Length} bytes)",
                    RegistryPath = productIdPath,
                    RegistryValue = "DigitalProductId binary data present"
                });
            }

            // Tokens path for licensing
            var tokensPath = @"Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform";
            var tokensDat = GetValue(tokensPath, "TokenStore");
            if (!string.IsNullOrEmpty(tokensDat))
            {
                section.Items.Add(new AnalysisItem
                {
                    Name = "TokenStore",
                    Value = tokensDat,
                    RegistryPath = tokensPath,
                    RegistryValue = $"TokenStore = {tokensDat}"
                });
            }

            // Look at SLP/OEM activation markers
            var slpPath = @"Microsoft\Windows\CurrentVersion\OEMInformation";
            var oemManufacturer = GetValue(slpPath, "Manufacturer");
            var oemModel = GetValue(slpPath, "Model");
            if (!string.IsNullOrEmpty(oemManufacturer))
            {
                section.Items.Add(new AnalysisItem
                {
                    Name = "Manufacturer",
                    Value = oemManufacturer,
                    RegistryPath = slpPath,
                    RegistryValue = $"Manufacturer = {oemManufacturer}"
                });
            }
            if (!string.IsNullOrEmpty(oemModel))
            {
                section.Items.Add(new AnalysisItem
                {
                    Name = "Model",
                    Value = oemModel,
                    RegistryPath = slpPath,
                    RegistryValue = $"Model = {oemModel}"
                });
            }

            // If no activation info found
            if (section.Items.Count == 0)
            {
                section.Items.Add(new AnalysisItem
                {
                    Name = "Status",
                    Value = "No KMS or special activation settings found",
                    RegistryPath = sppPath,
                    RegistryValue = "Standard retail/OEM activation assumed"
                });
            }

            return section;
        }

        #endregion

        #region Storage Analysis (SYSTEM hive)

        /// <summary>
        /// Get storage-related information as structured sections
        /// </summary>
        public List<AnalysisSection> GetStorageAnalysis()
        {
            var sections = new List<AnalysisSection>();

            // Filters section (placeholder - actual data loaded via GetDiskFilters/GetVolumeFilters)
            var filtersSection = new AnalysisSection { Title = "🔧 Filters" };
            filtersSection.Items.Add(new AnalysisItem
            {
                Name = "Select a filter type",
                Value = "Use the filter buttons above to view Disk Filters or Volume Filters"
            });
            sections.Add(filtersSection);

            // Mounted Devices section (placeholder - actual data loaded via GetMountedDevices)
            var partitionSection = new AnalysisSection { Title = "💿 Mounted Devices" };
            partitionSection.Items.Add(new AnalysisItem
            {
                Name = "Mounted Devices",
                Value = "Displays MBR signatures, GPT partition GUIDs, and device paths from MountedDevices"
            });
            sections.Add(partitionSection);

            // Physical Disks section (placeholder - actual data loaded via GetPhysicalDisks)
            var physicalDisksSection = new AnalysisSection { Title = "💽 Physical Disks" };
            physicalDisksSection.Items.Add(new AnalysisItem
            {
                Name = "Physical Disks",
                Value = "Enumerates all physical disk devices and detects probable Storage Spaces pool members"
            });
            sections.Add(physicalDisksSection);

            return sections;
        }

        /// <summary>
        /// Get disk class filter drivers configuration
        /// </summary>
        public List<AnalysisItem> GetDiskFilters()
        {
            return GetClassFilters("{4d36e967-e325-11ce-bfc1-08002be10318}", "Disk");
        }

        /// <summary>
        /// Get volume class filter drivers configuration
        /// </summary>
        public List<AnalysisItem> GetVolumeFilters()
        {
            return GetClassFilters("{71A27CDD-812A-11D0-BEC7-08002BE2092F}", "Volume");
        }

        /// <summary>
        /// Get filter drivers configuration for a given device class GUID
        /// </summary>
        private List<AnalysisItem> GetClassFilters(string classGuid, string className)
        {
            var items = new List<AnalysisItem>();

            var classPath = $@"{_currentControlSet}\Control\Class\{classGuid}";
            var classKey = _parser.GetKey(classPath);

            if (classKey != null)
            {
                // UpperFilters
                var upperFilters = classKey.Values.FirstOrDefault(v => v.ValueName == "UpperFilters")?.ValueData?.ToString() ?? "";
                if (!string.IsNullOrEmpty(upperFilters))
                {
                    var filterList = ParseMultiSzValue(upperFilters);
                    items.Add(new AnalysisItem
                    {
                        Name = "UpperFilters",
                        Value = filterList,
                        RegistryPath = classPath,
                        RegistryValue = $"UpperFilters = {upperFilters}"
                    });

                    foreach (var filter in filterList.Split(new[] { ", " }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        var filterServicePath = $@"{_currentControlSet}\Services\{filter.Trim()}";
                        var filterService = _parser.GetKey(filterServicePath);
                        if (filterService != null)
                        {
                            var imagePath = filterService.Values.FirstOrDefault(v => v.ValueName == "ImagePath")?.ValueData?.ToString() ?? "";
                            var description = filterService.Values.FirstOrDefault(v => v.ValueName == "Description")?.ValueData?.ToString() ?? "";

                            items.Add(new AnalysisItem
                            {
                                Name = $"  → {filter.Trim()}",
                                Value = !string.IsNullOrEmpty(description) ? TruncatePath(description, 80) :
                                       (!string.IsNullOrEmpty(imagePath) ? TruncatePath(imagePath, 80) : "Upper filter driver"),
                                RegistryPath = filterServicePath,
                                RegistryValue = !string.IsNullOrEmpty(imagePath) ? $"ImagePath = {imagePath}" : "Service entry found"
                            });
                        }
                    }
                }
                else
                {
                    items.Add(new AnalysisItem
                    {
                        Name = "UpperFilters",
                        Value = "(none configured)",
                        RegistryPath = classPath,
                        RegistryValue = "UpperFilters not present"
                    });
                }

                // LowerFilters
                var lowerFilters = classKey.Values.FirstOrDefault(v => v.ValueName == "LowerFilters")?.ValueData?.ToString() ?? "";
                if (!string.IsNullOrEmpty(lowerFilters))
                {
                    var filterList = ParseMultiSzValue(lowerFilters);
                    items.Add(new AnalysisItem
                    {
                        Name = "LowerFilters",
                        Value = filterList,
                        RegistryPath = classPath,
                        RegistryValue = $"LowerFilters = {lowerFilters}"
                    });

                    foreach (var filter in filterList.Split(new[] { ", " }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        var filterServicePath = $@"{_currentControlSet}\Services\{filter.Trim()}";
                        var filterService = _parser.GetKey(filterServicePath);
                        if (filterService != null)
                        {
                            var imagePath = filterService.Values.FirstOrDefault(v => v.ValueName == "ImagePath")?.ValueData?.ToString() ?? "";
                            var description = filterService.Values.FirstOrDefault(v => v.ValueName == "Description")?.ValueData?.ToString() ?? "";

                            items.Add(new AnalysisItem
                            {
                                Name = $"  → {filter.Trim()}",
                                Value = !string.IsNullOrEmpty(description) ? TruncatePath(description, 80) :
                                       (!string.IsNullOrEmpty(imagePath) ? TruncatePath(imagePath, 80) : "Lower filter driver"),
                                RegistryPath = filterServicePath,
                                RegistryValue = !string.IsNullOrEmpty(imagePath) ? $"ImagePath = {imagePath}" : "Service entry found"
                            });
                        }
                    }
                }
                else
                {
                    items.Add(new AnalysisItem
                    {
                        Name = "LowerFilters",
                        Value = "(none configured)",
                        RegistryPath = classPath,
                        RegistryValue = "LowerFilters not present"
                    });
                }
            }
            else
            {
                items.Add(new AnalysisItem
                {
                    Name = "Status",
                    Value = $"{className} class configuration not found",
                    RegistryPath = classPath,
                    RegistryValue = "Key not present - requires SYSTEM hive"
                });
            }

            return items;
        }


        /// <summary>
        /// Parse REG_MULTI_SZ value (handles various formats)
        /// </summary>
        private string ParseMultiSzValue(string value)
        {
            if (string.IsNullOrEmpty(value))
                return value;

            // Handle different possible formats from the registry library
            // Could be newline-separated, null-separated, or already comma-separated
            var separators = new[] { '\n', '\r', '\0' };
            var parts = value.Split(separators, StringSplitOptions.RemoveEmptyEntries)
                            .Select(s => s.Trim())
                            .Where(s => !string.IsNullOrEmpty(s))
                            .ToArray();

            return parts.Length > 1 ? string.Join(", ", parts) : value.Trim();
        }

        /// <summary>
        /// Get mounted devices information from SYSTEM\MountedDevices key.
        /// Parses MBR disk signatures, GPT partition GUIDs, and device paths.
        /// </summary>
        public List<MountedDeviceEntry> GetMountedDevices()
        {
            var entries = new List<MountedDeviceEntry>();

            var mountedDevicesKey = _parser.GetKey("MountedDevices");
            if (mountedDevicesKey?.Values == null)
                return entries;

            // Build disk partition registry for cross-referencing (GPT + MBR)
            var diskRegistry = BuildDiskPartitionRegistry();

            // Collect MBR entries for deferred cross-referencing (needs all entries first)
            var mbrEntries = new List<MountedDeviceEntry>();

            foreach (var value in mountedDevicesKey.Values)
            {
                var valueName = value.ValueName;
                if (string.IsNullOrEmpty(valueName))
                    continue;

                var entry = new MountedDeviceEntry
                {
                    RegistryValueName = valueName
                };

                // Determine mount point and type from value name
                if (valueName.StartsWith(@"\DosDevices\", StringComparison.OrdinalIgnoreCase))
                {
                    entry.MountPoint = valueName.Substring(@"\DosDevices\".Length);
                    entry.MountType = "Drive Letter";
                }
                else if (valueName.StartsWith(@"\??\Volume", StringComparison.OrdinalIgnoreCase))
                {
                    var guidPart = valueName.Substring(@"\??\Volume".Length);
                    entry.MountPoint = $"Volume{guidPart}";
                    entry.MountType = "Volume GUID";
                }
                else
                {
                    entry.MountPoint = valueName;
                    entry.MountType = "Other";
                }

                // Get raw binary data
                byte[]? rawData = value.ValueDataRaw;
                if (rawData == null || rawData.Length == 0)
                {
                    entry.PartitionStyle = "Empty";
                    entry.Identifier = "(no data)";
                    entries.Add(entry);
                    continue;
                }

                entry.DataLength = rawData.Length;
                // Parse based on data format
                if (rawData.Length == 12)
                {
                    // MBR: 4-byte disk signature + 8-byte partition byte offset
                    entry.PartitionStyle = "MBR";
                    uint diskSig = BitConverter.ToUInt32(rawData, 0);
                    long partOffset = BitConverter.ToInt64(rawData, 4);
                    long lbaSector = partOffset / 512;

                    entry.DiskSignature = $"0x{diskSig:X8}";
                    entry.PartitionOffset = $"{partOffset:N0} bytes (LBA {lbaSector:N0})";
                    entry.Identifier = $"Sig: 0x{diskSig:X8}, Offset: LBA {lbaSector:N0}";

                    // Collect for deferred cross-referencing after all entries are parsed
                    mbrEntries.Add(entry);
                }
                else if (rawData.Length == 24 && rawData[0] == 0x44 && rawData[1] == 0x4D &&
                         rawData[2] == 0x49 && rawData[3] == 0x4F && rawData[4] == 0x3A &&
                         rawData[5] == 0x49 && rawData[6] == 0x44 && rawData[7] == 0x3A)
                {
                    // GPT: "DMIO:ID:" (8 bytes) + 16-byte partition GUID
                    entry.PartitionStyle = "GPT";
                    var guidBytes = new byte[16];
                    Array.Copy(rawData, 8, guidBytes, 0, 16);
                    var partGuid = new Guid(guidBytes);

                    entry.PartitionGuid = partGuid.ToString("B");
                    entry.Identifier = partGuid.ToString("B");

                    // Cross-reference partition GUID against disk PartitionTableCache blobs
                    var matchedDisk = FindDiskForPartitionGuid(partGuid, diskRegistry);
                    if (matchedDisk != null)
                    {
                        entry.DiskId = matchedDisk.DiskId;
                        EnrichFromEnumPath(entry, matchedDisk.EnumPath);

                        // Update identifier to FriendlyName if available
                        if (!string.IsNullOrEmpty(entry.FriendlyName))
                            entry.Identifier = entry.FriendlyName;
                    }
                }
                else
                {
                    // Variable-length: likely UTF-16LE device path string
                    entry.PartitionStyle = "Device Path";

                    try
                    {
                        string devicePath = System.Text.Encoding.Unicode.GetString(rawData).TrimEnd('\0');
                        entry.DevicePath = devicePath;

                        ParseDevicePath(entry, devicePath);

                        // Enrich with data from ControlSet001\Enum
                        EnrichFromEnum(entry);

                        // Build concise identifier — prefer FriendlyName from Enum
                        if (!string.IsNullOrEmpty(entry.FriendlyName))
                        {
                            entry.Identifier = entry.FriendlyName;
                        }
                        else if (!string.IsNullOrEmpty(entry.Vendor) || !string.IsNullOrEmpty(entry.Product))
                        {
                            var parts = new List<string>();
                            if (!string.IsNullOrEmpty(entry.BusType)) parts.Add(entry.BusType);
                            if (!string.IsNullOrEmpty(entry.Vendor)) parts.Add(entry.Vendor);
                            if (!string.IsNullOrEmpty(entry.Product)) parts.Add(entry.Product);
                            entry.Identifier = string.Join(" | ", parts);
                        }
                        else if (!string.IsNullOrEmpty(entry.BusType))
                        {
                            entry.Identifier = entry.BusType;
                        }
                        else
                        {
                            entry.Identifier = devicePath.Length > 80
                                ? devicePath.Substring(0, 77) + "..."
                                : devicePath;
                        }
                    }
                    catch
                    {
                        entry.Identifier = $"Binary ({rawData.Length} bytes)";
                    }
                }

                entries.Add(entry);
            }

            // MBR cross-referencing: map disk signatures to DiskIds via STORAGE\Volume offsets
            if (mbrEntries.Count > 0)
            {
                // Collect DiskIds already matched by GPT entries so they can be excluded from MBR candidates
                var gptMatchedDiskIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var entry in entries)
                {
                    if (entry.PartitionStyle == "GPT" && !string.IsNullOrEmpty(entry.DiskId))
                    {
                        var id = entry.DiskId.Trim();
                        if (!id.StartsWith("{")) id = "{" + id + "}";
                        gptMatchedDiskIds.Add(id);
                    }
                }

                var mbrSigToDisk = BuildMbrSignatureToDiskMap(mbrEntries, diskRegistry, gptMatchedDiskIds);
                foreach (var entry in mbrEntries)
                {
                    if (string.IsNullOrEmpty(entry.DiskSignature)) continue;
                    if (!uint.TryParse(entry.DiskSignature.Replace("0x", ""),
                        System.Globalization.NumberStyles.HexNumber, null, out uint sig))
                        continue;

                    if (mbrSigToDisk.TryGetValue(sig, out var matchedDisk))
                    {
                        entry.DiskId = matchedDisk.DiskId;
                        entry.EnumPath = matchedDisk.EnumPath;
                        EnrichFromEnumPath(entry, matchedDisk.EnumPath);

                        // Update identifier to FriendlyName if available
                        if (!string.IsNullOrEmpty(entry.FriendlyName))
                            entry.Identifier = entry.FriendlyName;
                    }
                }
            }

            // Orphan/stale detection: compare DiskIds against active STORAGE\Volume registrations
            var activeDiskIds = BuildActiveStorageVolumeDiskIds();

            foreach (var entry in entries)
            {
                if (entry.PartitionStyle == "GPT")
                {
                    if (!string.IsNullOrEmpty(entry.DiskId))
                    {
                        // Extract just the GUID from DiskId (it may have braces already)
                        var diskIdGuid = entry.DiskId.Trim();
                        if (!diskIdGuid.StartsWith("{"))
                            diskIdGuid = "{" + diskIdGuid + "}";
                        entry.StaleStatus = activeDiskIds.Contains(diskIdGuid) ? "Active" : "Stale";
                    }
                    else
                    {
                        entry.StaleStatus = "Unknown";
                    }
                }
                else if (entry.PartitionStyle == "MBR")
                {
                    if (!string.IsNullOrEmpty(entry.DiskId))
                    {
                        var diskIdGuid = entry.DiskId.Trim();
                        if (!diskIdGuid.StartsWith("{"))
                            diskIdGuid = "{" + diskIdGuid + "}";
                        entry.StaleStatus = activeDiskIds.Contains(diskIdGuid) ? "Active" : "Stale";
                    }
                    // MBR entries without DiskId: leave StaleStatus as "" (no cross-reference found)
                }
                else if (entry.PartitionStyle == "Device Path")
                {
                    if (!string.IsNullOrEmpty(entry.EnumPath))
                    {
                        var enumKey = GetCachedKey(entry.EnumPath);
                        entry.StaleStatus = enumKey != null ? "Active" : "Stale";
                    }
                }
            }

            // Sort: drive letters first (alphabetical), then volume GUIDs
            return entries
                .OrderBy(e => e.MountType == "Drive Letter" ? 0 : e.MountType == "Volume GUID" ? 1 : 2)
                .ThenBy(e => e.MountPoint)
                .ToList();
        }

        /// <summary>
        /// Enumerate all physical disk devices from ControlSet001\Enum across all bus types
        /// and detect probable Storage Spaces pool members.
        /// </summary>
        public List<PhysicalDiskEntry> GetPhysicalDisks()
        {
            var disks = new List<PhysicalDiskEntry>();

            // Bus types that contain disk devices, and their key prefixes
            var busTypes = new[]
            {
                ("IDE", $@"{_currentControlSet}\Enum\IDE"),
                ("SCSI", $@"{_currentControlSet}\Enum\SCSI"),
                ("USBSTOR", $@"{_currentControlSet}\Enum\USBSTOR"),
                ("NVME", $@"{_currentControlSet}\Enum\NVME"),
            };

            // Build set of active STORAGE\Volume DiskIds for pool detection
            var activeDiskIds = BuildActiveStorageVolumeDiskIds();

            // Build MountedDevices drive letter mapping: DiskId -> list of drive letters
            var diskIdToDriveLetters = BuildDiskIdToDriveLetterMap();

            foreach (var (busType, basePath) in busTypes)
            {
                var busKey = GetCachedKey(basePath);
                if (busKey?.SubKeys == null) continue;

                foreach (var deviceTypeKey in busKey.SubKeys)
                {
                    var deviceTypeName = deviceTypeKey.KeyName;
                    if (string.IsNullOrEmpty(deviceTypeName)) continue;

                    // Only process disk devices (skip CdRom, other device types)
                    bool isDisk = deviceTypeName.StartsWith("Disk", StringComparison.OrdinalIgnoreCase);
                    if (!isDisk) continue;

                    // Read the device type subkey to get instances
                    var deviceTypeKeyPath = $@"{basePath}\{deviceTypeName}";
                    var fullDeviceTypeKey = GetCachedKey(deviceTypeKeyPath);
                    if (fullDeviceTypeKey?.SubKeys == null) continue;

                    foreach (var instanceKey in fullDeviceTypeKey.SubKeys)
                    {
                        var instanceName = instanceKey.KeyName;
                        if (string.IsNullOrEmpty(instanceName)) continue;

                        var instancePath = $@"{deviceTypeKeyPath}\{instanceName}";
                        var instanceFullKey = GetCachedKey(instancePath);
                        if (instanceFullKey?.Values == null) continue;

                        var entry = new PhysicalDiskEntry
                        {
                            BusType = busType,
                            DeviceId = $@"{busType}\{deviceTypeName}\{instanceName}",
                            EnumPath = instancePath,
                            RegistryPath = instancePath,
                        };

                        // Read standard values
                        string? ReadVal(string name)
                        {
                            return instanceFullKey.Values
                                .FirstOrDefault(v => string.Equals(v.ValueName, name, StringComparison.OrdinalIgnoreCase))
                                ?.ValueData?.ToString();
                        }

                        string? CleanResourceString(string? val)
                        {
                            if (val == null) return null;
                            var semi = val.LastIndexOf(';');
                            return semi >= 0 ? val.Substring(semi + 1) : val;
                        }

                        entry.FriendlyName = CleanResourceString(ReadVal("FriendlyName")) ?? "";
                        entry.DeviceDesc = CleanResourceString(ReadVal("DeviceDesc")) ?? "";
                        entry.Manufacturer = CleanResourceString(ReadVal("Mfg")) ?? "";
                        entry.Service = ReadVal("Service") ?? "";
                        entry.LocationInfo = ReadVal("LocationInformation") ?? "";
                        entry.DeviceClass = ReadVal("Class") ?? "";

                        // FriendlyName fallback to DeviceDesc
                        if (string.IsNullOrEmpty(entry.FriendlyName) && !string.IsNullOrEmpty(entry.DeviceDesc))
                            entry.FriendlyName = entry.DeviceDesc;

                        // HardwareID (first entry only)
                        var hwIdRaw = ReadVal("HardwareID");
                        if (!string.IsNullOrEmpty(hwIdRaw))
                        {
                            var firstId = hwIdRaw.Split(new[] { '\0', ' ' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
                            entry.HardwareId = firstId ?? hwIdRaw;
                        }

                        // CompatibleIDs
                        entry.CompatibleIds = ReadVal("CompatibleIDs") ?? "";

                        // ConfigFlags -> status
                        var configFlags = ReadVal("ConfigFlags");
                        if (!string.IsNullOrEmpty(configFlags))
                        {
                            if (int.TryParse(configFlags.Split(' ')[0], out int flags))
                                entry.DeviceStatus = (flags & 1) == 1 ? "Disabled" : "Enabled";
                        }
                        else
                        {
                            entry.DeviceStatus = "Enabled";
                        }

                        // Read DiskId from Device Parameters\Partmgr
                        var partmgrPath = $@"{instancePath}\Device Parameters\Partmgr";
                        var partmgrKey = GetCachedKey(partmgrPath);
                        if (partmgrKey?.Values != null)
                        {
                            var diskIdVal = partmgrKey.Values
                                .FirstOrDefault(v => string.Equals(v.ValueName, "DiskId", StringComparison.OrdinalIgnoreCase))
                                ?.ValueData?.ToString();
                            if (!string.IsNullOrEmpty(diskIdVal))
                            {
                                entry.DiskId = diskIdVal;

                                // Check if this DiskId has STORAGE\Volume entries
                                var normalizedId = diskIdVal.Trim();
                                if (!normalizedId.StartsWith("{"))
                                    normalizedId = "{" + normalizedId + "}";

                                if (activeDiskIds.Contains(normalizedId))
                                {
                                    entry.PoolStatus = "Normal";
                                    entry.VolumeCount = CountVolumesForDiskId(normalizedId);
                                }
                                else
                                {
                                    entry.PoolStatus = "Probable Pool Member";
                                    entry.VolumeCount = 0;
                                }

                                // Map drive letters
                                if (diskIdToDriveLetters.TryGetValue(normalizedId, out var letters))
                                {
                                    entry.DriveLetters = string.Join(", ", letters);
                                }
                            }
                        }

                        disks.Add(entry);
                    }
                }
            }

            // Sort: pool members first, then by bus type, then by location
            return disks
                .OrderByDescending(d => d.PoolStatus == "Probable Pool Member" ? 1 : 0)
                .ThenBy(d => d.BusType)
                .ThenBy(d => d.LocationInfo)
                .ToList();
        }

        /// <summary>
        /// Count STORAGE\Volume entries for a given DiskId GUID.
        /// </summary>
        private int CountVolumesForDiskId(string diskIdGuid)
        {
            var volumeKey = GetCachedKey($@"{_currentControlSet}\Enum\STORAGE\Volume");
            if (volumeKey?.SubKeys == null) return 0;

            int count = 0;
            foreach (var subKey in volumeKey.SubKeys)
            {
                var name = subKey.KeyName;
                if (!string.IsNullOrEmpty(name) &&
                    name.IndexOf(diskIdGuid, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    count++;
                }
            }
            return count;
        }

        /// <summary>
        /// Build a mapping from DiskId GUIDs to drive letters using MountedDevices + Partmgr cross-reference.
        /// </summary>
        private Dictionary<string, List<string>> BuildDiskIdToDriveLetterMap()
        {
            var map = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

            var mountedDevices = GetMountedDevices();
            foreach (var device in mountedDevices)
            {
                if (device.MountType != "Drive Letter" || string.IsNullOrEmpty(device.DiskId))
                    continue;

                var normalizedId = device.DiskId.Trim();
                if (!normalizedId.StartsWith("{"))
                    normalizedId = "{" + normalizedId + "}";

                if (!map.TryGetValue(normalizedId, out var letters))
                {
                    letters = new List<string>();
                    map[normalizedId] = letters;
                }
                letters.Add(device.MountPoint);
            }

            return map;
        }

        /// <summary>
        /// Parse device path string to extract bus type, vendor, product, and serial number.
        /// Handles paths like \??\SCSI#Disk&amp;Ven_Samsung&amp;Prod_SSD#serial#{guid}
        /// and \??\USBSTOR#Disk&amp;Ven_SanDisk&amp;Prod_Ultra&amp;Rev_1.00#SERIAL#{guid}
        /// </summary>
        private static void ParseDevicePath(MountedDeviceEntry entry, string devicePath)
        {
            var path = devicePath;
            if (path.StartsWith(@"\??\"))
                path = path.Substring(4);

            // Split by # to get segments: [BusAndDevice, DeviceId, InstanceId, InterfaceGuid]
            var segments = path.Split('#');
            if (segments.Length < 2)
                return;

            var firstSegment = segments[0];

            // Extract bus type (before first backslash)
            var busSlash = firstSegment.IndexOf('\\');
            if (busSlash >= 0)
            {
                entry.BusType = firstSegment.Substring(0, busSlash);
            }
            else
            {
                entry.BusType = firstSegment;
            }

            // Parse Ven_ and Prod_ - check both first and second segments
            // Standard Windows device paths put Ven_/Prod_ in segments[1]:
            //   \??\SCSI#CdRom&Ven_Msft&Prod_Virtual_DVD-ROM#serial#{guid}
            //   \??\USBSTOR#Disk&Ven_SanDisk&Prod_Ultra&Rev_1.00#serial#{guid}
            // But some may have them in segments[0] if joined by backslash:
            //   \??\SCSI\Disk&Ven_Samsung&Prod_SSD#serial#{guid}
            var searchSegment = firstSegment;
            if (segments.Length >= 2 && firstSegment.IndexOf("Ven_", StringComparison.OrdinalIgnoreCase) < 0)
            {
                searchSegment = segments[1];
            }

            var venIndex = searchSegment.IndexOf("Ven_", StringComparison.OrdinalIgnoreCase);
            if (venIndex >= 0)
            {
                var venStart = venIndex + 4;
                var venEnd = searchSegment.IndexOf('&', venStart);
                entry.Vendor = venEnd >= 0
                    ? searchSegment.Substring(venStart, venEnd - venStart).Replace('_', ' ').Trim()
                    : searchSegment.Substring(venStart).Replace('_', ' ').Trim();
            }

            var prodIndex = searchSegment.IndexOf("Prod_", StringComparison.OrdinalIgnoreCase);
            if (prodIndex >= 0)
            {
                var prodStart = prodIndex + 5;
                var prodEnd = searchSegment.IndexOf('&', prodStart);
                entry.Product = prodEnd >= 0
                    ? searchSegment.Substring(prodStart, prodEnd - prodStart).Replace('_', ' ').Trim()
                    : searchSegment.Substring(prodStart).Replace('_', ' ').Trim();
            }

            // Serial number is typically in the third segment (index 2)
            // Skip segments that look like GUIDs or contain Ven_/Prod_ device descriptors
            if (segments.Length >= 3 && !string.IsNullOrEmpty(segments[2])
                && !segments[2].StartsWith("{"))
            {
                entry.Serial = segments[2].Trim();
            }
        }

        /// <summary>
        /// Enrich a MountedDeviceEntry with data from ControlSet001\Enum\{Bus}\{Device}\{Instance}
        /// Parses the device path to construct the enum path, then delegates to EnrichFromEnumPath.
        /// </summary>
        private void EnrichFromEnum(MountedDeviceEntry entry)
        {
            if (string.IsNullOrEmpty(entry.DevicePath))
                return;

            // Build Enum path from device path segments
            var path = entry.DevicePath;
            if (path.StartsWith(@"\??\"))
                path = path.Substring(4);

            var segments = path.Split('#');
            if (segments.Length < 3)
                return;

            // Construct: ControlSet001\Enum\SCSI\CdRom&Ven_Msft&Prod_Virtual_DVD-ROM\5&394b69d0&0&000002
            var enumPath = $@"{_currentControlSet}\Enum\{segments[0]}\{segments[1]}\{segments[2]}";
            EnrichFromEnumPath(entry, enumPath);
        }

        /// <summary>
        /// Enrich a MountedDeviceEntry with data from a known Enum path.
        /// Reads FriendlyName, DeviceDesc, Class, Service, Mfg, LocationInformation, ConfigFlags.
        /// </summary>
        private void EnrichFromEnumPath(MountedDeviceEntry entry, string enumPath)
        {
            entry.EnumPath = enumPath;

            var enumKey = GetCachedKey(enumPath);
            if (enumKey?.Values == null)
                return;

            // Helper to read and clean localized strings (strips "@resource;Display Name" → "Display Name")
            string? ReadAndClean(string valueName, bool clean = false)
            {
                var val = enumKey.Values.FirstOrDefault(v => v.ValueName == valueName)?.ValueData?.ToString();
                if (clean && val != null)
                {
                    var semi = val.LastIndexOf(';');
                    if (semi >= 0) val = val.Substring(semi + 1);
                }
                return val;
            }

            var friendlyName = ReadAndClean("FriendlyName", clean: true);
            if (!string.IsNullOrEmpty(friendlyName))
                entry.FriendlyName = friendlyName;

            var deviceDesc = ReadAndClean("DeviceDesc", clean: true);
            // Use DeviceDesc as fallback FriendlyName if FriendlyName is empty
            if (string.IsNullOrEmpty(entry.FriendlyName) && !string.IsNullOrEmpty(deviceDesc))
                entry.FriendlyName = deviceDesc;

            var deviceClass = ReadAndClean("Class");
            if (!string.IsNullOrEmpty(deviceClass))
                entry.DeviceClass = deviceClass;

            var service = ReadAndClean("Service");
            if (!string.IsNullOrEmpty(service))
                entry.DeviceService = service;

            var mfg = ReadAndClean("Mfg", clean: true);
            if (!string.IsNullOrEmpty(mfg))
                entry.Manufacturer = mfg;

            var location = ReadAndClean("LocationInformation");
            if (!string.IsNullOrEmpty(location))
                entry.LocationInfo = location;

            // ConfigFlags: bit 0 = disabled
            var configFlags = ReadAndClean("ConfigFlags");
            if (!string.IsNullOrEmpty(configFlags))
            {
                if (int.TryParse(configFlags.Split(' ')[0], out int flags))
                {
                    entry.DeviceStatus = (flags & 1) == 1 ? "Disabled" : "Enabled";
                }
            }
            // If ConfigFlags is absent, leave DeviceStatus empty (unknown)
        }

        /// <summary>
        /// Internal class holding disk partition registry info for cross-referencing.
        /// </summary>
        private class DiskPartitionInfo
        {
            public string EnumPath { get; set; } = "";
            public string DiskId { get; set; } = "";
            public byte[] PartitionTableCache { get; set; } = Array.Empty<byte>();
        }

        /// <summary>
        /// Build a registry of all disk devices with their DiskId and optional PartitionTableCache.
        /// Iterates ControlSet001\Enum\*\*\* filtering by Service=disk.
        /// MBR disks may not have PartitionTableCache; they are still collected for DiskId-based matching.
        /// Results are lazy-cached for reuse.
        /// </summary>
        private List<DiskPartitionInfo> BuildDiskPartitionRegistry()
        {
            if (_diskPartitionRegistry != null)
                return _diskPartitionRegistry;

            _diskPartitionRegistry = new List<DiskPartitionInfo>();

            var enumKey = GetCachedKey($@"{_currentControlSet}\Enum");
            if (enumKey?.SubKeys == null)
                return _diskPartitionRegistry;

            // Three-level walk: bus → device → instance
            foreach (var busType in enumKey.SubKeys)
            {
                if (busType.SubKeys == null) continue;
                foreach (var device in busType.SubKeys)
                {
                    if (device.SubKeys == null) continue;
                    foreach (var instance in device.SubKeys)
                    {
                        // Filter by Service == "disk"
                        var serviceVal = instance.Values?.FirstOrDefault(v =>
                            string.Equals(v.ValueName, "Service", StringComparison.OrdinalIgnoreCase))?.ValueData?.ToString();
                        if (!string.Equals(serviceVal, "disk", StringComparison.OrdinalIgnoreCase))
                            continue;

                        // Navigate to Device Parameters\Partmgr via SubKeys
                        var deviceParams = instance.SubKeys?.FirstOrDefault(s =>
                            string.Equals(s.KeyName, "Device Parameters", StringComparison.OrdinalIgnoreCase));
                        var partmgr = deviceParams?.SubKeys?.FirstOrDefault(s =>
                            string.Equals(s.KeyName, "Partmgr", StringComparison.OrdinalIgnoreCase));
                        if (partmgr?.Values == null)
                            continue;

                        var diskIdVal = partmgr.Values.FirstOrDefault(v =>
                            string.Equals(v.ValueName, "DiskId", StringComparison.OrdinalIgnoreCase))?.ValueData?.ToString();

                        // DiskId is required; PartitionTableCache is optional (MBR disks may not have it)
                        if (string.IsNullOrEmpty(diskIdVal))
                            continue;

                        var cacheVal = partmgr.Values.FirstOrDefault(v =>
                            string.Equals(v.ValueName, "PartitionTableCache", StringComparison.OrdinalIgnoreCase));

                        var enumPath = $@"{_currentControlSet}\Enum\{busType.KeyName}\{device.KeyName}\{instance.KeyName}";

                        _diskPartitionRegistry.Add(new DiskPartitionInfo
                        {
                            EnumPath = enumPath,
                            DiskId = diskIdVal,
                            PartitionTableCache = cacheVal?.ValueDataRaw ?? Array.Empty<byte>()
                        });
                    }
                }
            }

            return _diskPartitionRegistry;
        }

        /// <summary>
        /// Find which disk contains a given partition GUID by byte-searching PartitionTableCache blobs.
        /// Uses Guid.ToByteArray() (mixed-endian) for the 16-byte search pattern.
        /// </summary>
        private static DiskPartitionInfo? FindDiskForPartitionGuid(Guid partitionGuid, List<DiskPartitionInfo> diskRegistry)
        {
            var guidBytes = partitionGuid.ToByteArray(); // 16 bytes, mixed-endian

            foreach (var disk in diskRegistry)
            {
                var cache = disk.PartitionTableCache;
                if (cache.Length < 16)
                    continue;

                // Brute-force byte search for the 16-byte GUID pattern
                for (int i = 0; i <= cache.Length - 16; i++)
                {
                    bool match = true;
                    for (int j = 0; j < 16; j++)
                    {
                        if (cache[i + j] != guidBytes[j])
                        {
                            match = false;
                            break;
                        }
                    }
                    if (match)
                        return disk;
                }
            }

            return null;
        }

        /// <summary>
        /// Build a set of DiskIds that have active STORAGE\Volume registrations.
        /// Parses subkey names under ControlSet001\Enum\STORAGE\Volume (format: {DiskId}#PartitionByteOffset).
        /// Results are lazy-cached for reuse.
        /// </summary>
        private HashSet<string> BuildActiveStorageVolumeDiskIds()
        {
            if (_activeStorageVolumeDiskIds != null)
                return _activeStorageVolumeDiskIds;

            _activeStorageVolumeDiskIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            var volumeKey = GetCachedKey($@"{_currentControlSet}\Enum\STORAGE\Volume");
            if (volumeKey?.SubKeys == null)
                return _activeStorageVolumeDiskIds;

            foreach (var subKey in volumeKey.SubKeys)
            {
                var name = subKey.KeyName;
                if (string.IsNullOrEmpty(name))
                    continue;

                // Format: _??_USBSTOR#Disk&Ven_...#Serial#{DiskId}#PartOffset
                // or simpler: {DiskId}#PartOffset
                // The DiskId is a GUID in braces — find it
                var hashIndex = name.LastIndexOf('#');
                if (hashIndex > 0)
                {
                    var beforeHash = name.Substring(0, hashIndex);
                    // Find the last GUID-like segment before the final #
                    // Look for a '{' that starts a GUID
                    var braceStart = beforeHash.LastIndexOf('{');
                    if (braceStart >= 0)
                    {
                        var braceEnd = beforeHash.IndexOf('}', braceStart);
                        if (braceEnd > braceStart)
                        {
                            var diskId = beforeHash.Substring(braceStart, braceEnd - braceStart + 1);
                            _activeStorageVolumeDiskIds.Add(diskId);
                        }
                    }
                }
            }

            return _activeStorageVolumeDiskIds;
        }

        /// <summary>
        /// Build a mapping of MBR disk signatures to DiskIds by cross-referencing
        /// MBR entries from MountedDevices against disk partition data.
        /// 
        /// Algorithm:
        /// 0. Direct match: scan PartitionTableCache blobs for MBR partition style (bytes[0..3]==0)
        ///    and extract the embedded MBR disk signature from bytes[8..11]
        /// 1. Fallback: parse STORAGE\Volume subkeys to build (DiskId, Offset) pairs
        /// 2. Group MBR entries by disk signature, collecting their partition offsets
        /// 3. For each signature, find DiskIds whose volume offsets contain all of the signature's offsets
        /// 4. Resolve unique matches first, then disambiguate remaining via elimination
        /// </summary>
        private Dictionary<uint, DiskPartitionInfo> BuildMbrSignatureToDiskMap(
            List<MountedDeviceEntry> mbrEntries, List<DiskPartitionInfo> diskRegistry,
            HashSet<string> excludedDiskIds)
        {
            var result = new Dictionary<uint, DiskPartitionInfo>();

            // Step 0: Direct MBR signature matching via PartitionTableCache
            // MBR disks with PartitionTableCache have: bytes[0..3] == 0 (MBR style),
            // bytes[8..11] = MBR disk signature (uint32 little-endian)
            var directMbrMap = new Dictionary<uint, DiskPartitionInfo>();
            foreach (var disk in diskRegistry)
            {
                var cache = disk.PartitionTableCache;
                if (cache == null || cache.Length < 12)
                    continue;

                // Check partition style: bytes 0-3 == 0 means MBR
                if (cache[0] != 0 || cache[1] != 0 || cache[2] != 0 || cache[3] != 0)
                    continue;

                uint mbrSig = BitConverter.ToUInt32(cache, 8);
                if (mbrSig != 0)
                    directMbrMap.TryAdd(mbrSig, disk);
            }

            // Apply direct matches immediately
            foreach (var (sig, disk) in directMbrMap)
            {
                result[sig] = disk;
            }

            if (mbrEntries.Count == 0 || diskRegistry.Count == 0)
                return result;

            // Step 1: Parse STORAGE\Volume subkeys → (DiskId, Offset) pairs
            // Format: {DiskId}#PartitionByteOffsetHex
            var volumeKey = GetCachedKey($@"{_currentControlSet}\Enum\STORAGE\Volume");
            if (volumeKey?.SubKeys == null)
                return result;

            // Build lookup: offset → set of DiskIds that have a volume at that offset
            var offsetToDiskIds = new Dictionary<long, HashSet<string>>();
            foreach (var subKey in volumeKey.SubKeys)
            {
                var name = subKey.KeyName;
                if (string.IsNullOrEmpty(name))
                    continue;

                var hashIdx = name.LastIndexOf('#');
                if (hashIdx <= 0) continue;

                var beforeHash = name.Substring(0, hashIdx);
                var offsetHex = name.Substring(hashIdx + 1);

                // Extract DiskId GUID from before the hash
                var braceStart = beforeHash.LastIndexOf('{');
                if (braceStart < 0) continue;
                var braceEnd = beforeHash.IndexOf('}', braceStart);
                if (braceEnd <= braceStart) continue;
                var diskIdStr = beforeHash.Substring(braceStart, braceEnd - braceStart + 1);

                if (long.TryParse(offsetHex, System.Globalization.NumberStyles.HexNumber, null, out long offset))
                {
                    if (!offsetToDiskIds.TryGetValue(offset, out var diskIdSet))
                    {
                        diskIdSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        offsetToDiskIds[offset] = diskIdSet;
                    }
                    diskIdSet.Add(diskIdStr);
                }
            }

            // Step 2: Group MBR entries by disk signature → collect partition offsets per signature
            var sigToOffsets = new Dictionary<uint, HashSet<long>>();
            foreach (var entry in mbrEntries)
            {
                if (string.IsNullOrEmpty(entry.DiskSignature)) continue;
                // Parse "0xHHHHHHHH" back to uint
                if (!uint.TryParse(entry.DiskSignature.Replace("0x", ""),
                    System.Globalization.NumberStyles.HexNumber, null, out uint sig))
                    continue;

                if (!sigToOffsets.TryGetValue(sig, out var offsets))
                {
                    offsets = new HashSet<long>();
                    sigToOffsets[sig] = offsets;
                }

                // Parse partition offset from entry (format: "NNN bytes (LBA NNN)")
                if (!string.IsNullOrEmpty(entry.PartitionOffset))
                {
                    var spaceIdx = entry.PartitionOffset.IndexOf(' ');
                    if (spaceIdx > 0)
                    {
                        var bytesStr = entry.PartitionOffset.Substring(0, spaceIdx).Replace(",", "");
                        if (long.TryParse(bytesStr, out long offsetVal))
                            offsets.Add(offsetVal);
                    }
                }
            }

            // Build DiskId → DiskPartitionInfo lookup from disk registry
            var diskIdToInfo = new Dictionary<string, DiskPartitionInfo>(StringComparer.OrdinalIgnoreCase);
            foreach (var disk in diskRegistry)
            {
                var key = disk.DiskId.Trim();
                if (!key.StartsWith("{")) key = "{" + key + "}";
                if (!diskIdToInfo.ContainsKey(key))
                    diskIdToInfo[key] = disk;
            }

            // Step 3: For each signature, find candidate DiskIds
            var sigToCandidates = new Dictionary<uint, HashSet<string>>();
            foreach (var (sig, offsets) in sigToOffsets)
            {
                HashSet<string>? candidates = null;
                foreach (var offset in offsets)
                {
                    if (offsetToDiskIds.TryGetValue(offset, out var diskIdsAtOffset))
                    {
                        if (candidates == null)
                            candidates = new HashSet<string>(diskIdsAtOffset, StringComparer.OrdinalIgnoreCase);
                        else
                            candidates.IntersectWith(diskIdsAtOffset);
                    }
                    else
                    {
                        // No volume at this offset — no candidates
                        candidates = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        break;
                    }
                }
                sigToCandidates[sig] = candidates ?? new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }

            // Step 4: Resolve unique matches first
            // Seed with GPT-matched DiskIds and direct-matched DiskIds to exclude from STORAGE\Volume candidates
            var matchedDiskIds = new HashSet<string>(excludedDiskIds, StringComparer.OrdinalIgnoreCase);
            foreach (var (_, disk) in directMbrMap)
            {
                var id = disk.DiskId.Trim();
                if (!id.StartsWith("{")) id = "{" + id + "}";
                matchedDiskIds.Add(id);
            }
            foreach (var (sig, candidates) in sigToCandidates)
            {
                if (result.ContainsKey(sig)) continue; // Skip sigs already resolved by direct match
                if (candidates.Count == 1)
                {
                    var diskIdStr = candidates.First();
                    if (diskIdToInfo.TryGetValue(diskIdStr, out var info))
                    {
                        result[sig] = info;
                        matchedDiskIds.Add(diskIdStr);
                    }
                }
            }

            // Step 5: Disambiguate remaining via elimination of already-matched DiskIds
            foreach (var (sig, candidates) in sigToCandidates)
            {
                if (result.ContainsKey(sig)) continue;
                var remaining = new HashSet<string>(candidates, StringComparer.OrdinalIgnoreCase);
                remaining.ExceptWith(matchedDiskIds);
                if (remaining.Count == 1)
                {
                    var diskIdStr = remaining.First();
                    if (diskIdToInfo.TryGetValue(diskIdStr, out var info))
                    {
                        result[sig] = info;
                        matchedDiskIds.Add(diskIdStr);
                    }
                }
            }

            return result;
        }

        #endregion

        #region Software Analysis (SOFTWARE hive)

        /// <summary>
        /// Detect installed .NET Framework versions from the SOFTWARE hive
        /// </summary>
        public AnalysisSection GetDotNetFrameworkAnalysis()
        {
            var section = new AnalysisSection { Title = "🔷 .NET Framework" };
            string ndpBasePath = @"Microsoft\NET Framework Setup\NDP";

            try
            {
                var ndpKey = _parser.GetKey(ndpBasePath);
                if (ndpKey?.SubKeys == null)
                {
                    section.Items.Add(new AnalysisItem
                    {
                        Name = "Info",
                        Value = "No .NET Framework versions detected",
                        RegistryPath = ndpBasePath
                    });
                    return section;
                }

                // Detect older versions (v1.1, v2.0, v3.0, v3.5)
                var olderVersionKeys = new[] { "v1.1.4322", "v2.0.50727", "v3.0", "v3.5" };
                foreach (var versionKeyName in olderVersionKeys)
                {
                    var versionKey = ndpKey.SubKeys
                        .FirstOrDefault(k => k.KeyName.Equals(versionKeyName, StringComparison.OrdinalIgnoreCase));
                    if (versionKey == null)
                        continue;

                    var install = versionKey.Values
                        .FirstOrDefault(v => v.ValueName == "Install")?.ValueData?.ToString() ?? "";
                    var version = versionKey.Values
                        .FirstOrDefault(v => v.ValueName == "Version")?.ValueData?.ToString() ?? "";
                    var sp = versionKey.Values
                        .FirstOrDefault(v => v.ValueName == "SP")?.ValueData?.ToString() ?? "";

                    // For v3.0, Install might be in the Setup sub-key
                    if (string.IsNullOrEmpty(install) || install != "1")
                    {
                        var setupKey = versionKey.SubKeys?
                            .FirstOrDefault(k => k.KeyName.Equals("Setup", StringComparison.OrdinalIgnoreCase));
                        if (setupKey != null)
                        {
                            var installSuccess = setupKey.Values
                                .FirstOrDefault(v => v.ValueName == "InstallSuccess")?.ValueData?.ToString() ?? "";
                            if (installSuccess == "1")
                                install = "1";
                        }
                    }

                    if (install != "1")
                        continue;

                    // Derive friendly name from key name
                    var friendlyName = versionKeyName.StartsWith("v")
                        ? versionKeyName.Substring(1) : versionKeyName;
                    // Trim build suffix for display (e.g. "2.0.50727" -> "2.0")
                    var dotCount = 0;
                    var trimIndex = friendlyName.Length;
                    for (int i = 0; i < friendlyName.Length; i++)
                    {
                        if (friendlyName[i] == '.')
                        {
                            dotCount++;
                            if (dotCount == 2) { trimIndex = i; break; }
                        }
                    }
                    var shortName = friendlyName.Substring(0, trimIndex);

                    var valueText = !string.IsNullOrEmpty(version) ? version : friendlyName;
                    if (!string.IsNullOrEmpty(sp) && sp != "0")
                        valueText += $" SP{sp}";

                    var regPath = $@"{ndpBasePath}\{versionKeyName}";
                    section.Items.Add(new AnalysisItem
                    {
                        Name = $".NET Framework {shortName}",
                        Value = valueText,
                        RegistryPath = regPath,
                        RegistryValue = "Version"
                    });
                }

                // Detect .NET Framework 4.5+ via v4\Full Release DWORD
                var v4FullPath = $@"{ndpBasePath}\v4\Full";
                var v4FullKey = _parser.GetKey(v4FullPath);
                var detectedV4 = false;

                if (v4FullKey != null)
                {
                    var releaseStr = v4FullKey.Values
                        .FirstOrDefault(v => v.ValueName == "Release")?.ValueData?.ToString() ?? "";
                    var exactVersion = v4FullKey.Values
                        .FirstOrDefault(v => v.ValueName == "Version")?.ValueData?.ToString() ?? "";

                    var releaseNum = ParseDwordValue(releaseStr);
                    var friendlyVersion = GetDotNet45PlusVersion(releaseNum);

                    if (!string.IsNullOrEmpty(friendlyVersion))
                    {
                        var valueText = !string.IsNullOrEmpty(exactVersion)
                            ? $"{exactVersion} (Release: {releaseNum})"
                            : $"Release: {releaseNum}";

                        section.Items.Add(new AnalysisItem
                        {
                            Name = $".NET Framework {friendlyVersion}",
                            Value = valueText,
                            RegistryPath = v4FullPath,
                            RegistryValue = "Release"
                        });
                        detectedV4 = true;
                    }
                    else
                    {
                        // v4\Full exists but no Release DWORD → .NET 4.0 Full profile
                        var install = v4FullKey.Values
                            .FirstOrDefault(v => v.ValueName == "Install")?.ValueData?.ToString() ?? "";
                        if (install == "1" && !string.IsNullOrEmpty(exactVersion))
                        {
                            section.Items.Add(new AnalysisItem
                            {
                                Name = ".NET Framework 4.0",
                                Value = exactVersion,
                                RegistryPath = v4FullPath,
                                RegistryValue = "Version"
                            });
                            detectedV4 = true;
                        }
                    }
                }

                // Fallback: v4\Client profile (4.0 without Full profile)
                if (!detectedV4)
                {
                    var v4ClientPath = $@"{ndpBasePath}\v4\Client";
                    var v4ClientKey = _parser.GetKey(v4ClientPath);
                    if (v4ClientKey != null)
                    {
                        var install = v4ClientKey.Values
                            .FirstOrDefault(v => v.ValueName == "Install")?.ValueData?.ToString() ?? "";
                        var version = v4ClientKey.Values
                            .FirstOrDefault(v => v.ValueName == "Version")?.ValueData?.ToString() ?? "";

                        if (install == "1" && !string.IsNullOrEmpty(version))
                        {
                            section.Items.Add(new AnalysisItem
                            {
                                Name = ".NET Framework 4.0 (Client Profile)",
                                Value = version,
                                RegistryPath = v4ClientPath,
                                RegistryValue = "Version"
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error reading .NET Framework info: {ex.Message}");
            }

            if (section.Items.Count == 0)
            {
                section.Items.Add(new AnalysisItem
                {
                    Name = "Info",
                    Value = "No .NET Framework versions detected",
                    RegistryPath = ndpBasePath
                });
            }

            return section;
        }

        /// <summary>
        /// Map .NET Framework 4.5+ Release DWORD to friendly version string
        /// </summary>
        private static string? GetDotNet45PlusVersion(int releaseKey)
        {
            if (releaseKey >= 533320) return "4.8.1";
            if (releaseKey >= 528040) return "4.8";
            if (releaseKey >= 461808) return "4.7.2";
            if (releaseKey >= 461308) return "4.7.1";
            if (releaseKey >= 460798) return "4.7";
            if (releaseKey >= 394802) return "4.6.2";
            if (releaseKey >= 394254) return "4.6.1";
            if (releaseKey >= 393295) return "4.6";
            if (releaseKey >= 379893) return "4.5.2";
            if (releaseKey >= 378675) return "4.5.1";
            if (releaseKey >= 378389) return "4.5";
            return null;
        }

        /// <summary>
        /// Get software-related information as structured sections
        /// </summary>
        public List<AnalysisSection> GetSoftwareAnalysis()
        {
            var sections = new List<AnalysisSection>();

            // Check if this is SOFTWARE hive
            if (_parser.CurrentHiveType != OfflineRegistryParser.HiveType.SOFTWARE)
            {
                var noticeSection = new AnalysisSection { Title = "ℹ️ Notice" };
                noticeSection.Items.Add(new AnalysisItem
                {
                    Name = "Info",
                    Value = "This information requires the SOFTWARE hive to be loaded",
                    RegistryPath = "",
                    RegistryValue = "Load SOFTWARE hive to view installed software"
                });
                sections.Add(noticeSection);
                return sections;
            }

            // Installed Programs (x64 and x86)
            sections.Add(GetInstalledProgramsAnalysis());

            // Startup Programs (Run and RunOnce)
            sections.Add(GetStartupProgramsAnalysis());
            
            // Appx Applications (combined section - UI will handle filtering)
            sections.Add(GetAppxAnalysis());

            // .NET Framework versions
            sections.Add(GetDotNetFrameworkAnalysis());

            // Scheduled Tasks
            sections.Add(GetScheduledTasksAnalysis());

            return sections;
        }

        /// <summary>
        /// Get combined Appx section (placeholder for UI with filter buttons)
        /// </summary>
        public AnalysisSection GetAppxAnalysis()
        {
            var section = new AnalysisSection { Title = "📱 Appx Packages" };
            
            // Get counts for display
            string inboxPath = @"Microsoft\Windows\CurrentVersion\Appx\AppxAllUserStore\InboxApplications";
            string applicationsPath = @"Microsoft\Windows\CurrentVersion\Appx\AppxAllUserStore\Applications";
            
            var inboxKey = _parser.GetKey(inboxPath);
            var applicationsKey = _parser.GetKey(applicationsPath);
            
            var inboxCount = inboxKey?.SubKeys?.Count ?? 0;
            var userCount = applicationsKey?.SubKeys?.Count ?? 0;
            
            section.Items.Add(new AnalysisItem
            {
                Name = "InBox (Preinstalled)",
                Value = $"{inboxCount} packages",
                RegistryPath = inboxPath
            });
            
            section.Items.Add(new AnalysisItem
            {
                Name = "User Installed",
                Value = $"{userCount} packages",
                RegistryPath = applicationsPath
            });
            
            return section;
        }

        /// <summary>
        /// Get Appx packages by type (for UI filtering)
        /// </summary>
        /// <param name="appxType">Type: "InBox" or "UserInstalled"</param>
        public List<(string PackageName, string Version, string Architecture, string RegistryPath)> GetAppxPackages(string appxType)
        {
            var apps = new List<(string PackageName, string Version, string Architecture, string RegistryPath)>();
            
            string basePath = appxType == "InBox" 
                ? @"Microsoft\Windows\CurrentVersion\Appx\AppxAllUserStore\InboxApplications"
                : @"Microsoft\Windows\CurrentVersion\Appx\AppxAllUserStore\Applications";
            
            var key = _parser.GetKey(basePath);
            if (key?.SubKeys == null) return apps;

            foreach (var app in key.SubKeys)
            {
                var keyName = app.KeyName ?? "";
                var parts = keyName.Split('_');
                
                var packageName = parts.Length > 0 ? parts[0] : keyName;
                var version = parts.Length > 1 ? parts[1] : "";
                var arch = parts.Length > 2 ? parts[2] : "";
                var appPath = $@"{basePath}\{keyName}";

                apps.Add((packageName, version, arch, appPath));
            }

            return apps.OrderBy(a => a.PackageName).ToList();
        }

        /// <summary>
        /// Get startup programs from Run and RunOnce registry keys
        /// </summary>
        public AnalysisSection GetStartupProgramsAnalysis()
        {
            var section = new AnalysisSection { Title = "🚀 Startup Programs" };

            string runPath = @"Microsoft\Windows\CurrentVersion\Run";
            string runOncePath = @"Microsoft\Windows\CurrentVersion\RunOnce";

            var entries = new List<(string Name, string Command, string Source, string RegistryPath)>();

            // Fetch Run entries
            var runKey = _parser.GetKey(runPath);
            if (runKey?.Values != null)
            {
                foreach (var val in runKey.Values)
                {
                    var name = val.ValueName;
                    var command = val.ValueData?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(name))
                    {
                        entries.Add((name, command, "Run", runPath));
                    }
                }
            }

            // Fetch RunOnce entries
            var runOnceKey = _parser.GetKey(runOncePath);
            if (runOnceKey?.Values != null)
            {
                foreach (var val in runOnceKey.Values)
                {
                    var name = val.ValueName;
                    var command = val.ValueData?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(name))
                    {
                        entries.Add((name, command, "RunOnce", runOncePath));
                    }
                }
            }

            // Sort alphabetically by name
            entries = entries.OrderBy(e => e.Name).ToList();

            foreach (var entry in entries)
            {
                section.Items.Add(new AnalysisItem
                {
                    Name = $"[{entry.Source}] {entry.Name}",
                    Value = entry.Command,
                    RegistryPath = entry.RegistryPath,
                    RegistryValue = entry.Name
                });
            }

            if (entries.Count == 0)
            {
                section.Items.Add(new AnalysisItem
                {
                    Name = "Info",
                    Value = "No startup programs found",
                    RegistryPath = runPath
                });
            }

            return section;
        }

        /// <summary>
        /// Get installed programs from Uninstall registry keys (x64 and x86)
        /// </summary>
        public AnalysisSection GetInstalledProgramsAnalysis()
        {
            var section = new AnalysisSection { Title = "📦 Installed Programs" };

            // x64 programs path
            string x64Path = @"Microsoft\Windows\CurrentVersion\Uninstall";
            // x86 (WOW6432Node) programs path
            string x86Path = @"WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall";

            var programs = new List<(string Name, string Version, string Publisher, string InstallDate, string Architecture, string RegistryPath)>();

            // Get x64 programs
            var x64Key = _parser.GetKey(x64Path);
            if (x64Key?.SubKeys != null)
            {
                foreach (var program in x64Key.SubKeys)
                {
                    var displayName = program.Values.FirstOrDefault(v => v.ValueName == "DisplayName")?.ValueData?.ToString();
                    if (string.IsNullOrEmpty(displayName)) continue;
                    
                    // Skip system components and child items (only show parent programs)
                    var systemComponent = program.Values.FirstOrDefault(v => v.ValueName == "SystemComponent")?.ValueData?.ToString() ?? "";
                    var parentKeyName = program.Values.FirstOrDefault(v => v.ValueName == "ParentKeyName")?.ValueData?.ToString() ?? "";
                    if (systemComponent == "1" || !string.IsNullOrEmpty(parentKeyName)) continue;
                    
                    var version = program.Values.FirstOrDefault(v => v.ValueName == "DisplayVersion")?.ValueData?.ToString() ?? "";
                    var publisher = program.Values.FirstOrDefault(v => v.ValueName == "Publisher")?.ValueData?.ToString() ?? "";
                    var installDate = program.Values.FirstOrDefault(v => v.ValueName == "InstallDate")?.ValueData?.ToString() ?? "";
                    var programPath = $@"{x64Path}\{program.KeyName}";
                    
                    programs.Add((displayName, version, publisher, FormatInstallDate(installDate), "x64", programPath));
                }
            }

            // Get x86 programs
            var x86Key = _parser.GetKey(x86Path);
            if (x86Key?.SubKeys != null)
            {
                foreach (var program in x86Key.SubKeys)
                {
                    var displayName = program.Values.FirstOrDefault(v => v.ValueName == "DisplayName")?.ValueData?.ToString();
                    if (string.IsNullOrEmpty(displayName)) continue;
                    
                    // Skip system components and child items (only show parent programs)
                    var systemComponent = program.Values.FirstOrDefault(v => v.ValueName == "SystemComponent")?.ValueData?.ToString() ?? "";
                    var parentKeyName = program.Values.FirstOrDefault(v => v.ValueName == "ParentKeyName")?.ValueData?.ToString() ?? "";
                    if (systemComponent == "1" || !string.IsNullOrEmpty(parentKeyName)) continue;
                    
                    var version = program.Values.FirstOrDefault(v => v.ValueName == "DisplayVersion")?.ValueData?.ToString() ?? "";
                    var publisher = program.Values.FirstOrDefault(v => v.ValueName == "Publisher")?.ValueData?.ToString() ?? "";
                    var installDate = program.Values.FirstOrDefault(v => v.ValueName == "InstallDate")?.ValueData?.ToString() ?? "";
                    var programPath = $@"{x86Path}\{program.KeyName}";
                    
                    // Check if already added from x64 (some programs register in both)
                    if (!programs.Any(p => p.Name == displayName && p.Version == version))
                    {
                        programs.Add((displayName, version, publisher, FormatInstallDate(installDate), "x86", programPath));
                    }
                }
            }

            // Sort by name
            programs = programs.OrderBy(p => p.Name).ToList();

            // Add counts header
            var x64Count = programs.Count(p => p.Architecture == "x64");
            var x86Count = programs.Count(p => p.Architecture == "x86");
            
            section.Items.Add(new AnalysisItem
            {
                Name = "Summary",
                Value = $"Total: {programs.Count} programs ({x64Count} x64, {x86Count} x86)",
                RegistryPath = x64Path,
                IsSubSection = false
            });

            // Add programs grouped by architecture
            foreach (var prog in programs)
            {
                var archIcon = prog.Architecture == "x64" ? "🔷" : "🔶";
                section.Items.Add(new AnalysisItem
                {
                    Name = $"{archIcon} {prog.Name}",
                    Value = $"[{prog.Architecture}]",
                    RegistryPath = prog.RegistryPath,
                    RegistryValue = $"DisplayName = {prog.Name}",
                    IsSubSection = true,
                    SubItems = new List<AnalysisItem>
                    {
                        new AnalysisItem { Name = "Version", Value = prog.Version, RegistryValue = "DisplayVersion" },
                        new AnalysisItem { Name = "Publisher", Value = prog.Publisher, RegistryValue = "Publisher" },
                        new AnalysisItem { Name = "Install Date", Value = prog.InstallDate, RegistryValue = "InstallDate" }
                    }
                });
            }

            if (programs.Count == 0)
            {
                section.Items.Add(new AnalysisItem
                {
                    Name = "Info",
                    Value = "No installed programs found in this hive"
                });
            }

            return section;
        }

        #endregion

        #region Firewall Rules Analysis

        /// <summary>
        /// Get the enabled/disabled status for a firewall profile
        /// </summary>
        /// <param name="profileName">Profile name: DomainProfile, StandardProfile, or PublicProfile</param>
        /// <returns>True if firewall is enabled for this profile, false otherwise</returns>
        public bool IsFirewallProfileEnabled(string profileName)
        {
            string firewallBasePath = $@"{_currentControlSet}\Services\SharedAccess\Parameters\FirewallPolicy";
            var profileKey = _parser.GetKey($@"{firewallBasePath}\{profileName}");
            if (profileKey == null) return false;
            
            var enableFirewall = profileKey.Values.FirstOrDefault(v => v.ValueName == "EnableFirewall")?.ValueData?.ToString() ?? "";
            return enableFirewall == "1";
        }

        /// <summary>
        /// Get all firewall rules for a specific profile
        /// </summary>
        public List<FirewallRuleInfo> GetFirewallRulesForProfile(string profileName)
        {
            var rules = new List<FirewallRuleInfo>();
            string rulesPath = $@"{_currentControlSet}\Services\SharedAccess\Parameters\FirewallPolicy\FirewallRules";
            var rulesKey = _parser.GetKey(rulesPath);
            
            if (rulesKey == null) return rules;

            // Map profile name to the profile code used in rules
            var profileCode = profileName switch
            {
                "DomainProfile" => "Domain",
                "StandardProfile" => "Private",
                "PublicProfile" => "Public",
                _ => profileName
            };

            foreach (var value in rulesKey.Values)
            {
                var ruleData = value.ValueData?.ToString() ?? "";
                if (string.IsNullOrEmpty(ruleData)) continue;

                var rule = ParseFirewallRule(value.ValueName, ruleData);
                if (rule == null) continue;

                // Check if this rule applies to the selected profile
                if (rule.Profiles.Contains(profileCode, StringComparison.OrdinalIgnoreCase) ||
                    rule.Profiles.Contains("*") ||
                    string.IsNullOrEmpty(rule.Profiles))
                {
                    rule.RegistryPath = rulesPath;
                    rule.RegistryValueName = value.ValueName;
                    rules.Add(rule);
                }
            }

            return rules;
        }


        /// <summary>
        /// Parse a firewall rule string into a FirewallRuleInfo object
        /// Format: v2.31|Action=Allow|Active=TRUE|Dir=In|Protocol=6|Profile=Domain|...
        /// </summary>
        private FirewallRuleInfo? ParseFirewallRule(string ruleName, string ruleData)
        {
            try
            {
                var rule = new FirewallRuleInfo { RuleId = ruleName, RawData = ruleData };
                var localPortsList = new List<string>();
                var remotePortsList = new List<string>();
                
                // Split by | and parse key=value pairs
                var parts = ruleData.Split('|');
                foreach (var part in parts)
                {
                    var kvp = part.Split(new[] { '=' }, 2);
                    if (kvp.Length != 2) continue;

                    var key = kvp[0].Trim();
                    var val = kvp[1].Trim();

                    switch (key.ToLowerInvariant())
                    {
                        case "action":
                            rule.Action = val;
                            break;
                        case "active":
                            rule.IsActive = val.Equals("TRUE", StringComparison.OrdinalIgnoreCase);
                            break;
                        case "dir":
                            rule.Direction = val.Equals("In", StringComparison.OrdinalIgnoreCase) ? "Inbound" : "Outbound";
                            break;
                        case "protocol":
                            rule.Protocol = ParseProtocol(val);
                            break;
                        case "profile":
                            if (string.IsNullOrEmpty(rule.Profiles))
                                rule.Profiles = val;
                            else
                                rule.Profiles += "|" + val;
                            break;
                        case "lport":
                            localPortsList.Add(val);
                            break;
                        case "rport":
                            remotePortsList.Add(val);
                            break;
                        case "la4":
                        case "la6":
                            if (string.IsNullOrEmpty(rule.LocalAddresses))
                                rule.LocalAddresses = val;
                            else
                                rule.LocalAddresses += ", " + val;
                            break;
                        case "ra4":
                        case "ra6":
                            if (string.IsNullOrEmpty(rule.RemoteAddresses))
                                rule.RemoteAddresses = val;
                            else
                                rule.RemoteAddresses += ", " + val;
                            break;
                        case "app":
                            rule.Application = val;
                            break;
                        case "svc":
                            rule.Service = val;
                            break;
                        case "name":
                            rule.Name = val;
                            break;
                        case "desc":
                            rule.Description = val;
                            break;
                        case "embedctxt":
                            rule.EmbedContext = val;
                            break;
                        case "pfn":
                            rule.PackageFamilyName = val;
                            break;
                    }
                }

                // Format accumulated port lists (ranges for consecutive, commas otherwise)
                if (localPortsList.Count > 0)
                    rule.LocalPorts = FormatPortList(localPortsList);
                if (remotePortsList.Count > 0)
                    rule.RemotePorts = FormatPortList(remotePortsList);

                // Use the parsed Name= field if the registry value name is a GUID
                // and the Name= value is a real name (not an unresolvable MUI resource);
                // otherwise fall back to the registry value name (which is already readable).
                if (string.IsNullOrEmpty(rule.Name) || !ruleName.StartsWith('{') ||
                    rule.Name.StartsWith('@'))
                    rule.Name = ruleName;

                return rule;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Format a list of port values into a compact string.
        /// Consecutive numeric ports are collapsed into ranges (e.g. "23554-23556").
        /// Non-consecutive ports are comma-separated (e.g. "80, 443, 8080").
        /// Non-numeric port values (e.g. "RPC", "IPHTTPS") are kept as-is.
        /// </summary>
        private string FormatPortList(List<string> ports)
        {
            if (ports.Count == 0) return "";
            if (ports.Count == 1) return ports[0];

            // Separate numeric and non-numeric ports
            var numericPorts = new List<int>();
            var nonNumericPorts = new List<string>();

            foreach (var p in ports)
            {
                if (int.TryParse(p, out int num))
                    numericPorts.Add(num);
                else
                    nonNumericPorts.Add(p);
            }

            // If no numeric ports, just join non-numeric
            if (numericPorts.Count == 0)
                return string.Join(", ", nonNumericPorts.Distinct());

            // Sort and deduplicate numeric ports
            numericPorts = numericPorts.Distinct().OrderBy(n => n).ToList();

            // Build ranges for consecutive numbers
            var parts = new List<string>();
            int rangeStart = numericPorts[0];
            int rangeEnd = numericPorts[0];

            for (int i = 1; i < numericPorts.Count; i++)
            {
                if (numericPorts[i] == rangeEnd + 1)
                {
                    rangeEnd = numericPorts[i];
                }
                else
                {
                    parts.Add(rangeStart == rangeEnd ? rangeStart.ToString() : $"{rangeStart}-{rangeEnd}");
                    rangeStart = numericPorts[i];
                    rangeEnd = numericPorts[i];
                }
            }
            parts.Add(rangeStart == rangeEnd ? rangeStart.ToString() : $"{rangeStart}-{rangeEnd}");

            // Append non-numeric ports
            parts.AddRange(nonNumericPorts.Distinct());

            return string.Join(", ", parts);
        }

        /// <summary>
        /// Parse protocol number to name
        /// </summary>
        private string ParseProtocol(string protocolValue)
        {
            if (int.TryParse(protocolValue, out int protocolNum))
            {
                return protocolNum switch
                {
                    1 => "ICMP",
                    6 => "TCP",
                    17 => "UDP",
                    47 => "GRE",
                    58 => "ICMPv6",
                    256 => "Any",
                    _ => $"Protocol {protocolNum}"
                };
            }
            return protocolValue;
        }

        #endregion

        #region Server Roles & Features

        /// <summary>
        /// Checks whether the SOFTWARE hive contains Windows Server roles/features data
        /// </summary>
        public bool HasServerRolesAndFeatures()
        {
            var key = _parser.GetKey(@"Microsoft\ServerManager\ServicingStorage\ServerComponentCache");
            return key?.SubKeys != null && key.SubKeys.Count > 0;
        }

        /// <summary>
        /// Checks whether this SOFTWARE hive contains Scheduled Tasks data.
        /// </summary>
        public bool HasScheduledTasks()
        {
            var key = _parser.GetKey(@"Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree");
            return key?.SubKeys != null && key.SubKeys.Count > 0;
        }

        /// <summary>
        /// Checks whether this SYSTEM hive contains boot configuration data.
        /// </summary>
        public bool HasBootConfigurations()
        {
            if (_parser.CurrentHiveType != OfflineRegistryParser.HiveType.SYSTEM)
                return false;
            var controlKey = _parser.GetKey($@"{_currentControlSet}\Control");
            return controlKey != null;
        }

        /// <summary>
        /// Extracts BIOS Mode, BIOS Version/Date, VM Generation, and TPM items
        /// for the Computer Information section.
        /// </summary>
        private List<AnalysisItem> GetBiosAndTpmItems()
        {
            var items = new List<AnalysisItem>();

            // ── BIOS Mode & Version ──────────────────────────────────────
            var isUefi = false;
            var biosVersion = "";
            var biosVendor = "";
            var biosReleaseDate = "";
            var systemVersion = "";
            var systemManufacturer = "";
            var systemProductName = "";
            var hwConfigPath = "HardwareConfig";
            var hwConfigKey = _parser.GetKey(hwConfigPath);
            if (hwConfigKey != null)
            {
                var lastConfig = hwConfigKey.Values
                    .FirstOrDefault(v => v.ValueName == "LastConfig")?.ValueData?.ToString() ?? "";
                if (!string.IsNullOrEmpty(lastConfig))
                {
                    var hwGuidPath = $@"HardwareConfig\{lastConfig}";
                    var hwGuidKey = _parser.GetKey(hwGuidPath);
                    if (hwGuidKey != null)
                    {
                        biosVersion = hwGuidKey.Values
                            .FirstOrDefault(v => v.ValueName == "BIOSVersion")?.ValueData?.ToString() ?? "";
                        biosVendor = hwGuidKey.Values
                            .FirstOrDefault(v => v.ValueName == "BIOSVendor")?.ValueData?.ToString() ?? "";
                        biosReleaseDate = hwGuidKey.Values
                            .FirstOrDefault(v => v.ValueName == "BIOSReleaseDate")?.ValueData?.ToString() ?? "";
                        systemVersion = hwGuidKey.Values
                            .FirstOrDefault(v => v.ValueName == "SystemVersion")?.ValueData?.ToString() ?? "";
                        systemManufacturer = hwGuidKey.Values
                            .FirstOrDefault(v => v.ValueName == "SystemManufacturer")?.ValueData?.ToString() ?? "";
                        systemProductName = hwGuidKey.Values
                            .FirstOrDefault(v => v.ValueName == "SystemProductName")?.ValueData?.ToString() ?? "";
                        var systemBiosVersion = hwGuidKey.Values
                            .FirstOrDefault(v => v.ValueName == "SystemBiosVersion")?.ValueData?.ToString() ?? "";

                        isUefi = biosVersion.Contains("UEFI", StringComparison.OrdinalIgnoreCase)
                              || systemVersion.Contains("UEFI", StringComparison.OrdinalIgnoreCase)
                              || systemBiosVersion.Contains("UEFI", StringComparison.OrdinalIgnoreCase);

                        hwConfigPath = hwGuidPath;
                    }
                }
            }

            var biosMode = isUefi ? "UEFI" : "Legacy";
            items.Add(new AnalysisItem
            {
                Name = "BIOS Mode",
                Value = biosMode,
                RegistryPath = hwConfigPath,
                RegistryValue = isUefi
                    ? $"BIOSVersion contains UEFI indicator — {biosVersion}"
                    : $"No UEFI indicator found in BIOS strings — {(string.IsNullOrEmpty(biosVersion) ? "HardwareConfig not available" : biosVersion)}"
            });

            var biosParts = new List<string>();
            if (!string.IsNullOrEmpty(biosVendor)) biosParts.Add(biosVendor);
            if (!string.IsNullOrEmpty(biosVersion)) biosParts.Add(biosVersion);
            if (!string.IsNullOrEmpty(biosReleaseDate)) biosParts.Add(biosReleaseDate);
            if (biosParts.Count > 0)
            {
                var biosLine = string.Join(" ", biosParts.Take(2));
                if (biosParts.Count == 3)
                    biosLine += $", {biosParts[2]}";

                items.Add(new AnalysisItem
                {
                    Name = "BIOS Version/Date",
                    Value = biosLine,
                    RegistryPath = hwConfigPath,
                    RegistryValue = $"BIOSVendor={biosVendor}, BIOSVersion={biosVersion}, BIOSReleaseDate={biosReleaseDate}"
                });
            }

            // ── VM Generation (Hyper-V) ──────────────────────────────────

            var isHyperV = systemManufacturer.Contains("Microsoft", StringComparison.OrdinalIgnoreCase)
                        && systemProductName.Contains("Virtual Machine", StringComparison.OrdinalIgnoreCase);

            if (isHyperV)
            {
                var vmGen = isUefi ? "Generation 2" : "Generation 1";
                var vmGenDetail = isUefi
                    ? "UEFI firmware — Hyper-V Generation 2 VM"
                    : "Legacy BIOS firmware — Hyper-V Generation 1 VM";
                items.Add(new AnalysisItem
                {
                    Name = "VM Generation",
                    Value = vmGen,
                    RegistryPath = hwConfigPath,
                    RegistryValue = vmGenDetail
                });
            }

            // ── TPM Hardware Detection ───────────────────────────────────
            var tpmClassGuid = "{d94ee5d8-d189-4994-83d2-f68d7d41b0e6}";
            var tpmClassPath = $@"{_currentControlSet}\Control\Class\{tpmClassGuid}\0000";
            var tpmClassKey = _parser.GetKey(tpmClassPath);

            if (tpmClassKey != null)
            {
                var infSection = tpmClassKey.Values
                    .FirstOrDefault(v => v.ValueName == "InfSection")?.ValueData?.ToString() ?? "";
                var driverDesc = tpmClassKey.Values
                    .FirstOrDefault(v => v.ValueName == "DriverDesc")?.ValueData?.ToString() ?? "";

                var tpmVersion = infSection.ToUpperInvariant().Contains("TPM2") ? "2.0" : "1.2";

                var dosDeviceName = "";
                var isVirtual = false;
                var compatibleIds = "";
                var enumPath = "";

                string[] tpmAcpiIds = { "MSFT1001", "MSFT0101", "INTC0102", "AMD0040", "BCM0101", "PNP0C31" };
                foreach (var acpiId in tpmAcpiIds)
                {
                    var acpiPath = $@"{_currentControlSet}\Enum\ACPI\{acpiId}";
                    var acpiKey = _parser.GetKey(acpiPath);
                    if (acpiKey?.SubKeys?.Count > 0)
                    {
                        var instanceKey = acpiKey.SubKeys[0];
                        enumPath = $@"{acpiPath}\{instanceKey.KeyName}";

                        compatibleIds = instanceKey.Values
                            .FirstOrDefault(v => v.ValueName == "CompatibleIDs")?.ValueData?.ToString() ?? "";
                        isVirtual = compatibleIds.ToUpperInvariant().Contains("VTPM");

                        var devParamsPath = $@"{enumPath}\Device Parameters";
                        var devParamsKey = _parser.GetKey(devParamsPath);
                        if (devParamsKey != null)
                        {
                            dosDeviceName = devParamsKey.Values
                                .FirstOrDefault(v => v.ValueName == "DosDeviceName")?.ValueData?.ToString() ?? "";
                        }

                        break;
                    }
                }

                var tpmParts = new List<string>();
                tpmParts.Add($"TPM {tpmVersion}");
                if (isVirtual) tpmParts[0] += " (Virtual)";

                var friendlyName = !string.IsNullOrEmpty(dosDeviceName) ? dosDeviceName : driverDesc;
                if (!string.IsNullOrEmpty(friendlyName))
                    tpmParts.Add(friendlyName);

                var tpmDisplayValue = string.Join(" — ", tpmParts);

                items.Add(new AnalysisItem
                {
                    Name = "TPM",
                    Value = tpmDisplayValue,
                    RegistryPath = !string.IsNullOrEmpty(enumPath) ? enumPath : tpmClassPath,
                    RegistryValue = $"InfSection={infSection}, DriverDesc={driverDesc}, CompatibleIDs={compatibleIds}, DosDeviceName={dosDeviceName}"
                });
            }
            else
            {
                items.Add(new AnalysisItem
                {
                    Name = "TPM",
                    Value = "Not Detected",
                    RegistryPath = tpmClassPath,
                    RegistryValue = "No TPM device class instance found (no hardware TPM installed)"
                });
            }

            return items;
        }

        /// <summary>
        /// Gets boot configuration analysis from the SYSTEM hive.
        /// Includes boot devices, start options, boot status, and BitLocker integration.
        /// BIOS Mode, BIOS Version/Date, VM Generation, and TPM are in Computer Information.
        /// </summary>
        public AnalysisSection GetBootConfigurationAnalysis()
        {
            var section = new AnalysisSection { Title = "\U0001f4bb Boot Configurations" };

            if (_parser.CurrentHiveType != OfflineRegistryParser.HiveType.SYSTEM)
            {
                section.Items.Add(new AnalysisItem
                {
                    Name = "Notice",
                    Value = "This information requires the SYSTEM hive to be loaded",
                    RegistryPath = "",
                    RegistryValue = "Load SYSTEM hive to view boot configuration"
                });
                return section;
            }

            var controlPath = $@"{_currentControlSet}\Control";

            // ── 1. Boot Device Paths ────────────────────────────────────
            var systemBootDevice = GetValue(controlPath, "SystemBootDevice");
            var firmwareBootDevice = GetValue(controlPath, "FirmwareBootDevice");

            // Resolve ARC paths to drive letters (MBR only; GPT falls back to raw paths)
            var driveLetterMap = ResolveBootDeviceDriveLetters(
                systemBootDevice ?? "", firmwareBootDevice ?? "");

            if (!string.IsNullOrEmpty(systemBootDevice))
            {
                var systemDisplay = driveLetterMap.TryGetValue("SystemBootDevice", out var sysLetter)
                    ? $"{systemBootDevice} ({sysLetter})"
                    : systemBootDevice;

                section.Items.Add(new AnalysisItem
                {
                    Name = "System Boot Device",
                    Value = systemDisplay,
                    RegistryPath = controlPath,
                    RegistryValue = $"SystemBootDevice = {systemBootDevice}"
                });
            }

            if (!string.IsNullOrEmpty(firmwareBootDevice))
            {
                var firmwareDisplay = driveLetterMap.TryGetValue("FirmwareBootDevice", out var fwLetter)
                    ? $"{firmwareBootDevice} ({fwLetter})"
                    : firmwareBootDevice;

                section.Items.Add(new AnalysisItem
                {
                    Name = "Firmware Boot Device",
                    Value = firmwareDisplay,
                    RegistryPath = controlPath,
                    RegistryValue = $"FirmwareBootDevice = {firmwareBootDevice}"
                });
            }

            // ── 2. System Start Options ──────────────────────────────────
            var startOptions = GetValue(controlPath, "SystemStartOptions");
            if (!string.IsNullOrEmpty(startOptions))
            {
                var hasWarningOptions = startOptions.Contains("SAFEBOOT", StringComparison.OrdinalIgnoreCase) ||
                                        startOptions.Contains("DEREPAIR", StringComparison.OrdinalIgnoreCase);
                section.Items.Add(new AnalysisItem
                {
                    Name = "System Start Options",
                    Value = startOptions,
                    IsWarning = hasWarningOptions,
                    RegistryPath = controlPath,
                    RegistryValue = $"SystemStartOptions = {startOptions}"
                });

                // Parse individual flags into sub-items with descriptions
                ParseBootStartOptions(section, controlPath, startOptions);
            }

            // ── 3. Boot Status ───────────────────────────────────────────
            var lastBootSucceeded = GetValue(controlPath, "LastBootSucceeded");
            if (!string.IsNullOrEmpty(lastBootSucceeded))
            {
                var succeeded = lastBootSucceeded == "1" || lastBootSucceeded.StartsWith("1 ");
                section.Items.Add(new AnalysisItem
                {
                    Name = "Last Boot Succeeded",
                    Value = succeeded ? "Yes" : "No",
                    RegistryPath = controlPath,
                    RegistryValue = $"LastBootSucceeded = {lastBootSucceeded}"
                });
            }

            var lastBootShutdown = GetValue(controlPath, "LastBootShutdown");
            if (!string.IsNullOrEmpty(lastBootShutdown))
            {
                var cleanShutdown = lastBootShutdown == "1" || lastBootShutdown.StartsWith("1 ");
                section.Items.Add(new AnalysisItem
                {
                    Name = "Last Boot Shutdown",
                    Value = cleanShutdown ? "Clean" : $"Abnormal ({lastBootShutdown})",
                    IsWarning = !cleanShutdown,
                    RegistryPath = controlPath,
                    RegistryValue = $"LastBootShutdown = {lastBootShutdown}"
                });
            }

            var dirtyShutdownCount = GetValue(controlPath, "DirtyShutdownCount");
            if (!string.IsNullOrEmpty(dirtyShutdownCount))
            {
                section.Items.Add(new AnalysisItem
                {
                    Name = "Dirty Shutdown Count",
                    Value = dirtyShutdownCount,
                    RegistryPath = controlPath,
                    RegistryValue = $"DirtyShutdownCount = {dirtyShutdownCount}"
                });
            }

            var bootDriverFlags = GetValue(controlPath, "BootDriverFlags");
            if (!string.IsNullOrEmpty(bootDriverFlags))
            {
                section.Items.Add(new AnalysisItem
                {
                    Name = "Boot Driver Flags",
                    Value = bootDriverFlags,
                    RegistryPath = controlPath,
                    RegistryValue = $"BootDriverFlags = {bootDriverFlags}"
                });
            }

            // ── 4. BitLocker boot integration
            var fvevolPath = $@"{_currentControlSet}\Services\fvevol";
            var fvevolKey = _parser.GetKey(fvevolPath);
            if (fvevolKey != null)
            {
                var fveStart = fvevolKey.Values
                    .FirstOrDefault(v => v.ValueName == "Start")?.ValueData?.ToString() ?? "";
                var fveStartDesc = fveStart switch
                {
                    "0" or "0 (0x0)" => "Boot (0) — loads at boot time",
                    "1" or "1 (0x1)" => "System (1)",
                    "3" or "3 (0x3)" => "Manual (3)",
                    "4" or "4 (0x4)" => "Disabled (4)",
                    _ => fveStart
                };
                section.Items.Add(new AnalysisItem
                {
                    Name = "BitLocker Volume Driver (fvevol)",
                    Value = fveStartDesc,
                    RegistryPath = fvevolPath,
                    RegistryValue = $"Start = {fveStart}"
                });
            }

            var bitlockerPath = $@"{_currentControlSet}\Control\BitLocker";
            var bitlockerKey = _parser.GetKey(bitlockerPath);
            if (bitlockerKey != null)
            {
                section.Items.Add(new AnalysisItem
                {
                    Name = "BitLocker Configuration",
                    Value = "Present",
                    RegistryPath = bitlockerPath,
                    RegistryValue = "BitLocker configuration key exists"
                });
            }

            // Fallback if nothing was found
            if (section.Items.Count == 0)
            {
                section.Items.Add(new AnalysisItem
                {
                    Name = "Status",
                    Value = "No boot configuration data found in this SYSTEM hive",
                    RegistryPath = controlPath,
                    RegistryValue = "Control key may be empty or malformed"
                });
            }

            return section;
        }

        /// <summary>
        /// Resolves boot device ARC paths to drive letters using MountedDevices cross-reference.
        /// For MBR disks: identifies the Windows drive letter from DriverDatabase\SystemRoot,
        /// then searches MountedDevices for other partitions on the same disk.
        /// For GPT disks: returns empty dictionary (caller should fall back to raw ARC paths).
        /// Returns: dictionary mapping "SystemBootDevice" and/or "FirmwareBootDevice" to drive letter annotations.
        /// </summary>
        private Dictionary<string, string> ResolveBootDeviceDriveLetters(string systemBootDevice, string firmwareBootDevice)
        {
            var result = new Dictionary<string, string>();

            // 1. Get the Windows drive letter from DriverDatabase\SystemRoot (e.g., "C:\Windows" → "C:")
            var systemRoot = GetValue(@"DriverDatabase", "SystemRoot");
            if (string.IsNullOrEmpty(systemRoot) || systemRoot.Length < 2 || systemRoot[1] != ':')
                return result;

            var windowsDriveLetter = systemRoot.Substring(0, 2); // e.g., "C:"

            // 2. Look up the Windows drive in MountedDevices to determine disk type
            var mountedDevicesKey = _parser.GetKey("MountedDevices");
            if (mountedDevicesKey?.Values == null)
                return result;

            var windowsValueName = $@"\DosDevices\{windowsDriveLetter}";
            var windowsEntry = mountedDevicesKey.Values
                .FirstOrDefault(v => string.Equals(v.ValueName, windowsValueName, StringComparison.OrdinalIgnoreCase));

            byte[]? windowsData = windowsEntry?.ValueDataRaw;
            if (windowsData == null || windowsData.Length != 12)
                return result; // Not MBR (GPT or missing) — fall back to raw paths

            // 3. MBR system: extract the boot disk signature from the Windows drive entry
            uint bootDiskSig = BitConverter.ToUInt32(windowsData, 0);

            // SystemBootDevice partition = the Windows partition
            if (!string.IsNullOrEmpty(systemBootDevice))
                result["SystemBootDevice"] = windowsDriveLetter;

            // 4. For FirmwareBootDevice: search for other drive letters on the same disk
            if (!string.IsNullOrEmpty(firmwareBootDevice) && firmwareBootDevice != systemBootDevice)
            {
                string? firmwareLetter = null;

                // Parse partition number from the firmware boot device ARC path
                var firmwarePartMatch = System.Text.RegularExpressions.Regex.Match(
                    firmwareBootDevice, @"partition\((\d+)\)");
                var systemPartMatch = System.Text.RegularExpressions.Regex.Match(
                    systemBootDevice ?? "", @"partition\((\d+)\)");

                if (firmwarePartMatch.Success && systemPartMatch.Success)
                {
                    int firmwarePartNum = int.Parse(firmwarePartMatch.Groups[1].Value);
                    int systemPartNum = int.Parse(systemPartMatch.Groups[1].Value);

                    if (firmwarePartNum != systemPartNum)
                    {
                        // Search all MountedDevices drive letter entries on the same disk
                        // to find which (if any) maps to the firmware partition
                        var sameDiskEntries = new List<(long offset, string driveLetter)>();

                        foreach (var val in mountedDevicesKey.Values)
                        {
                            if (!val.ValueName.StartsWith(@"\DosDevices\", StringComparison.OrdinalIgnoreCase))
                                continue;

                            byte[]? data = val.ValueDataRaw;
                            if (data == null || data.Length != 12)
                                continue;

                            uint sig = BitConverter.ToUInt32(data, 0);
                            if (sig != bootDiskSig)
                                continue;

                            long offset = BitConverter.ToInt64(data, 4);
                            string letter = val.ValueName.Substring(@"\DosDevices\".Length);
                            sameDiskEntries.Add((offset, letter));
                        }

                        // Sort by offset to determine partition ordering
                        sameDiskEntries.Sort((a, b) => a.offset.CompareTo(b.offset));

                        // Assign 1-indexed partition numbers based on sorted offset
                        // But there may be gaps (partitions without drive letters).
                        // We know systemPartNum maps to windowsDriveLetter.
                        // Find where windowsDriveLetter falls in the sorted list.
                        int windowsIndex = sameDiskEntries.FindIndex(e => e.driveLetter == windowsDriveLetter);
                        if (windowsIndex >= 0)
                        {
                            // The number of hidden partitions before the first lettered partition
                            // = systemPartNum - (number of lettered partitions up to and including Windows)
                            int hiddenBefore = systemPartNum - (windowsIndex + 1);

                            // Now check if firmwarePartNum maps to a lettered partition
                            int firmwareIndex = firmwarePartNum - hiddenBefore - 1; // 0-indexed into sameDiskEntries
                            if (firmwareIndex >= 0 && firmwareIndex < sameDiskEntries.Count
                                && firmwareIndex != windowsIndex)
                            {
                                firmwareLetter = sameDiskEntries[firmwareIndex].driveLetter;
                            }
                        }
                    }
                }

                result["FirmwareBootDevice"] = firmwareLetter ?? "no drive letter";
            }
            else if (!string.IsNullOrEmpty(firmwareBootDevice))
            {
                // Same partition as SystemBootDevice
                result["FirmwareBootDevice"] = windowsDriveLetter;
            }

            return result;
        }

        /// <summary>
        /// Parses individual flags from the SystemStartOptions string and adds them as AnalysisItems.
        /// </summary>
        private void ParseBootStartOptions(AnalysisSection section, string controlPath, string startOptions)
        {
            var options = startOptions.Split(' ', StringSplitOptions.RemoveEmptyEntries);

            foreach (var option in options)
            {
                string name;
                string description;
                bool isWarning = false;

                if (option.StartsWith("NOEXECUTE=", StringComparison.OrdinalIgnoreCase))
                {
                    var policy = option.Substring("NOEXECUTE=".Length);
                    name = "  DEP Policy (NOEXECUTE)";
                    description = policy.ToUpperInvariant() switch
                    {
                        "OPTIN" => "OptIn — DEP for Windows system components only",
                        "OPTOUT" => "OptOut — DEP for all processes except exclusion list",
                        "ALWAYSON" => "AlwaysOn — DEP for all processes, no exceptions",
                        "ALWAYSOFF" => "AlwaysOff — DEP disabled",
                        _ => policy
                    };
                }
                else if (option.Equals("REDIRECT", StringComparison.OrdinalIgnoreCase))
                {
                    name = "  Serial Console Redirect";
                    description = "Enabled — EMS/serial console redirection active";
                }
                else if (option.StartsWith("FVEBOOT=", StringComparison.OrdinalIgnoreCase))
                {
                    var offset = option.Substring("FVEBOOT=".Length);
                    name = "  BitLocker Boot (FVEBOOT)";
                    description = $"Boot partition offset: {offset}";
                }
                else if (option.Equals("NOVGA", StringComparison.OrdinalIgnoreCase))
                {
                    name = "  VGA Mode";
                    description = "Disabled (NOVGA) — no VGA-compatible display driver";
                }
                else if (option.StartsWith("SOS", StringComparison.OrdinalIgnoreCase))
                {
                    name = "  Verbose Boot";
                    description = "Enabled (SOS) — driver names displayed during boot";
                }
                else if (option.Equals("SAFEBOOT", StringComparison.OrdinalIgnoreCase) ||
                         option.StartsWith("SAFEBOOT:", StringComparison.OrdinalIgnoreCase))
                {
                    name = "  Safe Boot";
                    description = option.Contains(':')
                        ? $"Mode: {option.Substring(option.IndexOf(':') + 1)}"
                        : "Enabled";
                    isWarning = true;
                }
                else if (option.Equals("DEREPAIR", StringComparison.OrdinalIgnoreCase))
                {
                    name = "  Disaster Recovery Repair";
                    description = "Enabled (DEREPAIR) — system booted in repair mode";
                    isWarning = true;
                }
                else
                {
                    name = $"  {option}";
                    description = "(boot option flag)";
                }

                section.Items.Add(new AnalysisItem
                {
                    Name = name,
                    Value = description,
                    IsWarning = isWarning,
                    RegistryPath = controlPath,
                    RegistryValue = $"Parsed from SystemStartOptions: {option}"
                });
            }
        }

        #region Certificate Stores

        /// <summary>
        /// Store name mapping from registry key names to certmgr.msc friendly names.
        /// </summary>
        private static readonly Dictionary<string, string> CertStoreNameMap = new(StringComparer.OrdinalIgnoreCase)
        {
            ["MY"] = "Personal",
            ["ROOT"] = "Trusted Root Certification Authorities",
            ["CA"] = "Intermediate Certification Authorities",
            ["AuthRoot"] = "Third-Party Root Certification Authorities",
            ["Disallowed"] = "Untrusted Certificates",
            ["TrustedPublisher"] = "Trusted Publishers",
            ["TrustedPeople"] = "Trusted People",
            ["trust"] = "Enterprise Trust",
            ["Remote Desktop"] = "Remote Desktop",
            ["SmartCardRoot"] = "Smart Card Trusted Roots",
            ["AddressBook"] = "Other People",
            ["REQUEST"] = "Certificate Enrollment Requests",
            ["ClientAuthIssuers"] = "Client Authentication Issuers",
            ["TrustedDevices"] = "Trusted Devices"
        };

        /// <summary>
        /// Store display order matching certmgr.msc default layout.
        /// Stores not in this list sort alphabetically after the known ones.
        /// </summary>
        private static readonly Dictionary<string, int> CertStoreOrder = new(StringComparer.OrdinalIgnoreCase)
        {
            ["MY"] = 0,
            ["ROOT"] = 1,
            ["trust"] = 2,
            ["CA"] = 3,
            ["TrustedPublisher"] = 4,
            ["Disallowed"] = 5,
            ["AuthRoot"] = 6,
            ["TrustedPeople"] = 7,
            ["ClientAuthIssuers"] = 8,
            ["Remote Desktop"] = 10,
            ["SmartCardRoot"] = 11,
            ["TrustedDevices"] = 12,
            ["AddressBook"] = 13,
            ["REQUEST"] = 14
        };

        /// <summary>
        /// Checks whether this SOFTWARE hive contains any certificate stores.
        /// </summary>
        public bool HasCertificateStores()
        {
            var key = _parser.GetKey(@"Microsoft\SystemCertificates");
            return key?.SubKeys != null && key.SubKeys.Count > 0;
        }

        /// <summary>
        /// Extracts all certificate stores and their certificates from the SOFTWARE hive.
        /// Returns a list of CertificateStoreInfo, each containing the store name and its certificates.
        /// </summary>
        public List<CertificateStoreInfo> GetCertificateStoresData()
        {
            var stores = new List<CertificateStoreInfo>();
            var rootKey = _parser.GetKey(@"Microsoft\SystemCertificates");
            if (rootKey?.SubKeys == null) return stores;

            foreach (var storeKey in rootKey.SubKeys)
            {
                var storeName = storeKey.KeyName;
                var friendlyName = CertStoreNameMap.TryGetValue(storeName, out var mapped) ? mapped : storeName;

                var storeInfo = new CertificateStoreInfo
                {
                    RegistryName = storeName,
                    FriendlyName = friendlyName,
                    RegistryPath = $@"Microsoft\SystemCertificates\{storeName}"
                };

                // Look for Certificates subkey
                var certsKey = _parser.GetKey($@"Microsoft\SystemCertificates\{storeName}\Certificates");
                if (certsKey?.SubKeys != null)
                {
                    foreach (var certKey in certsKey.SubKeys)
                    {
                        var thumbprint = certKey.KeyName;
                        var certInfo = new CertificateInfo
                        {
                            Thumbprint = thumbprint,
                            RegistryPath = $@"Microsoft\SystemCertificates\{storeName}\Certificates\{thumbprint}"
                        };

                        // Find the Blob value
                        if (certKey.Values != null)
                        {
                            var blobValue = certKey.Values.FirstOrDefault(v =>
                                string.Equals(v.ValueName, "Blob", StringComparison.OrdinalIgnoreCase));
                            if (blobValue != null)
                            {
                                try
                                {
                                    var blobData = blobValue.ValueDataRaw;
                                    if (blobData != null && blobData.Length > 0)
                                    {
                                        DecodeCertificateBlob(blobData, certInfo);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine($"Error decoding cert blob for {thumbprint}: {ex.Message}");
                                }
                            }
                        }

                        // Fallback: if subject is still empty, use thumbprint
                        if (string.IsNullOrEmpty(certInfo.Subject))
                            certInfo.DisplayName = thumbprint;
                        else
                            certInfo.DisplayName = ExtractCN(certInfo.Subject) ?? certInfo.Subject;

                        storeInfo.Certificates.Add(certInfo);
                    }
                }

                stores.Add(storeInfo);
            }

            // Sort: certmgr.msc positional order for known stores, then alphabetical for unknown.
            // Empty stores sort to the bottom within each group.
            const int unknownOrder = 1000;
            stores.Sort((a, b) =>
            {
                bool aEmpty = a.Certificates.Count == 0;
                bool bEmpty = b.Certificates.Count == 0;
                if (aEmpty != bEmpty) return aEmpty ? 1 : -1;

                int aOrder = CertStoreOrder.TryGetValue(a.RegistryName, out var ao) ? ao : unknownOrder;
                int bOrder = CertStoreOrder.TryGetValue(b.RegistryName, out var bo) ? bo : unknownOrder;
                if (aOrder != bOrder) return aOrder.CompareTo(bOrder);
                // Both unknown (or same position): sort alphabetically by friendly name
                return string.Compare(a.FriendlyName, b.FriendlyName, StringComparison.OrdinalIgnoreCase);
            });

            return stores;
        }

        /// <summary>
        /// Decodes a certificate blob (sequence of property entries) and populates the CertificateInfo.
        /// Blob format: repeating { DWORD propId, DWORD reserved(=1), DWORD cbData, BYTE data[cbData] }
        /// </summary>
        private static void DecodeCertificateBlob(byte[] blob, CertificateInfo certInfo)
        {
            int offset = 0;
            while (offset + 12 <= blob.Length)
            {
                uint propId = BitConverter.ToUInt32(blob, offset);
                // uint reserved = BitConverter.ToUInt32(blob, offset + 4); // always 1
                uint cbData = BitConverter.ToUInt32(blob, offset + 8);
                offset += 12;

                if (cbData > (uint)(blob.Length - offset))
                    break; // malformed

                var propData = new byte[cbData];
                Array.Copy(blob, offset, propData, 0, (int)cbData);
                offset += (int)cbData;

                switch (propId)
                {
                    case 0x03: // SHA-1 hash (thumbprint)
                        if (propData.Length == 20)
                            certInfo.Sha1Hash = BitConverter.ToString(propData).Replace("-", "");
                        break;

                    case 0x0B: // Friendly Name (UTF-16LE, null-terminated)
                        certInfo.FriendlyName = Encoding.Unicode.GetString(propData).TrimEnd('\0');
                        break;

                    case 0x02: // Key Provider Info (CRYPT_KEY_PROV_INFO serialized)
                        DecodeKeyProviderInfo(propData, certInfo);
                        break;

                    case 0x20: // DER-encoded X.509 certificate
                        DecodeDerCertificate(propData, certInfo);
                        break;
                }
            }
        }

        /// <summary>
        /// Parses a DER-encoded certificate using X509Certificate2 to extract standard fields.
        /// </summary>
        private static void DecodeDerCertificate(byte[] derData, CertificateInfo certInfo)
        {
            try
            {
                using var cert = new X509Certificate2(derData);
                certInfo.Subject = cert.GetNameInfo(System.Security.Cryptography.X509Certificates.X509NameType.SimpleName, false);
                certInfo.Issuer = cert.GetNameInfo(System.Security.Cryptography.X509Certificates.X509NameType.SimpleName, true);
                certInfo.SerialNumber = cert.SerialNumber;
                certInfo.ValidFrom = cert.NotBefore;
                certInfo.ValidTo = cert.NotAfter;
                certInfo.SignatureAlgorithm = cert.SignatureAlgorithm.FriendlyName ?? cert.SignatureAlgorithm.Value ?? "";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error parsing DER certificate: {ex.Message}");
            }
        }

        /// <summary>
        /// Decodes the serialized CRYPT_KEY_PROV_INFO structure to extract CSP and key container names.
        /// Layout: fixed header (0x1C bytes), then container name (UTF-16LE) and provider name (UTF-16LE) inline.
        /// </summary>
        private static void DecodeKeyProviderInfo(byte[] data, CertificateInfo certInfo)
        {
            try
            {
                if (data.Length < 0x1C) return;

                // The serialized structure stores string offsets at the start:
                // offset 0x00: DWORD containerNameOffset (typically 0x1C)
                // offset 0x04: DWORD provNameOffset
                uint containerOffset = BitConverter.ToUInt32(data, 0);
                uint providerOffset = BitConverter.ToUInt32(data, 4);

                if (containerOffset < data.Length)
                {
                    certInfo.KeyContainer = ReadNullTerminatedUnicode(data, (int)containerOffset);
                }

                if (providerOffset < data.Length)
                {
                    certInfo.KeyProvider = ReadNullTerminatedUnicode(data, (int)providerOffset);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error parsing Key Provider Info: {ex.Message}");
            }
        }

        /// <summary>
        /// Reads a null-terminated UTF-16LE string starting at the given offset.
        /// </summary>
        private static string ReadNullTerminatedUnicode(byte[] data, int offset)
        {
            var sb = new StringBuilder();
            while (offset + 1 < data.Length)
            {
                char c = (char)(data[offset] | (data[offset + 1] << 8));
                if (c == '\0') break;
                sb.Append(c);
                offset += 2;
            }
            return sb.ToString();
        }

        /// <summary>
        /// Extracts the CN (Common Name) from an X.500 distinguished name string.
        /// Returns null if no CN is found.
        /// </summary>
        private static string? ExtractCN(string distinguishedName)
        {
            // Try CN= first, then OU=, then O=, then DC= (matches certmgr.msc behavior)
            string[] prefixes = { "CN=", "OU=", "O=", "DC=" };
            foreach (var prefix in prefixes)
            {
                var value = ExtractDNField(distinguishedName, prefix);
                if (value != null) return value;
            }
            return null;
        }

        private static string? ExtractDNField(string distinguishedName, string prefix)
        {
            int idx = distinguishedName.IndexOf(prefix, StringComparison.OrdinalIgnoreCase);
            if (idx < 0) return null;

            int start = idx + prefix.Length;
            if (start >= distinguishedName.Length) return null;

            // Handle quoted values
            if (distinguishedName[start] == '"')
            {
                int closeQuote = distinguishedName.IndexOf('"', start + 1);
                return closeQuote > start ? distinguishedName.Substring(start + 1, closeQuote - start - 1) : null;
            }

            // Unquoted: read until next comma or end
            int end = distinguishedName.IndexOf(", ", start, StringComparison.Ordinal);
            return end < 0 ? distinguishedName.Substring(start) : distinguishedName.Substring(start, end - start);
        }

        #endregion Certificate Stores

        /// <summary>
        /// Extracts all server roles and features from ServerComponentCache
        /// </summary>
        public List<RoleFeatureItem> GetRolesAndFeaturesData()
        {
            var results = new List<RoleFeatureItem>();
            var cacheKey = _parser.GetKey(@"Microsoft\ServerManager\ServicingStorage\ServerComponentCache");
            if (cacheKey?.SubKeys == null) return results;

            foreach (var subKey in cacheKey.SubKeys)
            {
                var item = new RoleFeatureItem
                {
                    KeyName = subKey.KeyName,
                    RegistryPath = $@"Microsoft\ServerManager\ServicingStorage\ServerComponentCache\{subKey.KeyName}"
                };

                if (subKey.Values != null)
                {
                    foreach (var val in subKey.Values)
                    {
                        var valName = val.ValueName ?? "";
                        var valData = val.ValueData?.ToString() ?? "";

                        switch (valName)
                        {
                            case "DisplayName":
                                item.DisplayName = valData;
                                break;
                            case "Description":
                                item.Description = valData;
                                break;
                            case "ServerComponentType":
                                if (int.TryParse(valData, out int compType))
                                    item.ServerComponentType = compType;
                                break;
                            case "InstallState":
                                if (int.TryParse(valData, out int installState))
                                    item.InstallState = installState;
                                break;
                            case "ParentName":
                                item.ParentName = valData;
                                break;
                            case "NumericId":
                                if (int.TryParse(valData, out int numId))
                                    item.NumericId = numId;
                                break;
                            case "MajorVersion":
                                if (int.TryParse(valData, out int major))
                                    item.MajorVersion = major;
                                break;
                            case "MinorVersion":
                                if (int.TryParse(valData, out int minor))
                                    item.MinorVersion = minor;
                                break;
                            case "SystemServices":
                                item.SystemServices = valData;
                                break;
                            case "NonAncestorDependencies":
                                item.Dependencies = valData;
                                break;
                        }
                    }
                }

                // Fall back to key name if no DisplayName
                if (string.IsNullOrEmpty(item.DisplayName))
                    item.DisplayName = subKey.KeyName;

                results.Add(item);
            }

            return results;
        }

        /// <summary>
        /// Get scheduled tasks from TaskCache as a hierarchical section.
        /// Tree entries map task names to GUIDs; Tasks entries hold the decoded details.
        /// </summary>
        public AnalysisSection GetScheduledTasksAnalysis()
        {
            var section = new AnalysisSection { Title = "📅 Scheduled Tasks" };

            try
            {
                string treePath = @"Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree";
                string tasksPath = @"Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks";
                var treeKey = GetCachedKey(treePath);
                if (treeKey == null)
                {
                    section.Items.Add(new AnalysisItem
                    {
                        Name = "No Data",
                        Value = "TaskCache\\Tree not found in this hive"
                    });
                    return section;
                }

                // Build trigger-type lookup: GUID -> trigger type(s)
                var triggerTypes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var triggerFolder in new[] { "Boot", "Logon", "Maintenance", "Plain" })
                {
                    var folderKey = GetCachedKey($@"Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\{triggerFolder}");
                    if (folderKey?.SubKeys != null)
                    {
                        foreach (var sub in folderKey.SubKeys)
                        {
                            var guid = sub.KeyName;
                            if (triggerTypes.ContainsKey(guid))
                                triggerTypes[guid] += ", " + triggerFolder;
                            else
                                triggerTypes[guid] = triggerFolder;
                        }
                    }
                }

                // Recursively walk the Tree, building hierarchical AnalysisItems
                CollectTreeTasks(treeKey, treePath, tasksPath, triggerTypes, section.Items);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error reading scheduled tasks: {ex.Message}");
                section.Items.Add(new AnalysisItem { Name = "Error", Value = ex.Message });
            }

            return section;
        }

        /// <summary>
        /// Recursively collect tasks from the TaskCache\Tree hierarchy.
        /// Folder keys (those with subkeys and no Id value) become parent items with IsSubSection=true.
        /// Leaf task keys (those with an Id value) become task items with SubItems for decoded details.
        /// </summary>
        private void CollectTreeTasks(RegistryKey parentKey, string treePath, string tasksPath,
            Dictionary<string, string> triggerTypes, List<AnalysisItem> items)
        {
            if (parentKey.SubKeys == null) return;

            foreach (var subKey in parentKey.SubKeys.OrderBy(k => k.KeyName))
            {
                var keyPath = $@"{treePath}\{subKey.KeyName}";
                var idValue = subKey.Values?.FirstOrDefault(v => v.ValueName == "Id")?.ValueData?.ToString();

                if (!string.IsNullOrEmpty(idValue))
                {
                    // Skip orphaned tree entries whose GUID has no matching Tasks\{GUID} key
                    var taskKeyPath = $@"{tasksPath}\{idValue}";
                    if (GetCachedKey(taskKeyPath) == null)
                        continue;

                    // This is a task leaf node - look up details from Tasks\{GUID}
                    var taskItem = BuildTaskItem(subKey.KeyName, idValue, tasksPath, triggerTypes, keyPath);
                    items.Add(taskItem);
                }
                else if (subKey.SubKeys != null && subKey.SubKeys.Count > 0)
                {
                    // This is a folder node - recurse
                    var subItems = new List<AnalysisItem>();
                    CollectTreeTasks(subKey, keyPath, tasksPath, triggerTypes, subItems);

                    // Skip empty folders (all children were orphaned)
                    if (subItems.Count == 0) continue;

                    var folderItem = new AnalysisItem
                    {
                        Name = subKey.KeyName,
                        Value = $"{subItems.Count} item(s)",
                        IsSubSection = true,
                        RegistryPath = keyPath,
                        SubItems = subItems
                    };

                    items.Add(folderItem);
                }
            }
        }

        /// <summary>
        /// Build an AnalysisItem for a single scheduled task by reading its Tasks\{GUID} entry.
        /// </summary>
        private AnalysisItem BuildTaskItem(string taskName, string guid, string tasksPath,
            Dictionary<string, string> triggerTypes, string treeKeyPath)
        {
            var taskKeyPath = $@"{tasksPath}\{guid}";
            var taskKey = GetCachedKey(taskKeyPath);

            var item = new AnalysisItem
            {
                Name = taskName,
                Value = guid,
                RegistryPath = treeKeyPath,
                RegistryValue = "Id",
                SubItems = new List<AnalysisItem>()
            };

            if (taskKey == null)
            {
                item.SubItems.Add(new AnalysisItem { Name = "Status", Value = "Task details not found" });
                return item;
            }

            // Determine task-level Enabled/Disabled status from Triggers blob
            // The JobBucket flags DWORD is at offset 0x28 in the Triggers binary.
            // Bit 0x400000 = enabled flag. If set, task is enabled; if not, disabled.
            var triggersRaw = taskKey.Values?.FirstOrDefault(v => v.ValueName == "Triggers")?.ValueDataRaw;
            bool isTaskEnabled = true; // default to enabled if no Triggers data
            if (triggersRaw != null && triggersRaw.Length >= 0x2C)
            {
                uint jobBucketFlags = BitConverter.ToUInt32(triggersRaw, 0x28);
                isTaskEnabled = (jobBucketFlags & 0x400000) != 0;
            }

            item.SubItems.Add(new AnalysisItem
            {
                Name = "Status",
                Value = isTaskEnabled ? "Ready" : "Disabled",
                RegistryPath = taskKeyPath,
                RegistryValue = "Triggers"
            });

            if (!isTaskEnabled)
                item.Name = taskName + " [Disabled]";

            // Read string values
            var path = taskKey.Values?.FirstOrDefault(v => v.ValueName == "Path")?.ValueData?.ToString();
            var author = taskKey.Values?.FirstOrDefault(v => v.ValueName == "Author")?.ValueData?.ToString();
            var description = taskKey.Values?.FirstOrDefault(v => v.ValueName == "Description")?.ValueData?.ToString();
            var source = taskKey.Values?.FirstOrDefault(v => v.ValueName == "Source")?.ValueData?.ToString();
            var uri = taskKey.Values?.FirstOrDefault(v => v.ValueName == "URI")?.ValueData?.ToString();
            var version = taskKey.Values?.FirstOrDefault(v => v.ValueName == "Version")?.ValueData?.ToString();
            var securityDescriptor = taskKey.Values?.FirstOrDefault(v => v.ValueName == "SecurityDescriptor")?.ValueData?.ToString();
            var dateStr = taskKey.Values?.FirstOrDefault(v => v.ValueName == "Date")?.ValueData?.ToString();

            if (!string.IsNullOrEmpty(path))
                item.SubItems.Add(new AnalysisItem { Name = "Path", Value = path, RegistryPath = taskKeyPath, RegistryValue = "Path" });
            if (!string.IsNullOrEmpty(uri) && uri != path)
                item.SubItems.Add(new AnalysisItem { Name = "URI", Value = uri, RegistryPath = taskKeyPath, RegistryValue = "URI" });
            if (!string.IsNullOrEmpty(author))
                item.SubItems.Add(new AnalysisItem { Name = "Author", Value = author, RegistryPath = taskKeyPath, RegistryValue = "Author" });
            if (!string.IsNullOrEmpty(description))
                item.SubItems.Add(new AnalysisItem { Name = "Description", Value = description, RegistryPath = taskKeyPath, RegistryValue = "Description" });
            if (!string.IsNullOrEmpty(source))
                item.SubItems.Add(new AnalysisItem { Name = "Source", Value = source, RegistryPath = taskKeyPath, RegistryValue = "Source" });
            if (!string.IsNullOrEmpty(version))
                item.SubItems.Add(new AnalysisItem { Name = "Version", Value = version, RegistryPath = taskKeyPath, RegistryValue = "Version" });

            // Trigger type from folder lookup
            if (triggerTypes.TryGetValue(guid, out var triggerType))
                item.SubItems.Add(new AnalysisItem { Name = "Trigger Type", Value = triggerType });

            // Decode Actions binary
            var actionsData = taskKey.Values?.FirstOrDefault(v => v.ValueName == "Actions")?.ValueDataRaw;
            if (actionsData != null && actionsData.Length > 0)
            {
                var actionsText = DecodeTaskActions(actionsData);
                item.SubItems.Add(new AnalysisItem { Name = "Action", Value = actionsText, RegistryPath = taskKeyPath, RegistryValue = "Actions" });
            }

            // Decode Triggers binary
            var triggersData = taskKey.Values?.FirstOrDefault(v => v.ValueName == "Triggers")?.ValueDataRaw;
            if (triggersData != null && triggersData.Length > 0)
            {
                var triggerDescriptions = DecodeTaskTriggers(triggersData);
                if (triggerDescriptions.Count > 0)
                {
                    for (int i = 0; i < triggerDescriptions.Count; i++)
                    {
                        var label = triggerDescriptions.Count == 1 ? "Trigger" : $"Trigger {i + 1}";
                        item.SubItems.Add(new AnalysisItem { Name = label, Value = triggerDescriptions[i], RegistryPath = taskKeyPath, RegistryValue = "Triggers" });
                    }
                }
            }

            // Decode DynamicInfo binary (timestamps)
            var dynamicData = taskKey.Values?.FirstOrDefault(v => v.ValueName == "DynamicInfo")?.ValueDataRaw;
            if (dynamicData != null && dynamicData.Length >= 28)
            {
                var (registered, lastRun, lastSuccessfulRun) = DecodeTaskDynamicInfo(dynamicData);

                // Show "Created" from the Date string value (original task creation date)
                if (!string.IsNullOrEmpty(dateStr) && DateTime.TryParse(dateStr, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out var createdDate))
                    item.SubItems.Add(new AnalysisItem { Name = "Created", Value = createdDate.ToString("yyyy-MM-dd HH:mm:ss"), RegistryPath = taskKeyPath, RegistryValue = "Date" });

                // Show "Last Registered" from DynamicInfo (when task was last re-registered with the service)
                if (registered.HasValue)
                    item.SubItems.Add(new AnalysisItem { Name = "Last Registered", Value = registered.Value.ToString("yyyy-MM-dd HH:mm:ss UTC"), RegistryPath = taskKeyPath, RegistryValue = "DynamicInfo" });

                if (lastSuccessfulRun.HasValue)
                    item.SubItems.Add(new AnalysisItem { Name = "Last Successful Run", Value = lastSuccessfulRun.Value.ToString("yyyy-MM-dd HH:mm:ss UTC"), RegistryPath = taskKeyPath, RegistryValue = "DynamicInfo" });
            }
            else if (!string.IsNullOrEmpty(dateStr) && DateTime.TryParse(dateStr, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out var createdDateOnly))
            {
                // No DynamicInfo but we still have a Date value
                item.SubItems.Add(new AnalysisItem { Name = "Created", Value = createdDateOnly.ToString("yyyy-MM-dd HH:mm:ss"), RegistryPath = taskKeyPath, RegistryValue = "Date" });
            }


            return item;
        }

        /// <summary>
        /// Decode the Actions binary value from a scheduled task.
        /// Format: version(2) + context_len(4) + context_utf16 + action_type(2+2+2) + payload
        /// Action type 0x6666 = command line, 0x7777 = COM handler
        /// </summary>
        private string DecodeTaskActions(byte[] data)
        {
            try
            {
                if (data.Length < 8) return BitConverter.ToString(data);

                int offset = 0;

                // Version (2 bytes, typically 0x0003)
                offset += 2;

                // Context string length in bytes
                int contextLen = BitConverter.ToInt32(data, offset);
                offset += 4;

                // Skip context string (UTF-16LE, contextLen bytes)
                if (contextLen > 0 && offset + contextLen <= data.Length)
                    offset += contextLen;

                // Decode one or more actions
                var actions = new List<string>();
                while (offset + 6 <= data.Length)
                {
                    ushort marker = BitConverter.ToUInt16(data, offset);
                    if (marker == 0) break; // end of actions

                    offset += 2; // marker
                    offset += 4; // 4 zero bytes after marker

                    string actionText;
                    if (marker == 0x6666)
                    {
                        actionText = DecodeCommandLineAction(data, offset);
                    }
                    else if (marker == 0x7777)
                    {
                        actionText = DecodeComHandlerAction(data, offset);
                    }
                    else if (marker == 0x8888)
                    {
                        actionText = "(Send Email - deprecated)";
                    }
                    else if (marker == 0x9999)
                    {
                        actionText = "(Show Message - deprecated)";
                    }
                    else
                    {
                        actionText = $"(action type 0x{marker:X4})";
                    }

                    actions.Add(actionText);

                    // Single-action tasks are ~99% of real tasks.
                    // Multi-action iteration would need ref-offset versions of decoders.
                    break;
                }

                return actions.Count > 0 ? string.Join(" | ", actions) : "(no actions)";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error decoding task actions: {ex.Message}");
                return "(decode error)";
            }
        }

        /// <summary>
        /// Decode a command-line action from the Actions binary.
        /// Format: command_len(4) + command_utf16 + args_len(4) + args_utf16 + workdir_len(4) + workdir_utf16
        /// </summary>
        private string DecodeCommandLineAction(byte[] data, int offset)
        {
            string command = ReadActionString(data, ref offset);
            string args = ReadActionString(data, ref offset);
            string workDir = ReadActionString(data, ref offset);

            var result = command;
            if (!string.IsNullOrEmpty(args))
                result += " " + args;
            if (!string.IsNullOrEmpty(workDir))
                result += $" (in {workDir})";

            return string.IsNullOrEmpty(result) ? "(empty command)" : result;
        }

        /// <summary>
        /// Decode a COM handler action from the Actions binary.
        /// Format: CLSID(16 bytes) + data_len(4) + data_utf16
        /// </summary>
        private string DecodeComHandlerAction(byte[] data, int offset)
        {
            return "Custom Handler";
        }

        /// <summary>
        /// Read a length-prefixed UTF-16LE string from the Actions binary.
        /// Format: length_in_bytes(4 bytes) + utf16_data (length bytes)
        /// </summary>
        private string ReadActionString(byte[] data, ref int offset)
        {
            if (offset + 4 > data.Length) return "";

            int byteCount = BitConverter.ToInt32(data, offset);
            offset += 4;

            if (byteCount <= 0 || offset + byteCount > data.Length) return "";

            var str = Encoding.Unicode.GetString(data, offset, byteCount).TrimEnd('\0');
            offset += byteCount;
            return str;
        }

        /// <summary>
        /// Decode the DynamicInfo binary value from a scheduled task.
        /// 36 bytes: version(4) + created_filetime(8) + last_run_filetime(8) + unknown(8) + last_successful_run_filetime(8)
        /// </summary>
        private (DateTime? Created, DateTime? LastRun, DateTime? LastSuccessfulRun) DecodeTaskDynamicInfo(byte[] data)
        {
            DateTime? created = null;
            DateTime? lastRun = null;
            DateTime? lastSuccessfulRun = null;

            try
            {
                if (data.Length >= 12)
                    created = FileTimeToDateTimeUtc(BitConverter.ToInt64(data, 4));
                if (data.Length >= 20)
                    lastRun = FileTimeToDateTimeUtc(BitConverter.ToInt64(data, 12));
                if (data.Length >= 36)
                    lastSuccessfulRun = FileTimeToDateTimeUtc(BitConverter.ToInt64(data, 28));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error decoding DynamicInfo: {ex.Message}");
            }

            return (created, lastRun, lastSuccessfulRun);
        }

        /// <summary>
        /// Convert a Windows FILETIME (100-nanosecond intervals since 1601-01-01) to UTC DateTime.
        /// Returns null for zero/invalid values.
        /// </summary>
        private DateTime? FileTimeToDateTimeUtc(long fileTime)
        {
            if (fileTime <= 0) return null;
            try
            {
                return DateTime.FromFileTimeUtc(fileTime);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Align a position to the next 8-byte boundary.
        /// </summary>
        private static int Align8(int v) => (v + 7) & ~7;

        /// <summary>
        /// Compute the size of a string block in the Triggers binary.
        /// If len == 0: 8 bytes (zero DWORD + 4 pad bytes).
        /// If len > 0: Align8(4 + len) bytes (4-byte length prefix + string data + pad).
        /// </summary>
        private static int TriggerStrBlock(uint len) => len == 0 ? 8 : Align8(4 + (int)len);

        /// <summary>
        /// Format a number of seconds into a human-readable duration string.
        /// </summary>
        private static string FormatSeconds(uint seconds)
        {
            if (seconds >= 86400 && seconds % 86400 == 0)
                return seconds / 86400 == 1 ? "1 day" : $"{seconds / 86400} days";
            if (seconds >= 3600 && seconds % 3600 == 0)
                return seconds / 3600 == 1 ? "1 hour" : $"{seconds / 3600} hours";
            if (seconds >= 60 && seconds % 60 == 0)
                return seconds / 60 == 1 ? "1 minute" : $"{seconds / 60} minutes";
            return $"{seconds}s";
        }

        /// <summary>
        /// Check if a trigger marker is a top-level marker (starts an independent trigger entry).
        /// </summary>
        private static bool IsTopLevelTriggerMarker(ushort marker) =>
            marker == 0xFFFF || marker == 0xDDDD || marker == 0xAAAA ||
            marker == 0x8888 || marker == 0x6666 || marker == 0xCCCC || marker == 0xBBBB;

        /// <summary>
        /// Check if a trigger marker is a sub-trigger marker (embedded inside a compound trigger).
        /// </summary>
        private static bool IsSubTriggerMarker(ushort marker) => marker == 0x7777 || marker == 0xEEEE;

        /// <summary>
        /// Convert a trigger type marker to a human-readable name.
        /// </summary>
        private static string TriggerMarkerToType(ushort marker) => marker switch
        {
            0xFFFF => "At startup",
            0xDDDD => "On a schedule",
            0xAAAA => "At log on",
            0x8888 => "On idle",
            0x6666 => "Custom Trigger",
            0xCCCC => "At task creation/modification",
            0xBBBB => "Session State Change",
            0x7777 => "Session State Change",
            0xEEEE => "Custom",
            _ => $"Unknown (0x{marker:X4})"
        };

        /// <summary>
        /// Convert a session state change code to a human-readable description.
        /// Used by BBBB/7777 (Session) triggers.
        /// </summary>
        private static string SessionStateToDescription(uint stateCode) => stateCode switch
        {
            1 => "On connection to local session",
            2 => "On disconnect from local session",
            3 => "On connection to remote session",
            4 => "On disconnect from remote session",
            5 => "On session logon",
            6 => "On session logoff",
            7 => "On session lock",
            8 => "On session unlock",
            _ => $"Session state change ({stateCode})"
        };

        /// <summary>
        /// Parse the Triggers binary header and return the byte offset where trigger entries begin.
        /// The header contains: version(4) + reserved(4) + 3 FILETIMEs(24) + duration(8) = 40 fixed bytes,
        /// followed by padded fields (job buckets, user string, flags, SID, display name, settings).
        /// All fields use 8-byte alignment with 0x48 as the padding byte.
        /// Returns -1 if the blob is too small or malformed.
        /// </summary>
        private static int SkipTriggerHeader(byte[] data)
        {
            if (data.Length < 0x30) return -1;

            try
            {
                int pos = 0x28; // After 40-byte fixed header

                // JobBucket1 (DWORD padded = 8 bytes) + JobBucket2 (DWORD padded = 8 bytes)
                pos += 16;
                if (pos + 8 > data.Length) return -1;

                // UserStrLen (DWORD padded = 8 bytes) + UserString data + align
                uint usrLen = BitConverter.ToUInt32(data, pos);
                pos += 8;
                if (usrLen > 0x1000 || pos + (int)usrLen > data.Length) return -1;
                pos += (int)usrLen;
                pos = Align8(pos);

                // PostStr (DWORD padded) + Flags1 (BYTE padded) + Flags2 (BYTE padded) = 24 bytes
                pos += 24;
                if (pos + 8 > data.Length) return -1;

                // SidSubAuthCount (DWORD padded = 8 bytes)
                pos += 8;
                if (pos + 8 > data.Length) return -1;

                // SidByteLen (DWORD padded = 8 bytes) + SID data + align
                uint sidLen = BitConverter.ToUInt32(data, pos);
                pos += 8;
                if (sidLen > 0x200 || pos + (int)sidLen > data.Length) return -1;
                pos += (int)sidLen;
                pos = Align8(pos);
                if (pos + 8 > data.Length) return -1;

                // DisplayNameLen (DWORD padded = 8 bytes) + optional DisplayName data + align
                uint dispLen = BitConverter.ToUInt32(data, pos);
                pos += 8;
                if (dispLen > 0 && dispLen < 0x200)
                {
                    if (pos + (int)dispLen > data.Length) return -1;
                    pos += (int)dispLen;
                    pos = Align8(pos);
                }
                if (pos + 8 > data.Length) return -1;

                // SettingsSize (DWORD padded = 8 bytes) + Settings data (NOT aligned)
                uint setLen = BitConverter.ToUInt32(data, pos);
                pos += 8;
                if (setLen > 0x1000 || pos + (int)setLen > data.Length) return -1;
                pos += (int)setLen;

                // Optional trailing 0x48484848 pad
                if (pos + 4 <= data.Length && data[pos] == 0x48 && data[pos + 1] == 0x48
                    && data[pos + 2] == 0x48 && data[pos + 3] == 0x48)
                    pos += 4;

                return pos;
            }
            catch
            {
                return -1;
            }
        }

        /// <summary>
        /// Walk embedded sub-triggers (0x7777 SESSION or 0xEEEE UNKNOWN) starting at the given position.
        /// Returns the total bytes consumed by all sub-triggers.
        /// </summary>
        private static int WalkSubTriggers(byte[] data, int startPos)
        {
            int pos = startPos;
            while (pos + 48 <= data.Length)
            {
                ushort sm = BitConverter.ToUInt16(data, pos);
                if (!IsSubTriggerMarker(sm)) break;
                if (pos + 8 > data.Length) break;
                if (BitConverter.ToUInt16(data, pos + 2) != 0 || BitConverter.ToUInt32(data, pos + 4) != 0) break;

                int sz = ComputeTriggerEntrySize(data, pos, sm);
                if (sz <= 0) break;
                pos += sz;
            }
            return pos - startPos;
        }

        /// <summary>
        /// Compute the total size (48-byte header + tail) of a trigger entry at the given offset.
        /// Returns -1 if the data is malformed.
        /// </summary>
        private static int ComputeTriggerEntrySize(byte[] data, int trigOffset, ushort marker)
        {
            int t = trigOffset + 48; // tail start
            if (t > data.Length) return -1;

            try
            {
                switch (marker)
                {
                    case 0xFFFF: // BOOT
                    case 0x8888: // IDLE
                    case 0xEEEE: // UNKNOWN sub-trigger
                    {
                        // Tail: 32-byte prefix + StrBlock(nameLen)
                        if (t + 0x24 > data.Length) return -1;
                        uint nameLen = BitConverter.ToUInt32(data, t + 0x20);
                        if (nameLen > 0x2000) return -1;
                        int tailSize = 32 + TriggerStrBlock(nameLen);
                        return 48 + tailSize;
                    }

                    case 0xDDDD: // TIME
                    {
                        // Tail: 48-byte prefix + StrBlock(nameLen) + optional sub-triggers
                        if (t + 0x34 > data.Length) return -1;
                        uint nameLen = BitConverter.ToUInt32(data, t + 0x30);
                        if (nameLen > 0x2000) return -1;
                        int baseTail = 48 + TriggerStrBlock(nameLen);
                        int subs = WalkSubTriggers(data, t + baseTail);
                        return 48 + baseTail + subs;
                    }

                    case 0x6666: // EVENT
                    {
                        // Tail: 32-byte prefix + StrBlock(nameLen) + WNF(8) + subCount(4) + pad(4) + subData + optional sub-triggers
                        if (t + 0x24 > data.Length) return -1;
                        uint nameLen = BitConverter.ToUInt32(data, t + 0x20);
                        if (nameLen > 0x2000) return -1;
                        int strBlk = TriggerStrBlock(nameLen);
                        int afterStr = t + 32 + strBlk;
                        if (afterStr + 12 > data.Length) return -1;
                        uint subCount = BitConverter.ToUInt32(data, afterStr + 8);
                        if (subCount > 0x10000) return -1;
                        int baseTail = 32 + strBlk + 8 + 8 + (int)subCount;
                        int subs = WalkSubTriggers(data, t + baseTail);
                        return 48 + baseTail + subs;
                    }

                    case 0xAAAA: // LOGON
                    {
                        // Tail: 32-byte prefix + StrBlock(nameLen) + filterFlag(8) [+ user identity block]
                        if (t + 0x24 > data.Length) return -1;
                        uint nameLen = BitConverter.ToUInt32(data, t + 0x20);
                        if (nameLen > 0x2000) return -1;
                        int strBlk = TriggerStrBlock(nameLen);
                        int flagPos = t + 32 + strBlk;
                        if (flagPos >= data.Length) return -1;
                        byte filterFlag = data[flagPos];
                        int baseTail = 32 + strBlk + 8;

                        if (filterFlag == 0x00)
                        {
                            // Specific user: UnknownByte(8) + SidSubAuthCount(8) + SidByteLen(8) + SidData + DisplayNameLen(8) + DisplayName
                            int pos = t + baseTail;
                            pos += 8; // UnknownByte
                            pos += 8; // SidSubAuthCount
                            if (pos + 4 > data.Length) return -1;
                            uint sidLen = BitConverter.ToUInt32(data, pos);
                            if (sidLen > 0x200) return -1;
                            pos += 8; // SidByteLen field
                            pos += (int)sidLen;
                            pos = Align8(pos);
                            if (pos + 4 > data.Length) return -1;
                            uint dispLen = BitConverter.ToUInt32(data, pos);
                            if (dispLen > 0x2000) return -1;
                            pos += 8; // DisplayNameLen field
                            pos += (int)dispLen;
                            pos = Align8(pos);
                            return pos - trigOffset;
                        }
                        else
                        {
                            int subs = WalkSubTriggers(data, t + baseTail);
                            return 48 + baseTail + subs;
                        }
                    }

                    case 0xCCCC: // REGISTRATION
                    {
                        // Tail: 32-byte prefix + StrBlock(nameLen) + xmlLenChars(8) + xmlData + align + trailing(24)
                        if (t + 0x24 > data.Length) return -1;
                        uint nameLen = BitConverter.ToUInt32(data, t + 0x20);
                        if (nameLen > 0x2000) return -1;
                        int strBlk = TriggerStrBlock(nameLen);
                        int xmlLenPos = t + 32 + strBlk;
                        if (xmlLenPos + 8 > data.Length) return -1;
                        uint xmlChars = BitConverter.ToUInt32(data, xmlLenPos);
                        if (xmlChars > 0x10000) return -1;
                        int xmlBytes = (int)xmlChars * 2 + 2; // UTF-16 + null terminator
                        int xmlEnd = xmlLenPos + 8 + xmlBytes;
                        int xmlAligned = Align8(xmlEnd);
                        return (xmlAligned - trigOffset) + 24;
                    }

                    case 0xBBBB: // SESSION (top-level)
                    case 0x7777: // SESSION (sub-trigger)
                    {
                        // Tail: 32-byte prefix + StrBlock(nameLen) + sessionStateChange(8) + trailingByte(8) = +16
                        if (t + 0x24 > data.Length) return -1;
                        uint nameLen = BitConverter.ToUInt32(data, t + 0x20);
                        if (nameLen > 0x2000) return -1;
                        int tailSize = 32 + TriggerStrBlock(nameLen) + 16;
                        return 48 + tailSize;
                    }

                    default:
                        return -1;
                }
            }
            catch
            {
                return -1;
            }
        }

        /// <summary>
        /// Decode the Triggers binary value from a scheduled task.
        /// Returns a list of human-readable trigger descriptions.
        /// Format: header (variable, 0x48-padded fields) + trigger entries (48-byte header + variable tail each).
        /// Trigger types: 0xFFFF=Boot, 0xDDDD=Time, 0xAAAA=Logon, 0x8888=Idle,
        ///   0x6666=Event, 0xCCCC=Registration, 0xBBBB/0x7777=Session, 0xEEEE=Custom.
        /// </summary>
        private List<string> DecodeTaskTriggers(byte[] data)
        {
            var results = new List<string>();

            try
            {
                int pos = SkipTriggerHeader(data);
                if (pos < 0 || pos >= data.Length) return results;

                while (pos + 48 <= data.Length)
                {
                    ushort marker = BitConverter.ToUInt16(data, pos);
                    if (!IsTopLevelTriggerMarker(marker)) break;

                    // Validate marker signature: marker(2) + 00 00 + 00 00 00 00
                    if (pos + 8 > data.Length) break;
                    if (BitConverter.ToUInt16(data, pos + 2) != 0 || BitConverter.ToUInt32(data, pos + 4) != 0) break;

                    string typeName = TriggerMarkerToType(marker);

                    // Extract trigger name from tail
                    int t = pos + 48;
                    string? triggerName = null;
                    string? xmlQuery = null;

                    int nameOffset = (marker == 0xDDDD) ? 0x30 : 0x20;
                    if (t + nameOffset + 4 <= data.Length)
                    {
                        uint nameLen = BitConverter.ToUInt32(data, t + nameOffset);
                        if (nameLen > 0 && nameLen < 0x1000 && t + nameOffset + 4 + (int)nameLen <= data.Length)
                        {
                            triggerName = Encoding.Unicode.GetString(data, t + nameOffset + 4, (int)nameLen).TrimEnd('\0');
                        }
                    }

                    // Extract repetition interval and per-trigger enabled/disabled flag.
                    // GenericData triggers (Boot, Idle, Event, Logon, Registration, Session) store
                    // repetition_interval at entry offset +0x30 and enabled at +0x40.
                    // JobSchedule triggers (DDDD Time/Calendar) store repetition_interval at +0x38
                    // and enabled at +0x44 (different layout).
                    string? repetitionInfo = null;
                    bool triggerEnabled = true;
                    if (marker == 0xDDDD)
                    {
                        // DDDD uses JobSchedule: interval at tail+0x08 (pos+0x38), enabled at tail+0x14 (pos+0x44)
                        int repPos = pos + 0x38;
                        if (repPos + 8 <= data.Length)
                        {
                            uint repInterval = BitConverter.ToUInt32(data, repPos);
                            uint repDuration = BitConverter.ToUInt32(data, repPos + 4);
                            if (repInterval > 0 && repInterval != 0xFFFFFFFF)
                            {
                                var parts = new List<string>();
                                parts.Add($"repeat every {FormatSeconds(repInterval)}");
                                if (repDuration > 0 && repDuration != 0xFFFFFFFF)
                                    parts.Add($"for {FormatSeconds(repDuration)}");
                                repetitionInfo = string.Join(" ", parts);
                            }
                        }

                        int enabledPos = pos + 0x44;
                        if (enabledPos + 4 <= data.Length)
                        {
                            uint enabledFlag = BitConverter.ToUInt32(data, enabledPos);
                            triggerEnabled = enabledFlag != 0;
                        }
                    }
                    else
                    {
                        // GenericData layout: interval at entry+0x30, duration at +0x34, enabled at +0x40
                        int repPos = pos + 0x30;
                        if (repPos + 8 <= data.Length)
                        {
                            uint repInterval = BitConverter.ToUInt32(data, repPos);
                            uint repDuration = BitConverter.ToUInt32(data, repPos + 4);
                            if (repInterval > 0 && repInterval != 0xFFFFFFFF)
                            {
                                var parts = new List<string>();
                                parts.Add($"repeat every {FormatSeconds(repInterval)}");
                                if (repDuration > 0 && repDuration != 0xFFFFFFFF)
                                    parts.Add($"for {FormatSeconds(repDuration)}");
                                repetitionInfo = string.Join(" ", parts);
                            }
                        }

                        int enabledPos = pos + 0x40;
                        if (enabledPos + 1 <= data.Length)
                        {
                            triggerEnabled = data[enabledPos] != 0;
                        }
                    }

                    // Extract XML query for Registration triggers
                    if (marker == 0xCCCC && t + 0x24 <= data.Length)
                    {
                        uint nameLen = BitConverter.ToUInt32(data, t + 0x20);
                        int strBlk = TriggerStrBlock(nameLen);
                        int xmlPos = t + 32 + strBlk;
                        if (xmlPos + 8 <= data.Length)
                        {
                            uint xmlChars = BitConverter.ToUInt32(data, xmlPos);
                            if (xmlChars > 0 && xmlChars < 0x10000 && xmlPos + 8 + (int)xmlChars * 2 <= data.Length)
                            {
                                xmlQuery = Encoding.Unicode.GetString(data, xmlPos + 8, (int)xmlChars * 2).TrimEnd('\0');
                            }
                        }
                    }

                    // Extract user info for Logon triggers
                    string? logonUserInfo = null;
                    if (marker == 0xAAAA && t + 0x24 <= data.Length)
                    {
                        uint nameLen = BitConverter.ToUInt32(data, t + 0x20);
                        if (nameLen <= 0x2000)
                        {
                            int strBlk = TriggerStrBlock(nameLen);
                            int filterFlagPos = t + 32 + strBlk;
                            if (filterFlagPos < data.Length)
                            {
                                byte filterFlag = data[filterFlagPos];
                                if (filterFlag != 0x00)
                                {
                                    // Any user
                                    logonUserInfo = "of any user";
                                }
                                else
                                {
                                    // Specific user — try to extract display name
                                    // Layout after filterFlag: filterFlag(8) + UnknownByte(8) + SidSubAuthCount(8)
                                    //   + SidByteLen(8) + SidData + align + DisplayNameLen(8) + DisplayName
                                    int userPos = filterFlagPos + 8; // skip filterFlag padded field
                                    userPos += 8; // skip UnknownByte
                                    userPos += 8; // skip SidSubAuthCount
                                    if (userPos + 4 <= data.Length)
                                    {
                                        uint sidLen = BitConverter.ToUInt32(data, userPos);
                                        if (sidLen > 0 && sidLen <= 0x200)
                                        {
                                            userPos += 8; // skip SidByteLen field (padded)
                                            userPos += (int)sidLen;
                                            userPos = Align8(userPos);
                                            if (userPos + 4 <= data.Length)
                                            {
                                                uint dispLen = BitConverter.ToUInt32(data, userPos);
                                                if (dispLen > 0 && dispLen < 0x2000)
                                                {
                                                    userPos += 8; // skip DisplayNameLen field (padded)
                                                    if (userPos + (int)dispLen <= data.Length)
                                                    {
                                                        string displayName = Encoding.Unicode.GetString(data, userPos, (int)dispLen).TrimEnd('\0');
                                                        if (!string.IsNullOrEmpty(displayName))
                                                            logonUserInfo = $"of {displayName}";
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // Build human-readable description
                    var desc = new StringBuilder(typeName);

                    if (!string.IsNullOrEmpty(logonUserInfo))
                        desc.Append($" {logonUserInfo}");

                    if (!string.IsNullOrEmpty(triggerName))
                        desc.Append($": {triggerName}");

                    if (!string.IsNullOrEmpty(repetitionInfo))
                        desc.Append($" ({repetitionInfo})");

                    if (!string.IsNullOrEmpty(xmlQuery))
                    {
                        // Try to extract EventID from XML query for a cleaner display
                        var eventIdMatch = System.Text.RegularExpressions.Regex.Match(xmlQuery, @"EventID=(\d+)");
                        if (eventIdMatch.Success)
                            desc.Append($" [EventID={eventIdMatch.Groups[1].Value}]");
                        else
                            desc.Append($" [{xmlQuery}]");
                    }

                    if (!triggerEnabled)
                        desc.Append(" [Disabled]");

                    results.Add(desc.ToString());

                    // Decode sub-triggers (7777 Session / EEEE Custom) embedded in compound triggers
                    int entrySize = ComputeTriggerEntrySize(data, pos, marker);
                    if (entrySize > 0)
                    {
                        int subStart = -1;
                        if (marker == 0xDDDD && t + 0x34 <= data.Length)
                        {
                            uint nLen = BitConverter.ToUInt32(data, t + 0x30);
                            if (nLen <= 0x2000)
                                subStart = t + 48 + TriggerStrBlock(nLen);
                        }
                        else if (marker == 0x6666 && t + 0x24 <= data.Length)
                        {
                            uint nLen = BitConverter.ToUInt32(data, t + 0x20);
                            if (nLen <= 0x2000)
                            {
                                int strBlk = TriggerStrBlock(nLen);
                                int afterStr = t + 32 + strBlk;
                                if (afterStr + 12 <= data.Length)
                                {
                                    uint subCount = BitConverter.ToUInt32(data, afterStr + 8);
                                    if (subCount <= 0x10000)
                                        subStart = t + 32 + strBlk + 8 + 8 + (int)subCount;
                                }
                            }
                        }

                        if (subStart > 0)
                        {
                            int subEnd = pos + entrySize;
                            DecodeSubTriggers(data, subStart, subEnd, results);
                        }
                    }

                    // Advance past this trigger entry (including its sub-triggers)
                    if (entrySize <= 0) break;
                    pos += entrySize;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error decoding task triggers: {ex.Message}");
            }

            return results;
        }

        /// <summary>
        /// Decode embedded sub-triggers (0x7777 Session or 0xEEEE Custom) and add their descriptions to results.
        /// Sub-triggers are found inside compound triggers (DDDD Time/Calendar and 6666 Event).
        /// </summary>
        private void DecodeSubTriggers(byte[] data, int startPos, int endPos, List<string> results)
        {
            int pos = startPos;
            while (pos + 48 <= endPos && pos + 48 <= data.Length)
            {
                ushort sm = BitConverter.ToUInt16(data, pos);
                if (!IsSubTriggerMarker(sm)) break;
                if (pos + 8 > data.Length) break;
                if (BitConverter.ToUInt16(data, pos + 2) != 0 || BitConverter.ToUInt32(data, pos + 4) != 0) break;

                string typeName = TriggerMarkerToType(sm);
                int st = pos + 48; // sub-trigger tail start

                // For 7777 (Session) sub-triggers: extract session state change type
                // Tail: 32-byte prefix + StrBlock(nameLen) + sessionStateChange(8) + disabledFlag(8)
                if (sm == 0x7777 && st + 0x24 <= data.Length)
                {
                    uint nameLen = BitConverter.ToUInt32(data, st + 0x20);
                    if (nameLen <= 0x2000)
                    {
                        int strBlk = TriggerStrBlock(nameLen);
                        int statePos = st + 32 + strBlk;
                        if (statePos + 8 <= data.Length)
                        {
                            uint stateCode = BitConverter.ToUInt32(data, statePos);
                            var desc = new StringBuilder(SessionStateToDescription(stateCode));

                            // Disabled flag: byte at statePos + 8 (padded to 8 bytes)
                            // Observed: 0x01 = disabled for all known disabled 7777 triggers
                            int flagPos = statePos + 8;
                            if (flagPos < data.Length && data[flagPos] == 0x01)
                                desc.Append(" [Disabled]");

                            results.Add(desc.ToString());
                        }
                        else
                        {
                            results.Add(typeName);
                        }
                    }
                    else
                    {
                        results.Add(typeName);
                    }
                }
                else
                {
                    // EEEE (Custom) or unknown sub-trigger — just show the type name
                    results.Add(typeName);
                }

                int sz = ComputeTriggerEntrySize(data, pos, sm);
                if (sz <= 0) break;
                pos += sz;
            }
        }

        #endregion

        #region Health Analysis

        /// <summary>
        /// Get hive health analysis as structured sections.
        /// All checks are structurally verifiable — only spec-defined integrity checks are performed.
        /// </summary>
        public List<AnalysisSection> GetHealthAnalysis()
        {
            var sections = new List<AnalysisSection>();
            int errorCount = 0;
            int warningCount = 0;

            // --- Hive Integrity ---
            var integritySection = new AnalysisSection { Title = "Hive Integrity" };

            // Signature check
            bool signatureValid = _parser.HasValidSignature;
            integritySection.Items.Add(new AnalysisItem
            {
                Name = "Signature",
                Value = signatureValid ? "OK (regf)" : "FAILED - Invalid signature",
                IsWarning = !signatureValid
            });
            if (!signatureValid) errorCount++;

            // Header checksum
            bool checksumValid = _parser.IsChecksumValid;
            int storedChecksum = _parser.CheckSum;
            int calculatedChecksum = _parser.CalculatedChecksum;
            integritySection.Items.Add(new AnalysisItem
            {
                Name = "Header Checksum",
                Value = checksumValid
                    ? $"OK (0x{storedChecksum:X8})"
                    : $"FAILED - Stored: 0x{storedChecksum:X8}, Calculated: 0x{calculatedChecksum:X8}",
                IsWarning = !checksumValid
            });
            if (!checksumValid) errorCount++;

            // Sequence numbers
            uint primarySeq = _parser.PrimarySequenceNumber;
            uint secondarySeq = _parser.SecondarySequenceNumber;
            bool seqMatch = primarySeq == secondarySeq;
            integritySection.Items.Add(new AnalysisItem
            {
                Name = "Sequence Numbers",
                Value = seqMatch
                    ? $"OK (Primary: {primarySeq}, Secondary: {secondarySeq})"
                    : $"Mismatch - Primary: {primarySeq}, Secondary: {secondarySeq} (normal for offline hives)",
                IsWarning = false  // Mismatch is expected for offline hives — transaction logs handle replay on boot
            });

            // File size vs header hive bins data size
            long fileLength = _parser.FileLength;
            uint headerLength = _parser.HeaderLength;
            long expectedMinSize = headerLength + 4096; // hive bins data + base block
            bool sizeConsistent = fileLength >= expectedMinSize;
            if (headerLength > 0)
            {
                integritySection.Items.Add(new AnalysisItem
                {
                    Name = "File Size Consistency",
                    Value = sizeConsistent
                        ? $"OK (File: {fileLength:N0} bytes, Header declares: {headerLength:N0} bytes of hive data)"
                        : $"WARNING - File ({fileLength:N0} bytes) is smaller than header expects ({expectedMinSize:N0} bytes)",
                    IsWarning = !sizeConsistent
                });
                if (!sizeConsistent) warningCount++;
            }

            sections.Add(integritySection);

            // --- Header Details ---
            var headerSection = new AnalysisSection { Title = "Header Details" };

            headerSection.Items.Add(new AnalysisItem
            {
                Name = "Version",
                Value = $"{_parser.MajorVersion}.{_parser.MinorVersion}"
            });

            var lastWrite = _parser.LastWriteTimestamp;
            headerSection.Items.Add(new AnalysisItem
            {
                Name = "Last Write Timestamp",
                Value = lastWrite.HasValue ? lastWrite.Value.UtcDateTime.ToString("yyyy-MM-dd HH:mm:ss UTC") : "Not available"
            });

            headerSection.Items.Add(new AnalysisItem
            {
                Name = "Hive Bins Data Size",
                Value = $"{headerLength:N0} bytes"
            });

            headerSection.Items.Add(new AnalysisItem
            {
                Name = "File Size",
                Value = $"{fileLength:N0} bytes"
            });

            headerSection.Items.Add(new AnalysisItem
            {
                Name = "Root Cell Offset",
                Value = $"0x{_parser.RootCellOffset:X8}"
            });

            string headerFileName = _parser.HeaderFileName;
            if (!string.IsNullOrEmpty(headerFileName))
            {
                headerSection.Items.Add(new AnalysisItem
                {
                    Name = "Embedded File Name",
                    Value = headerFileName
                });
            }

            sections.Add(headerSection);

            // --- Parsing Results ---
            var parsingSection = new AnalysisSection { Title = "Parsing Results" };

            int hardErrors = _parser.HardParsingErrors;
            int softErrors = _parser.SoftParsingErrors;

            parsingSection.Items.Add(new AnalysisItem
            {
                Name = "Hard Parsing Errors",
                Value = hardErrors == 0 ? "0" : $"{hardErrors} (cells that could not be read)",
                IsWarning = hardErrors > 0
            });
            if (hardErrors > 0) warningCount++;

            parsingSection.Items.Add(new AnalysisItem
            {
                Name = "Soft Parsing Errors",
                Value = softErrors == 0 ? "0" : $"{softErrors} (non-critical read issues — normal for used hives)",
                IsWarning = false  // Soft errors are informational; free cells with unrecognized signatures are expected
            });

            parsingSection.Items.Add(new AnalysisItem
            {
                Name = "HBin Records",
                Value = _parser.HBinRecordCount.ToString("N0")
            });

            parsingSection.Items.Add(new AnalysisItem
            {
                Name = "HBin Total Size",
                Value = $"{_parser.HBinRecordTotalSize:N0} bytes"
            });

            sections.Add(parsingSection);

            // --- Security Descriptors ---
            var securitySection = new AnalysisSection { Title = "Security Descriptors" };

            var skResult = ValidateSecurityChain();
            securitySection.Items.Add(new AnalysisItem
            {
                Name = "SK Records Found",
                Value = skResult.TotalRecords.ToString("N0")
            });

            securitySection.Items.Add(new AnalysisItem
            {
                Name = "Chain Integrity",
                Value = skResult.IsChainValid
                    ? "OK (circular doubly-linked list intact)"
                    : $"FAILED - {skResult.ChainError}",
                IsWarning = !skResult.IsChainValid
            });
            if (!skResult.IsChainValid) errorCount++;

            if (skResult.OrphanedRecords > 0)
            {
                securitySection.Items.Add(new AnalysisItem
                {
                    Name = "Orphaned SK Records",
                    Value = $"{skResult.OrphanedRecords} (not in FLink/BLink chain — normal for used hives)",
                    IsWarning = false  // Orphaned SK records are common when keys are deleted; not a data integrity concern
                });
            }

            sections.Add(securitySection);

            // --- Key Tree Statistics ---
            var treeSection = new AnalysisSection { Title = "Key Tree Statistics" };

            var treeResult = WalkKeyTree();
            treeSection.Items.Add(new AnalysisItem
            {
                Name = "Total Keys",
                Value = treeResult.TotalKeys.ToString("N0")
            });

            treeSection.Items.Add(new AnalysisItem
            {
                Name = "Total Values",
                Value = treeResult.TotalValues.ToString("N0")
            });

            treeSection.Items.Add(new AnalysisItem
            {
                Name = "Maximum Depth",
                Value = treeResult.MaxDepth.ToString()
            });

            if (treeResult.InvalidParentPointers > 0)
            {
                treeSection.Items.Add(new AnalysisItem
                {
                    Name = "Invalid Parent Pointers",
                    Value = $"{treeResult.InvalidParentPointers} keys have parent pointers that don't match their actual parent",
                    IsWarning = true
                });
                errorCount++;
            }
            else
            {
                treeSection.Items.Add(new AnalysisItem
                {
                    Name = "Parent Pointer Validation",
                    Value = "OK (all parent pointers consistent)"
                });
            }

            sections.Add(treeSection);

            // --- Overall Status (inserted at the top) ---
            var statusSection = new AnalysisSection { Title = "Overall Status" };
            string statusValue;
            bool statusWarning;

            if (errorCount > 0)
            {
                statusValue = $"ERRORS DETECTED ({errorCount} error{(errorCount != 1 ? "s" : "")}, {warningCount} warning{(warningCount != 1 ? "s" : "")})";
                statusWarning = true;
            }
            else if (warningCount > 0)
            {
                statusValue = $"WARNINGS ({warningCount} warning{(warningCount != 1 ? "s" : "")})";
                statusWarning = true;
            }
            else
            {
                statusValue = "Healthy (0 errors, 0 warnings)";
                statusWarning = false;
            }

            statusSection.Items.Add(new AnalysisItem
            {
                Name = "Status",
                Value = statusValue,
                IsWarning = statusWarning
            });

            sections.Insert(0, statusSection);

            return sections;
        }

        /// <summary>
        /// Validates the security descriptor (SK) doubly-linked list chain.
        /// SK records form a circular linked list via FLink/BLink.
        /// </summary>
        private SecurityChainResult ValidateSecurityChain()
        {
            var result = new SecurityChainResult();

            try
            {
                var skRecords = _parser.GetSecurityRecords().ToList();
                result.TotalRecords = skRecords.Count;

                if (skRecords.Count == 0)
                {
                    // No SK records is unusual but not necessarily an error for empty hives
                    result.IsChainValid = true;
                    return result;
                }

                // Build a lookup by relative offset for chain traversal
                var skByOffset = new Dictionary<long, SkCellRecord>();
                foreach (var sk in skRecords)
                {
                    skByOffset[sk.RelativeOffset] = sk;
                }

                // Start from the first SK record and walk FLink
                var startSk = skRecords[0];
                var visited = new HashSet<long>();
                var current = startSk;
                bool chainBroken = false;

                while (true)
                {
                    if (!visited.Add(current.RelativeOffset))
                    {
                        // We've looped back — check if it's to the start (valid circular list)
                        if (current.RelativeOffset == startSk.RelativeOffset)
                        {
                            // Valid circular chain
                            break;
                        }
                        else
                        {
                            chainBroken = true;
                            result.ChainError = $"FLink chain loops to offset 0x{current.RelativeOffset:X} instead of returning to start";
                            break;
                        }
                    }

                    // Follow FLink
                    uint flink = current.FLink;
                    if (!skByOffset.TryGetValue(flink, out var next))
                    {
                        chainBroken = true;
                        result.ChainError = $"FLink at offset 0x{current.RelativeOffset:X} points to 0x{flink:X} which is not a valid SK record";
                        break;
                    }

                    // Verify BLink consistency: next.BLink should point back to current
                    if (next.BLink != (uint)current.RelativeOffset)
                    {
                        chainBroken = true;
                        result.ChainError = $"BLink inconsistency: SK at 0x{next.RelativeOffset:X} has BLink 0x{next.BLink:X}, expected 0x{current.RelativeOffset:X}";
                        break;
                    }

                    current = next;
                }

                result.IsChainValid = !chainBroken;
                result.OrphanedRecords = skRecords.Count - visited.Count;
            }
            catch (Exception ex)
            {
                result.IsChainValid = false;
                result.ChainError = $"Validation failed: {ex.Message}";
            }

            return result;
        }

        /// <summary>
        /// Walks the key tree from root, counting keys/values/depth
        /// and validating parent pointers.
        /// </summary>
        private KeyTreeResult WalkKeyTree()
        {
            var result = new KeyTreeResult();

            try
            {
                var rootKey = _parser.GetRootKey();
                if (rootKey == null)
                {
                    return result;
                }

                // Iterative walk to avoid stack overflow on deep hives
                var stack = new Stack<(RegistryKey key, RegistryKey? parent, int depth)>();
                stack.Push((rootKey, null, 0));

                while (stack.Count > 0)
                {
                    var (key, parent, depth) = stack.Pop();
                    result.TotalKeys++;

                    if (depth > result.MaxDepth)
                        result.MaxDepth = depth;

                    // Count values
                    try
                    {
                        result.TotalValues += key.Values?.Count ?? 0;
                    }
                    catch { /* value list may be corrupt */ }

                    // Validate parent pointer (skip root which has no meaningful parent)
                    if (parent != null)
                    {
                        try
                        {
                            uint parentCellIndex = key.NkRecord.ParentCellIndex;
                            long parentActualOffset = parent.NkRecord.RelativeOffset;
                            if (parentCellIndex != (uint)parentActualOffset)
                            {
                                result.InvalidParentPointers++;
                            }
                        }
                        catch { /* NkRecord access may fail on corrupt keys */ }
                    }

                    // Push children
                    try
                    {
                        if (key.SubKeys != null)
                        {
                            foreach (var child in key.SubKeys)
                            {
                                stack.Push((child, key, depth + 1));
                            }
                        }
                    }
                    catch { /* subkey list may be corrupt */ }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Health check tree walk error: {ex.Message}");
            }

            return result;
        }

        private class SecurityChainResult
        {
            public int TotalRecords { get; set; }
            public bool IsChainValid { get; set; }
            public string ChainError { get; set; } = "";
            public int OrphanedRecords { get; set; }
        }

        private class KeyTreeResult
        {
            public int TotalKeys { get; set; }
            public int TotalValues { get; set; }
            public int MaxDepth { get; set; }
            public int InvalidParentPointers { get; set; }
        }

        #endregion
    }

    /// <summary>
    /// Firewall rule information
    /// </summary>
    public class FirewallRuleInfo
    {
        public string RuleId { get; set; } = "";
        public string Name { get; set; } = "";
        public string Description { get; set; } = "";
        public string Action { get; set; } = "";
        public bool IsActive { get; set; }
        public string Direction { get; set; } = "";
        public string Protocol { get; set; } = "";
        public string Profiles { get; set; } = "";
        public string LocalPorts { get; set; } = "";
        public string RemotePorts { get; set; } = "";
        public string LocalAddresses { get; set; } = "";
        public string RemoteAddresses { get; set; } = "";
        public string Application { get; set; } = "";
        public string Service { get; set; } = "";
        public string EmbedContext { get; set; } = "";
        public string PackageFamilyName { get; set; } = "";
        public string RegistryPath { get; set; } = "";
        public string RegistryValueName { get; set; } = "";
        public string RawData { get; set; } = "";

    }

    /// <summary>
    /// Service information for display in the Services view
    /// </summary>
    public class ServiceInfo
    {
        public string ServiceName { get; set; } = "";
        public string DisplayName { get; set; } = "";
        public string Description { get; set; } = "";
        public string StartType { get; set; } = "";
        public string StartTypeName { get; set; } = "";
        public string ImagePath { get; set; } = "";
        public string RegistryPath { get; set; } = "";
        public bool IsDisabled { get; set; }
        public bool IsAutoStart { get; set; }
        public bool IsDelayedAutoStart { get; set; }
        public bool IsBoot { get; set; }
        public bool IsSystem { get; set; }
        public bool IsManual { get; set; }
    }

    /// <summary>
    /// Represents a certificate store (e.g., Personal, Root, CA).
    /// </summary>
    public class CertificateStoreInfo
    {
        public string RegistryName { get; set; } = "";
        public string FriendlyName { get; set; } = "";
        public string RegistryPath { get; set; } = "";
        public List<CertificateInfo> Certificates { get; set; } = new();
    }

    /// <summary>
    /// Represents a single certificate extracted from a registry certificate store.
    /// </summary>
    public class CertificateInfo
    {
        public string DisplayName { get; set; } = "";
        public string Thumbprint { get; set; } = "";
        public string Sha1Hash { get; set; } = "";
        public string Subject { get; set; } = "";
        public string Issuer { get; set; } = "";
        public string SerialNumber { get; set; } = "";
        public DateTime? ValidFrom { get; set; }
        public DateTime? ValidTo { get; set; }
        public string SignatureAlgorithm { get; set; } = "";
        public string FriendlyName { get; set; } = "";
        public string KeyProvider { get; set; } = "";
        public string KeyContainer { get; set; } = "";
        public string RegistryPath { get; set; } = "";
    }
}

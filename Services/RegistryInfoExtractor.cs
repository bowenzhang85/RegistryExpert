using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections.Concurrent;
using RegistryParser.Abstractions;

namespace RegistryExpert
{
    /// <summary>
    /// Extracts useful information from registry hives
    /// </summary>
    public class RegistryInfoExtractor
    {
        private readonly OfflineRegistryParser _parser;
        
        // Cache for frequently accessed registry keys
        private readonly ConcurrentDictionary<string, RegistryKey?> _keyCache = new();
        private readonly ConcurrentDictionary<string, string?> _valueCache = new();

        public RegistryInfoExtractor(OfflineRegistryParser parser)
        {
            _parser = parser;
        }

        /// <summary>
        /// Get a cached registry key
        /// </summary>
        private RegistryKey? GetCachedKey(string path)
        {
            return _keyCache.GetOrAdd(path, p => _parser.GetKey(p));
        }


        #region Helper Methods

        private string? GetValue(string keyPath, string valueName)
        {
            // Use cache key that includes value name
            var cacheKey = $"{keyPath}\\{valueName}";
            return _valueCache.GetOrAdd(cacheKey, _ =>
            {
                try
                {
                    var key = GetCachedKey(keyPath);
                    return key?.Values.FirstOrDefault(v => v.ValueName == valueName)?.ValueData?.ToString();
                }
                catch
                {
                    return null;
                }
            });
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
            var computerNamePath = @"ControlSet001\Control\ComputerName\ComputerName";
            var computerName = GetValue(computerNamePath, "ComputerName");
            if (string.IsNullOrEmpty(computerName))
            {
                computerNamePath = @"ControlSet001\Control\ComputerName\ActiveComputerName";
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

            var timezonePath = @"ControlSet001\Control\TimeZoneInformation";
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

            var shutdownPath = @"ControlSet001\Control\Windows";
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

            // Always add for UI consistency (will appear greyed out if empty)
            sections.Add(computerSection);

            // OS Information Section (SOFTWARE hive) - Always include for UI consistency
            const string osRegPath = @"Microsoft\Windows NT\CurrentVersion";
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
            const string memMgmtPath = @"ControlSet001\Control\Session Manager\Memory Management";
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
            var crashControlKey = _parser.GetKey(@"ControlSet001\Control\CrashControl");
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
                    RegistryPath = @"ControlSet001\Control\CrashControl",
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
                        RegistryPath = @"ControlSet001\Control\CrashControl",
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
                        RegistryPath = @"ControlSet001\Control\CrashControl",
                        RegistryValue = $"MinidumpDir = {minidumpDir}"
                    });
                }

                // AutoReboot
                var autoReboot = crashControlKey.Values.FirstOrDefault(v => v.ValueName == "AutoReboot")?.ValueData?.ToString() ?? "";
                dumpSection.Items.Add(new AnalysisItem 
                { 
                    Name = "AutoReboot", 
                    Value = autoReboot == "1" ? "1 - Enabled" : (autoReboot == "0" ? "0 - Disabled" : "Not Configured"),
                    RegistryPath = @"ControlSet001\Control\CrashControl",
                    RegistryValue = $"AutoReboot = {autoReboot}"
                });

                // Overwrite
                var overwrite = crashControlKey.Values.FirstOrDefault(v => v.ValueName == "Overwrite")?.ValueData?.ToString() ?? "";
                dumpSection.Items.Add(new AnalysisItem 
                { 
                    Name = "Overwrite", 
                    Value = overwrite == "1" ? "1 - Overwrite existing file" : (overwrite == "0" ? "0 - Keep existing file" : "Not Configured"),
                    RegistryPath = @"ControlSet001\Control\CrashControl",
                    RegistryValue = $"Overwrite = {overwrite}"
                });

                // LogEvent
                var logEvent = crashControlKey.Values.FirstOrDefault(v => v.ValueName == "LogEvent")?.ValueData?.ToString() ?? "";
                dumpSection.Items.Add(new AnalysisItem 
                { 
                    Name = "LogEvent", 
                    Value = logEvent == "1" ? "1 - Log to System Event Log" : (logEvent == "0" ? "0 - Do not log" : "Not Configured"),
                    RegistryPath = @"ControlSet001\Control\CrashControl",
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
                        RegistryPath = @"ControlSet001\Control\CrashControl",
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
                        RegistryPath = @"ControlSet001\Control\CrashControl",
                        RegistryValue = $"DumpFilters = {dumpFilters}"
                    });
                }
            }
            sections.Add(dumpSection); // Always add for UI consistency

            // Guest Agent section - content depends on hive type
            var guestSection = new AnalysisSection { Title = "☁️ Guest Agent" };
            
            if (_parser.CurrentHiveType == OfflineRegistryParser.HiveType.SOFTWARE)
            {
                // SOFTWARE hive: Show VmId, VmType, and grayed out agent status
                const string azurePath = @"Microsoft\Windows Azure";
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
                
                // Show grayed out agent status with hint to load SYSTEM hive
                guestSection.Items.Add(new AnalysisItem
                {
                    Name = "WindowsAzureGuestAgent",
                    Value = "(Load SYSTEM hive to view)",
                    RegistryPath = @"ControlSet001\Services\WindowsAzureGuestAgent",
                    RegistryValue = "Agent service information requires SYSTEM hive.\nLoad the SYSTEM hive to see agent version and path."
                });
                
                guestSection.Items.Add(new AnalysisItem
                {
                    Name = "RdAgent",
                    Value = "(Load SYSTEM hive to view)",
                    RegistryPath = @"ControlSet001\Services\RdAgent",
                    RegistryValue = "Agent service information requires SYSTEM hive.\nLoad the SYSTEM hive to see agent version and path."
                });
                
                // Extensions are now shown via separate sub-tab (↳ Extensions)
            }
            else
            {
                // SYSTEM hive: Show agent service info
                const string waGuestAgentPath = @"ControlSet001\Services\WindowsAzureGuestAgent";
                var waImagePath = GetValue(waGuestAgentPath, "ImagePath") ?? "";
                var waVersion = ExtractAgentVersionInfo(waImagePath);
                guestSection.Items.Add(new AnalysisItem
                {
                    Name = "WindowsAzureGuestAgent",
                    Value = string.IsNullOrEmpty(waImagePath) ? "Not Found" : (string.IsNullOrEmpty(waVersion) ? TruncatePath(waImagePath, 100) : waVersion),
                    RegistryPath = waGuestAgentPath,
                    RegistryValue = string.IsNullOrEmpty(waImagePath) ? "ImagePath missing" : $"ImagePath = {TruncatePath(waImagePath, 140)}"
                });

                const string rdAgentPath = @"ControlSet001\Services\RdAgent";
                var rdImagePath = GetValue(rdAgentPath, "ImagePath") ?? "";
                var rdVersion = ExtractAgentVersionInfo(rdImagePath);
                guestSection.Items.Add(new AnalysisItem
                {
                    Name = "RdAgent",
                    Value = string.IsNullOrEmpty(rdImagePath) ? "Not Found" : (string.IsNullOrEmpty(rdVersion) ? TruncatePath(rdImagePath, 100) : rdVersion),
                    RegistryPath = rdAgentPath,
                    RegistryValue = string.IsNullOrEmpty(rdImagePath) ? "ImagePath missing" : $"ImagePath = {TruncatePath(rdImagePath, 140)}"
                });
                
                // Show hint about VmId in SOFTWARE hive
                guestSection.Items.Add(new AnalysisItem
                {
                    Name = "VmId",
                    Value = "(Load SOFTWARE hive to view)",
                    RegistryPath = @"Microsoft\Windows Azure",
                    RegistryValue = "VM identification requires SOFTWARE hive.\nLoad the SOFTWARE hive to see VmId and Extensions."
                });
            }

            sections.Add(guestSection);

            // System Time Configuration (SYSTEM hive) - Windows Time Service (w32time)
            sections.Add(GetSystemTimeAnalysis());

            return sections;
        }

        /// <summary>
        /// Get Windows Time Service (w32time) configuration as a structured section
        /// </summary>
        public AnalysisSection GetSystemTimeAnalysis()
        {
            var timeSection = new AnalysisSection { Title = "🕐 System Time Config" };

            // w32time Parameters - Type (NTP vs NT5DS)
            const string w32timeParamsPath = @"ControlSet001\Services\w32time\Parameters";
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
            const string w32timeConfigPath = @"ControlSet001\Services\w32time\Config";
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
            const string vmicProviderPath = @"ControlSet001\Services\w32time\TimeProviders\VMICTimeProvider";
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
            const string ntpClientPath = @"ControlSet001\Services\w32time\TimeProviders\NtpClient";
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
            const string ntpServerPath = @"ControlSet001\Services\w32time\TimeProviders\NtpServer";
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
            const string w32timeServicePath = @"ControlSet001\Services\w32time";
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
            
            const string handlerStatePath = @"Microsoft\Windows Azure\HandlerState";
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

            // Local Users (SAM hive)
            var usersPath = @"SAM\Domains\Account\Users\Names";
            var usersKey = _parser.GetKey(usersPath);
            if (usersKey?.SubKeys != null && usersKey.SubKeys.Count > 0)
            {
                var usersSection = new AnalysisSection { Title = "👤 Local User Accounts" };
                foreach (var userKey in usersKey.SubKeys.OrderBy(k => k.KeyName))
                {
                    usersSection.Items.Add(new AnalysisItem 
                    { 
                        Name = userKey.KeyName, 
                        Value = "Local Account",
                        RegistryPath = $"{usersPath}\\{userKey.KeyName}",
                        RegistryValue = $"User account entry in SAM database"
                    });
                }
                sections.Add(usersSection);
            }

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
                if (profilesSection.Items.Count > 0)
                    sections.Add(profilesSection);
            }

            // Recent Documents (NTUSER.DAT)
            var recentDocsPath = @"Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs";
            var recentDocs = _parser.GetKey(recentDocsPath);
            if (recentDocs?.SubKeys != null)
            {
                var recentSection = new AnalysisSection { Title = "📄 Recent Documents" };
                recentSection.Items.Add(new AnalysisItem 
                { 
                    Name = "Total Extensions", 
                    Value = recentDocs.SubKeys.Count.ToString(),
                    RegistryPath = recentDocsPath,
                    RegistryValue = $"Contains {recentDocs.SubKeys.Count} file extension subkeys"
                });
                foreach (var ext in recentDocs.SubKeys.Take(10))
                {
                    recentSection.Items.Add(new AnalysisItem 
                    { 
                        Name = ext.KeyName, 
                        Value = $"{ext.Values.Count} files",
                        RegistryPath = $"{recentDocsPath}\\{ext.KeyName}",
                        RegistryValue = $"Extension: {ext.KeyName}\nFiles tracked: {ext.Values.Count}"
                    });
                }
                sections.Add(recentSection);
            }

            // Typed Paths (NTUSER.DAT)
            var typedPathsPath = @"Software\Microsoft\Windows\CurrentVersion\Explorer\TypedPaths";
            var typedPaths = _parser.GetKey(typedPathsPath);
            if (typedPaths != null && typedPaths.Values.Count > 0)
            {
                var pathsSection = new AnalysisSection { Title = "📍 Explorer Typed Paths" };
                foreach (var value in typedPaths.Values.Where(v => !string.IsNullOrEmpty(v.ValueData?.ToString())).Take(15))
                {
                    pathsSection.Items.Add(new AnalysisItem 
                    { 
                        Name = value.ValueName, 
                        Value = value.ValueData?.ToString() ?? "",
                        RegistryPath = typedPathsPath,
                        RegistryValue = $"{value.ValueName} = {value.ValueData}"
                    });
                }
                sections.Add(pathsSection);
            }

            // Run MRU (NTUSER.DAT)
            var runMruPath = @"Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU";
            var runMru = _parser.GetKey(runMruPath);
            if (runMru != null && runMru.Values.Count > 1)
            {
                var runSection = new AnalysisSection { Title = "▶️ Run Dialog History" };
                foreach (var value in runMru.Values.Where(v => v.ValueName != "MRUList" && !string.IsNullOrEmpty(v.ValueData?.ToString())).Take(10))
                {
                    var cmd = value.ValueData?.ToString()?.TrimEnd('\\', '1') ?? "";
                    runSection.Items.Add(new AnalysisItem 
                    { 
                        Name = value.ValueName, 
                        Value = cmd,
                        RegistryPath = runMruPath,
                        RegistryValue = $"{value.ValueName} = {value.ValueData}"
                    });
                }
                sections.Add(runSection);
            }

            // UserAssist (NTUSER.DAT)
            var userAssistPath = @"Software\Microsoft\Windows\CurrentVersion\Explorer\UserAssist";
            var userAssist = _parser.GetKey(userAssistPath);
            if (userAssist?.SubKeys != null && userAssist.SubKeys.Count > 0)
            {
                var uaSection = new AnalysisSection { Title = "📊 UserAssist (Program Usage)" };
                int programCount = 0;
                foreach (var guid in userAssist.SubKeys)
                {
                    var countKey = guid.SubKeys?.FirstOrDefault(k => k.KeyName == "Count");
                    if (countKey != null)
                        programCount += countKey.Values.Count;
                }
                uaSection.Items.Add(new AnalysisItem 
                { 
                    Name = "Tracked Programs", 
                    Value = programCount.ToString(),
                    RegistryPath = userAssistPath,
                    RegistryValue = $"Total tracked program executions: {programCount}\nData is ROT13 encoded in registry"
                });
                sections.Add(uaSection);
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
            var servicesKey = _parser.GetKey(@"ControlSet001\Services");
            
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
                        DisplayName = displayName,
                        Description = description,
                        StartType = startValue,
                        StartTypeName = isDelayedAuto ? "Automatic (Delayed)" : GetStartTypeName(startValue),
                        ImagePath = imagePath,
                        RegistryPath = $@"ControlSet001\Services\{k.KeyName}",
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
            const string tsPath = @"ControlSet001\Control\Terminal Server";
            const string rdpTcpPath = @"ControlSet001\Control\Terminal Server\WinStations\RDP-Tcp";
            const string termServicePath = @"ControlSet001\Services\TermService";

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
                const string winStationsPathForCert = @"ControlSet001\Control\Terminal Server\WinStations";
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
                    const string tsPolicyPath = @"SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services";
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
            const string winStationsPath = @"ControlSet001\Control\Terminal Server\WinStations";
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
            const string licensingPath = @"ControlSet001\Control\Terminal Server\Licensing Core";
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
            const string termServiceParamsPath = @"ControlSet001\Services\TermService\Parameters";
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
            const string rcmLicensingPath = @"ControlSet001\Control\Terminal Server\RCM\Licensing Core";
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
            const string rdsPolicyPath = @"ControlSet001\Control\Terminal Server\RCM";
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
            var classKey = _parser.GetKey(@"ControlSet001\Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}");
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
                        var classPath = $@"ControlSet001\Control\Class\{{4d36e972-e325-11ce-bfc1-08002be10318}}\{adapter.KeyName}";
                        
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
            var enumKey = _parser.GetKey(@"ControlSet001\Enum");
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
                                    var enumPath = $@"ControlSet001\Enum\{deviceInstancePath}";
                                    
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
            var interfacesKey = _parser.GetKey(@"ControlSet001\Services\Tcpip\Parameters\Interfaces");
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

                    var registryPath = $@"ControlSet001\Services\Tcpip\Parameters\Interfaces\{iface.KeyName}";

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
                        var dnsRegAdapterPath = $@"ControlSet001\Services\Tcpip\Parameters\DNSRegisteredAdapters\{iface.KeyName}";
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
            const string dnsRegBasePath = @"ControlSet001\Services\Tcpip\Parameters\DNSRegisteredAdapters";
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

            // Network Profiles (SOFTWARE hive)
            var profilesKey = _parser.GetKey(@"Microsoft\Windows NT\CurrentVersion\NetworkList\Profiles");
            if (profilesKey?.SubKeys != null && profilesKey.SubKeys.Count > 0)
            {
                var profilesSection = new AnalysisSection { Title = "📶 Network Profiles" };
                foreach (var profile in profilesKey.SubKeys.Take(20))
                {
                    var name = profile.Values.FirstOrDefault(v => v.ValueName == "ProfileName")?.ValueData?.ToString();
                    var category = profile.Values.FirstOrDefault(v => v.ValueName == "Category")?.ValueData?.ToString();
                    var description = profile.Values.FirstOrDefault(v => v.ValueName == "Description")?.ValueData?.ToString();
                    var managed = profile.Values.FirstOrDefault(v => v.ValueName == "Managed")?.ValueData?.ToString();

                    if (!string.IsNullOrEmpty(name))
                    {
                        var registryPath = $@"Microsoft\Windows NT\CurrentVersion\NetworkList\Profiles\{profile.KeyName}";
                        var details = new List<string>();
                        details.Add($"Category: {GetNetworkCategory(category ?? "")}");
                        if (!string.IsNullOrEmpty(description)) details.Add($"Description: {description}");
                        details.Add($"Managed: {(managed == "1" ? "Yes" : "No")}");

                        profilesSection.Items.Add(new AnalysisItem
                        {
                            Name = name,
                            Value = GetNetworkCategory(category ?? ""),
                            RegistryPath = registryPath,
                            RegistryValue = string.Join(" | ", details)
                        });
                    }
                }

                if (profilesSection.Items.Count > 0)
                    sections.Add(profilesSection);
            }

            // WiFi Networks (SOFTWARE hive)
            var wlanKey = _parser.GetKey(@"Microsoft\WlanSvc\Profiles");
            if (wlanKey?.SubKeys != null && wlanKey.SubKeys.Count > 0)
            {
                var wifiSection = new AnalysisSection { Title = "📡 WiFi Networks" };
                wifiSection.Items.Add(new AnalysisItem 
                { 
                    Name = "Saved Profiles", 
                    Value = wlanKey.SubKeys.Count.ToString(),
                    RegistryPath = @"Microsoft\WlanSvc\Profiles",
                    RegistryValue = $"Total saved WiFi profiles: {wlanKey.SubKeys.Count}"
                });
                sections.Add(wifiSection);
            }

            // Shares (SYSTEM hive)
            var sharesKey = _parser.GetKey(@"ControlSet001\Services\LanmanServer\Shares");
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
                        RegistryPath = @"ControlSet001\Services\LanmanServer\Shares",
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
                    RegistryPath = @"ControlSet001\Services\LanmanServer\Shares",
                    RegistryValue = "No shares found in the registry"
                });
            }
            sections.Add(sharesSection);

            // NTLM Authentication Settings (SYSTEM hive)
            var ntlmSection = new AnalysisSection { Title = "🔑 NTLM Authentication" };
            const string lsaPath = @"ControlSet001\Control\Lsa";
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
                const string msv1Path = @"ControlSet001\Control\Lsa\MSV1_0";
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
            var schannelPath = @"ControlSet001\Control\SecurityProviders\SCHANNEL\Protocols";
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
            const string firewallBasePath = @"ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy";
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
            if (string.IsNullOrEmpty(str) || str.Length <= maxLength)
                return str;
            return str.Substring(0, maxLength - 3) + "...";
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
                // Only show last 5 chars for privacy
                var maskedKey = backupKey.Length > 5 
                    ? new string('*', backupKey.Length - 5) + backupKey.Substring(backupKey.Length - 5) 
                    : backupKey;
                section.Items.Add(new AnalysisItem
                {
                    Name = "BackupProductKeyDefault",
                    Value = maskedKey,
                    RegistryPath = backupProductKeyPath,
                    RegistryValue = $"BackupProductKeyDefault = {maskedKey}"
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

            return sections;
        }

        /// <summary>
        /// Get disk class filter drivers configuration
        /// </summary>
        public List<AnalysisItem> GetDiskFilters()
        {
            var items = new List<AnalysisItem>();

            // Disk class GUID: {4d36e967-e325-11ce-bfc1-08002be10318}
            const string diskClassPath = @"ControlSet001\Control\Class\{4d36e967-e325-11ce-bfc1-08002be10318}";
            var diskClassKey = _parser.GetKey(diskClassPath);

            if (diskClassKey != null)
            {
                // UpperFilters - filter drivers loaded above the disk class driver
                var upperFilters = diskClassKey.Values.FirstOrDefault(v => v.ValueName == "UpperFilters")?.ValueData?.ToString() ?? "";
                if (!string.IsNullOrEmpty(upperFilters))
                {
                    var filterList = ParseMultiSzValue(upperFilters);
                    items.Add(new AnalysisItem
                    {
                        Name = "UpperFilters",
                        Value = filterList,
                        RegistryPath = diskClassPath,
                        RegistryValue = $"UpperFilters = {upperFilters}"
                    });

                    // Add individual filter details
                    foreach (var filter in filterList.Split(new[] { ", " }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        var filterServicePath = $@"ControlSet001\Services\{filter.Trim()}";
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
                        RegistryPath = diskClassPath,
                        RegistryValue = "UpperFilters not present"
                    });
                }

                // LowerFilters - filter drivers loaded below the disk class driver
                var lowerFilters = diskClassKey.Values.FirstOrDefault(v => v.ValueName == "LowerFilters")?.ValueData?.ToString() ?? "";
                if (!string.IsNullOrEmpty(lowerFilters))
                {
                    var filterList = ParseMultiSzValue(lowerFilters);
                    items.Add(new AnalysisItem
                    {
                        Name = "LowerFilters",
                        Value = filterList,
                        RegistryPath = diskClassPath,
                        RegistryValue = $"LowerFilters = {lowerFilters}"
                    });

                    // Add individual filter details
                    foreach (var filter in filterList.Split(new[] { ", " }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        var filterServicePath = $@"ControlSet001\Services\{filter.Trim()}";
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
                        RegistryPath = diskClassPath,
                        RegistryValue = "LowerFilters not present"
                    });
                }
            }
            else
            {
                items.Add(new AnalysisItem
                {
                    Name = "Status",
                    Value = "Disk class configuration not found",
                    RegistryPath = diskClassPath,
                    RegistryValue = "Key not present - requires SYSTEM hive"
                });
            }

            return items;
        }

        /// <summary>
        /// Get volume class filter drivers configuration
        /// </summary>
        public List<AnalysisItem> GetVolumeFilters()
        {
            var items = new List<AnalysisItem>();

            // Volume class GUID: {71A27CDD-812A-11D0-BEC7-08002BE2092F}
            const string volumeClassPath = @"ControlSet001\Control\Class\{71A27CDD-812A-11D0-BEC7-08002BE2092F}";
            var volumeClassKey = _parser.GetKey(volumeClassPath);

            if (volumeClassKey != null)
            {
                // UpperFilters - filter drivers loaded above the volume class driver
                var upperFilters = volumeClassKey.Values.FirstOrDefault(v => v.ValueName == "UpperFilters")?.ValueData?.ToString() ?? "";
                if (!string.IsNullOrEmpty(upperFilters))
                {
                    var filterList = ParseMultiSzValue(upperFilters);
                    items.Add(new AnalysisItem
                    {
                        Name = "UpperFilters",
                        Value = filterList,
                        RegistryPath = volumeClassPath,
                        RegistryValue = $"UpperFilters = {upperFilters}"
                    });

                    // Add individual filter details
                    foreach (var filter in filterList.Split(new[] { ", " }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        var filterServicePath = $@"ControlSet001\Services\{filter.Trim()}";
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
                        RegistryPath = volumeClassPath,
                        RegistryValue = "UpperFilters not present"
                    });
                }

                // LowerFilters - filter drivers loaded below the volume class driver
                var lowerFilters = volumeClassKey.Values.FirstOrDefault(v => v.ValueName == "LowerFilters")?.ValueData?.ToString() ?? "";
                if (!string.IsNullOrEmpty(lowerFilters))
                {
                    var filterList = ParseMultiSzValue(lowerFilters);
                    items.Add(new AnalysisItem
                    {
                        Name = "LowerFilters",
                        Value = filterList,
                        RegistryPath = volumeClassPath,
                        RegistryValue = $"LowerFilters = {lowerFilters}"
                    });

                    // Add individual filter details
                    foreach (var filter in filterList.Split(new[] { ", " }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        var filterServicePath = $@"ControlSet001\Services\{filter.Trim()}";
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
                        RegistryPath = volumeClassPath,
                        RegistryValue = "LowerFilters not present"
                    });
                }
            }
            else
            {
                items.Add(new AnalysisItem
                {
                    Name = "Status",
                    Value = "Volume class configuration not found",
                    RegistryPath = volumeClassPath,
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

        #endregion

        #region Software Analysis (SOFTWARE hive)

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

            return sections;
        }

        /// <summary>
        /// Get combined Appx section (placeholder for UI with filter buttons)
        /// </summary>
        public AnalysisSection GetAppxAnalysis()
        {
            var section = new AnalysisSection { Title = "📱 Appx Packages" };
            
            // Get counts for display
            const string inboxPath = @"Microsoft\Windows\CurrentVersion\Appx\AppxAllUserStore\InboxApplications";
            const string applicationsPath = @"Microsoft\Windows\CurrentVersion\Appx\AppxAllUserStore\Applications";
            
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

            const string runPath = @"Microsoft\Windows\CurrentVersion\Run";
            const string runOncePath = @"Microsoft\Windows\CurrentVersion\RunOnce";

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
            const string x64Path = @"Microsoft\Windows\CurrentVersion\Uninstall";
            // x86 (WOW6432Node) programs path
            const string x86Path = @"WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall";

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
            const string firewallBasePath = @"ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy";
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
            const string rulesPath = @"ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy\FirewallRules";
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
                var rule = new FirewallRuleInfo { RuleId = ruleName };
                
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
                            rule.LocalPorts = val;
                            break;
                        case "rport":
                            rule.RemotePorts = val;
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
                    }
                }

                // Always use the registry value name as the rule name (it's the unique identifier)
                rule.Name = ruleName;

                return rule;
            }
            catch
            {
                return null;
            }
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
        public string RegistryPath { get; set; } = "";
        public string RegistryValueName { get; set; } = "";

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
}

using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows;
using Microsoft.Win32;
using RegistryExpert.Core;
using RegistryExpert.Wpf.Helpers;
using RegistryParser.Abstractions;
using HiveType = RegistryExpert.Core.OfflineRegistryParser.HiveType;
using RegistryKey = RegistryParser.Abstractions.RegistryKey;

namespace RegistryExpert.Wpf.ViewModels
{
    public class MainViewModel : ViewModelBase
    {
        // ── Backing fields ──────────────────────────────────────────────────

        private string _windowTitle = "Registry Expert";
        private string _statusText = "Ready";
        private string _hiveStatusText = "";
        private bool _isLoading;
        private double _loadProgress;
        private bool _hasLoadedHives;
        private RegistryKeyNode? _selectedTreeNode;
        private RegistryValueItem? _selectedValue;
        private string _detailsText = "";
        private bool _isBookmarkPanelExpanded;
        private bool _hasBookmarks;

        private readonly Dictionary<HiveType, LoadedHiveInfo> _loadedHives = new();
        private CancellationTokenSource? _loadCts;
        private readonly AppSettings _settings;

        // ── Properties ──────────────────────────────────────────────────────

        public string WindowTitle
        {
            get => _windowTitle;
            set => SetProperty(ref _windowTitle, value);
        }

        public string StatusText
        {
            get => _statusText;
            set => SetProperty(ref _statusText, value);
        }

        public string HiveStatusText
        {
            get => _hiveStatusText;
            set => SetProperty(ref _hiveStatusText, value);
        }

        public bool IsLoading
        {
            get => _isLoading;
            set => SetProperty(ref _isLoading, value);
        }

        public double LoadProgress
        {
            get => _loadProgress;
            set => SetProperty(ref _loadProgress, value);
        }

        public bool HasLoadedHives
        {
            get => _hasLoadedHives;
            set => SetProperty(ref _hasLoadedHives, value);
        }

        public ObservableCollection<RegistryKeyNode> TreeRoots { get; } = new();

        public RegistryKeyNode? SelectedTreeNode
        {
            get => _selectedTreeNode;
            set
            {
                if (SetProperty(ref _selectedTreeNode, value))
                    OnTreeNodeSelected();
            }
        }

        public ObservableCollection<RegistryValueItem> Values { get; } = new();

        public RegistryValueItem? SelectedValue
        {
            get => _selectedValue;
            set
            {
                if (SetProperty(ref _selectedValue, value))
                    OnValueSelected();
            }
        }

        public string DetailsText
        {
            get => _detailsText;
            set => SetProperty(ref _detailsText, value);
        }

        public ObservableCollection<BookmarkItem> Bookmarks { get; } = new();

        public bool IsBookmarkPanelExpanded
        {
            get => _isBookmarkPanelExpanded;
            set => SetProperty(ref _isBookmarkPanelExpanded, value);
        }

        public bool HasBookmarks
        {
            get => _hasBookmarks;
            set => SetProperty(ref _hasBookmarks, value);
        }

        public IReadOnlyDictionary<HiveType, LoadedHiveInfo> LoadedHives => _loadedHives;

        /// <summary>
        /// Find the LoadedHiveInfo for the currently selected tree node by walking up to its root.
        /// </summary>
        public LoadedHiveInfo? ActiveHive
        {
            get
            {
                var node = SelectedTreeNode;
                if (node == null) return null;

                // Walk up to the root node by checking TreeRoots
                var rootNode = FindRootNode(node);
                if (rootNode == null) return null;

                foreach (var hive in _loadedHives.Values)
                {
                    if (hive.RootNode == rootNode)
                        return hive;
                }
                return null;
            }
        }

        public AppSettings Settings => _settings;

        // ── Commands ────────────────────────────────────────────────────────

        public AsyncRelayCommand OpenHiveCommand { get; }
        public RelayCommand CloseHiveCommand { get; }
        public RelayCommand ExportKeyCommand { get; }
        public RelayCommand CopyPathCommand { get; }
        public RelayCommand CopyValueCommand { get; }
        public RelayCommand CancelLoadCommand { get; }
        public RelayCommand ToggleBookmarksCommand { get; }
        public RelayCommand NavigateToKeyCommand { get; }
        public RelayCommand SwitchThemeCommand { get; }
        public RelayCommand ExitCommand { get; }
        public RelayCommand OpenSearchCommand { get; }
        public RelayCommand OpenAnalyzeCommand { get; }
        public RelayCommand OpenStatisticsCommand { get; }
        public RelayCommand OpenCompareCommand { get; }
        public RelayCommand OpenTimelineCommand { get; }
        public RelayCommand AboutCommand { get; }
        public AsyncRelayCommand CheckForUpdatesCommand { get; }

        // ── Events ──────────────────────────────────────────────────────────

        /// <summary>Raised when the View should open the Search window.</summary>
        public event Action? RequestOpenSearch;

        /// <summary>Raised when the View should open the Analyze window.</summary>
        public event Action? RequestOpenAnalyze;

        /// <summary>Raised when the View should open the Statistics window.</summary>
        public event Action? RequestOpenStatistics;

        /// <summary>Raised when the View should open the Compare window.</summary>
        public event Action? RequestOpenCompare;

        /// <summary>Raised when the View should open the Timeline window.</summary>
        public event Action? RequestOpenTimeline;

        /// <summary>Raised when the View should open the About dialog.</summary>
        public event Action? RequestOpenAbout;

        /// <summary>Raised when the View should show the update check result.</summary>
        public event Action<UpdateInfo?, bool>? RequestShowUpdateResult;

        /// <summary>Raised when the View should scroll a navigated node (and optional value) into view.</summary>
        public event Action<RegistryKeyNode, string?>? RequestScrollToNode;

        // ── Constructor ─────────────────────────────────────────────────────

        public MainViewModel()
        {
            _settings = AppSettings.Load();

            OpenHiveCommand = new AsyncRelayCommand(OnOpenHive);
            CloseHiveCommand = new RelayCommand(p => OnCloseHive(p));
            ExportKeyCommand = new RelayCommand(OnExportKey, CanExportKey);
            CopyPathCommand = new RelayCommand(OnCopyPath, CanCopyPath);
            CopyValueCommand = new RelayCommand(OnCopyValue, CanCopyValue);
            CancelLoadCommand = new RelayCommand(OnCancelLoad, () => IsLoading);
            ToggleBookmarksCommand = new RelayCommand(() => IsBookmarkPanelExpanded = !IsBookmarkPanelExpanded);
            NavigateToKeyCommand = new RelayCommand(p => OnNavigateToKey(p));
            SwitchThemeCommand = new RelayCommand(p => OnSwitchTheme(p));
            ExitCommand = new RelayCommand(() => Application.Current.Shutdown());
            OpenSearchCommand = new RelayCommand(() => RequestOpenSearch?.Invoke(), () => HasLoadedHives);
            OpenAnalyzeCommand = new RelayCommand(() => RequestOpenAnalyze?.Invoke(), () => HasLoadedHives);
            OpenStatisticsCommand = new RelayCommand(() => RequestOpenStatistics?.Invoke(), () => HasLoadedHives);
            OpenCompareCommand = new RelayCommand(() => RequestOpenCompare?.Invoke(), () => HasLoadedHives);
            OpenTimelineCommand = new RelayCommand(() => RequestOpenTimeline?.Invoke(), () => HasLoadedHives);
            AboutCommand = new RelayCommand(() => RequestOpenAbout?.Invoke());
            CheckForUpdatesCommand = new AsyncRelayCommand(OnCheckForUpdates);
        }

        // ── Command handlers ────────────────────────────────────────────────

        private async Task OnOpenHive(object? _)
        {
            var dlg = new OpenFileDialog
            {
                Title = "Open Registry Hive",
                Filter = "Registry Hive Files|*.dat;SYSTEM;SOFTWARE;SAM;SECURITY;NTUSER*;USRCLASS*;DEFAULT;AMCACHE*;BCD;COMPONENTS|All Files|*.*",
                FilterIndex = 2,
                CheckFileExists = true,
                Multiselect = true
            };

            if (dlg.ShowDialog() == true)
            {
                foreach (var file in dlg.FileNames)
                {
                    await LoadHiveFileAsync(file);
                }
            }
        }

        private void OnCloseHive(object? parameter)
        {
            if (parameter is HiveType hiveType)
            {
                CloseHive(hiveType);
            }
            else if (parameter is string hiveTypeStr && Enum.TryParse<HiveType>(hiveTypeStr, true, out var parsed))
            {
                CloseHive(parsed);
            }
        }

        private void OnExportKey()
        {
            var node = SelectedTreeNode;
            if (node?.RegistryKey == null) return;

            var dlg = new SaveFileDialog
            {
                Title = "Export Registry Key",
                Filter = "Text Files|*.txt|All Files|*.*",
                DefaultExt = ".txt",
                FileName = $"{node.DisplayName}_export.txt"
            };

            if (dlg.ShowDialog() == true)
            {
                try
                {
                    using var writer = new StreamWriter(dlg.FileName, false, Encoding.UTF8);

                    var hive = ActiveHive;
                    var displayPath = hive != null
                        ? hive.Parser.ConvertRootPath(node.RegistryKey.KeyPath)
                        : node.RegistryKey.KeyPath;
                    writer.WriteLine($"Registry Export: {displayPath}");
                    writer.WriteLine($"Exported: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                    writer.WriteLine(new string('=', 60));
                    writer.WriteLine();

                    ExportKeyRecursive(node.RegistryKey, writer, 0);

                    StatusText = $"Exported to {Path.GetFileName(dlg.FileName)}";
                }
                catch (Exception ex)
                {
                    StatusText = $"Export failed: {ex.Message}";
                    Debug.WriteLine($"Export error: {ex.Message}");
                }
            }
        }

        private bool CanExportKey() => SelectedTreeNode?.RegistryKey != null;

        private void OnCopyPath()
        {
            var node = SelectedTreeNode;
            if (node?.RegistryKey == null) return;

            var hive = ActiveHive;
            var path = hive != null
                ? hive.Parser.ConvertRootPath(node.RegistryKey.KeyPath)
                : node.RegistryKey.KeyPath;

            Clipboard.SetText(path);
            StatusText = "Path copied to clipboard";
        }

        private bool CanCopyPath() => SelectedTreeNode?.RegistryKey != null;

        private void OnCopyValue()
        {
            var val = SelectedValue;
            if (val == null) return;

            Clipboard.SetText(val.Data ?? "");
            StatusText = "Value copied to clipboard";
        }

        private bool CanCopyValue() => SelectedValue != null;

        private void OnCancelLoad()
        {
            _loadCts?.Cancel();
        }

        private void OnNavigateToKey(object? parameter)
        {
            if (parameter is string path && !string.IsNullOrWhiteSpace(path))
            {
                NavigateToKey(path);
            }
        }

        private void OnSwitchTheme(object? parameter)
        {
            var theme = parameter?.ToString() == "Dark"
                ? ThemeManager.Theme.Dark
                : ThemeManager.Theme.Light;
            ThemeManager.SetTheme(theme);

            // Persist theme choice
            _settings.Theme = parameter?.ToString() ?? "Dark";
            _settings.Save();
        }

        private async Task OnCheckForUpdates(object? _)
        {
            StatusText = "Checking for updates...";
            try
            {
                var info = await UpdateChecker.CheckForUpdatesAsync();
                RequestShowUpdateResult?.Invoke(info, true);
            }
            catch
            {
                RequestShowUpdateResult?.Invoke(null, true);
            }
            finally
            {
                StatusText = "Ready";
            }
        }

        // ── Core methods ────────────────────────────────────────────────────

        /// <summary>
        /// Load a registry hive file asynchronously with progress and cancellation support.
        /// </summary>
        public async Task LoadHiveFileAsync(string filePath)
        {
            // Cancel any existing load operation
            _loadCts?.Cancel();
            _loadCts?.Dispose();
            _loadCts = new CancellationTokenSource();
            var ct = _loadCts.Token;

            IsLoading = true;
            LoadProgress = 0;
            StatusText = $"Loading {Path.GetFileName(filePath)}...";

            var parser = new OfflineRegistryParser();
            try
            {
                var progress = new Progress<(string phase, double percent)>(p =>
                {
                    LoadProgress = p.percent * 100;
                    StatusText = $"Loading: {p.phase} ({p.percent * 100:F0}%)";
                });

                var success = await Task.Run(() => parser.LoadHive(filePath, progress, ct), ct)
                    .ConfigureAwait(true);

                if (!success)
                {
                    parser.Dispose();
                    StatusText = "Failed to load hive file";
                    return;
                }

                ct.ThrowIfCancellationRequested();

                var rootKey = parser.GetRootKey();
                if (rootKey == null)
                {
                    parser.Dispose();
                    StatusText = "Hive has no root key";
                    return;
                }

                var hiveType = parser.CurrentHiveType;
                var hiveName = hiveType.ToString();

                // If this hive type is already loaded, ask to replace
                if (_loadedHives.TryGetValue(hiveType, out var existing))
                {
                    var existingFile = Path.GetFileName(existing.FilePath);
                    var result = System.Windows.MessageBox.Show(
                        $"A {hiveName} hive is already loaded ({existingFile}).\n\nReplace it?",
                        "Duplicate Hive Type",
                        System.Windows.MessageBoxButton.YesNo,
                        System.Windows.MessageBoxImage.Question);

                    if (result != System.Windows.MessageBoxResult.Yes)
                    {
                        parser.Dispose();
                        StatusText = "Load cancelled — duplicate hive type";
                        return;
                    }

                    TreeRoots.Remove(existing.RootNode);
                    existing.Dispose();
                    _loadedHives.Remove(hiveType);
                }

                var extractor = new RegistryInfoExtractor(parser);
                var rootNode = new RegistryKeyNode(rootKey, hiveName);
                var hiveInfo = new LoadedHiveInfo
                {
                    Parser = parser,
                    InfoExtractor = extractor,
                    FilePath = filePath,
                    RootNode = rootNode
                };

                _loadedHives[hiveType] = hiveInfo;
                TreeRoots.Add(rootNode);
                rootNode.IsExpanded = true;

                UpdateHiveSeparators();
                UpdateBookmarks();
                UpdateTitleBar();
                UpdateHiveStatus();
                HasLoadedHives = _loadedHives.Count > 0;

                StatusText = $"Loaded {hiveName} hive ({Path.GetFileName(filePath)})";
            }
            catch (OperationCanceledException)
            {
                parser.Dispose();
                StatusText = "Loading cancelled";
            }
            catch (Exception ex)
            {
                parser.Dispose();
                StatusText = $"Error loading hive: {ex.Message}";
                Debug.WriteLine($"LoadHive error: {ex}");
            }
            finally
            {
                IsLoading = false;
                LoadProgress = 0;
            }
        }

        /// <summary>
        /// Close and remove a loaded hive by type.
        /// </summary>
        public void CloseHive(HiveType hiveType)
        {
            if (!_loadedHives.TryGetValue(hiveType, out var hiveInfo))
                return;

            TreeRoots.Remove(hiveInfo.RootNode);
            _loadedHives.Remove(hiveType);
            hiveInfo.Dispose();

            UpdateHiveSeparators();
            UpdateBookmarks();
            UpdateTitleBar();
            UpdateHiveStatus();
            HasLoadedHives = _loadedHives.Count > 0;

            if (_loadedHives.Count == 0)
            {
                Values.Clear();
                DetailsText = "";
                SelectedTreeNode = null;
            }

            StatusText = $"Unloaded {hiveType} hive";

            // Help free memory from large hive files
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        /// Called when SelectedTreeNode changes — populate Values and show key details.
        /// </summary>
        private void OnTreeNodeSelected()
        {
            var node = SelectedTreeNode;
            if (node?.RegistryKey == null)
            {
                Values.Clear();
                DetailsText = "";
                return;
            }

            // Populate values: default value first, then alphabetical by name
            Values.Clear();
            var key = node.RegistryKey;
            if (key.Values != null)
            {
                var ordered = key.Values
                    .Select(v => new RegistryValueItem(v))
                    .OrderBy(v => string.IsNullOrEmpty(v.KeyValue.ValueName) ? 0 : 1)
                    .ThenBy(v => v.Name, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                foreach (var item in ordered)
                    Values.Add(item);
            }

            // Show key metadata in details panel
            var hive = ActiveHive;
            var displayPath = hive != null
                ? hive.Parser.ConvertRootPath(key.KeyPath)
                : key.KeyPath;

            var sb = new StringBuilder();
            sb.AppendLine("Key Details");
            sb.AppendLine(new string('─', 40));
            sb.AppendLine($"Path:       {displayPath}");
            sb.AppendLine($"Name:       {key.KeyName}");
            sb.AppendLine($"Modified:   {(key.LastWriteTime.HasValue ? key.LastWriteTime.Value.LocalDateTime.ToString("yyyy-MM-dd HH:mm:ss") : "Unknown")}");
            sb.AppendLine($"Subkeys:    {key.SubKeys?.Count ?? 0}");
            sb.AppendLine($"Values:     {key.Values?.Count ?? 0}");

            DetailsText = sb.ToString();
        }

        /// <summary>
        /// Called when SelectedValue changes — show value details or fall back to key details.
        /// </summary>
        private void OnValueSelected()
        {
            var val = SelectedValue;
            if (val == null)
            {
                // Revert to key details
                OnTreeNodeSelected();
                return;
            }

            var hive = ActiveHive;
            var node = SelectedTreeNode;
            var keyPath = node?.RegistryKey != null && hive != null
                ? hive.Parser.ConvertRootPath(node.RegistryKey.KeyPath)
                : node?.RegistryKey?.KeyPath ?? "";

            var sb = new StringBuilder();
            sb.AppendLine("Value Details");
            sb.AppendLine(new string('─', 40));
            sb.AppendLine($"Path:       {keyPath}");
            sb.AppendLine($"Name:       {val.Name}");
            sb.AppendLine($"Type:       {val.Type}");
            sb.AppendLine($"Slack:      {val.SlackSize} bytes");
            sb.AppendLine();
            sb.AppendLine("Data:");
            sb.AppendLine(val.Data);

            if (val.RawBytes.Length > 0)
            {
                sb.AppendLine();
                sb.AppendLine("Hex Dump:");
                sb.Append(FormatHexDump(val.RawBytes));
            }

            DetailsText = sb.ToString();
        }

        /// <summary>
        /// Format a byte array as a standard hex dump (16 bytes per line, offset + hex + ASCII).
        /// </summary>
        public static string FormatHexDump(byte[] data)
        {
            if (data == null || data.Length == 0) return "";

            var sb = new StringBuilder();
            var maxBytes = Math.Min(data.Length, 1024);

            for (int offset = 0; offset < maxBytes; offset += 16)
            {
                sb.Append($"{offset:X8}  ");

                // Hex section
                int lineLen = Math.Min(16, maxBytes - offset);
                for (int i = 0; i < 16; i++)
                {
                    if (i < lineLen)
                        sb.Append($"{data[offset + i]:X2} ");
                    else
                        sb.Append("   ");

                    if (i == 7) sb.Append(' ');
                }

                sb.Append(' ');

                // ASCII section
                for (int i = 0; i < lineLen; i++)
                {
                    byte b = data[offset + i];
                    sb.Append(b >= 0x20 && b <= 0x7E ? (char)b : '.');
                }

                sb.AppendLine();
            }

            if (data.Length > 1024)
            {
                sb.AppendLine($"... ({data.Length - 1024} more bytes truncated)");
            }

            return sb.ToString();
        }

        /// <summary>
        /// Recursively export a registry key and its values to a StreamWriter.
        /// </summary>
        private void ExportKeyRecursive(RegistryKey key, StreamWriter writer, int indent)
        {
            var prefix = new string(' ', indent * 2);

            writer.WriteLine($"{prefix}[{key.KeyName}]");

            if (key.LastWriteTime.HasValue)
                writer.WriteLine($"{prefix}  Last Modified: {key.LastWriteTime.Value.LocalDateTime:yyyy-MM-dd HH:mm:ss}");

            if (key.Values != null)
            {
                foreach (var value in key.Values)
                {
                    var name = string.IsNullOrEmpty(value.ValueName) ? "(Default)" : value.ValueName;
                    var type = value.ValueType ?? "Unknown";
                    var data = value.ValueData ?? "";
                    writer.WriteLine($"{prefix}  {name} ({type}) = {data}");
                }
            }

            writer.WriteLine();

            if (key.SubKeys != null)
            {
                foreach (var subKey in key.SubKeys.OrderBy(k => k.KeyName))
                {
                    ExportKeyRecursive(subKey, writer, indent + 1);
                }
            }
        }

        /// <summary>
        /// Repopulate Bookmarks collection from static definitions matching loaded hive types.
        /// </summary>
        private void UpdateBookmarks()
        {
            Bookmarks.Clear();

            foreach (var hiveType in _loadedHives.Keys.OrderBy(h => h.ToString()))
            {
                var hiveName = hiveType.ToString();
                if (_bookmarkDefinitions.TryGetValue(hiveName, out var defs))
                {
                    foreach (var (name, path) in defs)
                    {
                        Bookmarks.Add(new BookmarkItem
                        {
                            Name = name,
                            Path = path,
                            HiveType = hiveName
                        });
                    }
                }
            }

            HasBookmarks = Bookmarks.Count > 0;
        }

        /// <summary>
        /// Update the window title based on loaded hives.
        /// </summary>
        private void UpdateTitleBar()
        {
            if (_loadedHives.Count == 0)
            {
                WindowTitle = "Registry Expert";
            }
            else
            {
                var hiveNames = _loadedHives.Keys
                    .OrderBy(h => h.ToString())
                    .Select(h => h.ToString());
                WindowTitle = $"Registry Expert - {string.Join(", ", hiveNames)}";
            }
        }

        /// <summary>
        /// Update the hive status indicator text.
        /// </summary>
        private void UpdateHiveStatus()
        {
            if (_loadedHives.Count == 0)
            {
                HiveStatusText = "";
            }
            else if (_loadedHives.Count == 1)
            {
                HiveStatusText = $"● {_loadedHives.Keys.First()}";
            }
            else
            {
                HiveStatusText = $"● {_loadedHives.Count} hives";
            }
        }

        /// <summary>
        /// Update separator visibility on root hive nodes. The first root gets no separator;
        /// subsequent roots get a visible separator line above them.
        /// </summary>
        private void UpdateHiveSeparators()
        {
            for (int i = 0; i < TreeRoots.Count; i++)
            {
                TreeRoots[i].ShowSeparator = i > 0;
            }
        }

        /// <summary>
        /// Navigate the tree to a specific key path (e.g. "SYSTEM\ControlSet001\Services").
        /// Optionally select a value by name after navigation.
        /// </summary>
        public void NavigateToKey(string keyPath, string? valueName = null)
        {
            if (string.IsNullOrWhiteSpace(keyPath)) return;

            var parts = keyPath.Split('\\');
            if (parts.Length == 0) return;

            // Find the matching root node (first segment is the hive name)
            RegistryKeyNode? current = null;
            foreach (var root in TreeRoots)
            {
                if (string.Equals(root.DisplayName, parts[0], StringComparison.OrdinalIgnoreCase))
                {
                    current = root;
                    break;
                }
            }

            if (current == null) return;

            // Walk down the tree expanding as we go
            for (int i = 1; i < parts.Length; i++)
            {
                current.EnsureChildrenLoaded();
                current.IsExpanded = true;

                RegistryKeyNode? match = null;
                foreach (var child in current.Children)
                {
                    if (string.Equals(child.DisplayName, parts[i], StringComparison.OrdinalIgnoreCase))
                    {
                        match = child;
                        break;
                    }
                }

                if (match == null) break;
                current = match;
            }

            if (current != null)
            {
                current.IsSelected = true;
                SelectedTreeNode = current;

                // If a value name was provided, select it in the Values list
                if (valueName != null)
                {
                    foreach (var val in Values)
                    {
                        if (string.Equals(val.Name, valueName, StringComparison.OrdinalIgnoreCase) ||
                            (valueName == "(Default)" && val.Name == "(Default)") ||
                            (string.IsNullOrEmpty(valueName) && val.Name == "(Default)"))
                        {
                            SelectedValue = val;
                            break;
                        }
                    }
                }

                // Ask the View to scroll the tree node and value into view
                RequestScrollToNode?.Invoke(current, valueName);
            }
        }

        /// <summary>
        /// Find the root node (in TreeRoots) for a given node by walking up via TreeRoots matching.
        /// Since RegistryKeyNode doesn't have a Parent property, we search TreeRoots for the root
        /// whose subtree contains the target node's key path.
        /// </summary>
        private RegistryKeyNode? FindRootNode(RegistryKeyNode node)
        {
            if (node.IsRootNode)
                return node;

            // Match by key path prefix: the root node's RegistryKey.KeyPath should be
            // a prefix of the given node's KeyPath
            foreach (var root in TreeRoots)
            {
                if (root.RegistryKey == null) continue;

                var rootPath = root.RegistryKey.KeyPath;
                var nodePath = node.KeyPath;

                if (nodePath.StartsWith(rootPath, StringComparison.OrdinalIgnoreCase))
                    return root;
            }

            return null;
        }

        // ── Static bookmark definitions ─────────────────────────────────────

        private static readonly Dictionary<string, List<(string Name, string Path)>> _bookmarkDefinitions = new()
        {
            ["SOFTWARE"] = new()
            {
                ("Activation", @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"),
                ("Component Based Servicing", @"SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing"),
                ("Current Version", @"SOFTWARE\Microsoft\Windows NT\CurrentVersion"),
                ("Installed Programs", @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
                ("Logon UI", @"SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI"),
                ("Profile List", @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"),
                ("Startup Programs", @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run"),
                ("Windows Azure", @"SOFTWARE\Microsoft\Windows Azure"),
                ("Windows Defender", @"SOFTWARE\Microsoft\Windows Defender"),
                ("Windows Update", @"SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate"),
            },
            ["SYSTEM"] = new()
            {
                ("Class: Adapter", @"SYSTEM\ControlSet001\Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}"),
                ("Disk Filters", @"SYSTEM\ControlSet001\Control\Class\{4d36e967-e325-11ce-bfc1-08002be10318}"),
                ("Crash Control", @"SYSTEM\ControlSet001\Control\CrashControl"),
                ("Firewall", @"SYSTEM\ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy"),
                ("Guest Agent", @"SYSTEM\ControlSet001\Services\WindowsAzureGuestAgent"),
                ("Network Adapters", @"SYSTEM\ControlSet001\Services\Tcpip\Parameters\Interfaces"),
                ("Network Shares", @"SYSTEM\ControlSet001\Services\LanmanServer\Shares"),
                ("NTLM", @"SYSTEM\ControlSet001\Control\Lsa\MSV1_0"),
                ("RDP-Tcp", @"SYSTEM\ControlSet001\Control\Terminal Server\WinStations\RDP-Tcp"),
                ("Services", @"SYSTEM\ControlSet001\Services"),
                ("TLS/SSL", @"SYSTEM\ControlSet001\Control\SecurityProviders\SCHANNEL\Protocols"),
            },
        };

        // ── Nested types ────────────────────────────────────────────────────

        public class BookmarkItem
        {
            public required string Name { get; init; }
            public required string Path { get; init; }
            public required string HiveType { get; init; }
        }
    }
}

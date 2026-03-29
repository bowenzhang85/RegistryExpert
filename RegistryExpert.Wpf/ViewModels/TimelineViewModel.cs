using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using RegistryExpert.Core;
using RegistryParser.Abstractions;

namespace RegistryExpert.Wpf.ViewModels
{
    public class TimelineEntry : ViewModelBase
    {
        public DateTime LastModified { get; set; }
        public string KeyPath { get; set; } = "";
        public string DisplayPath { get; set; } = "";
        public TransactionLogDiff? TxLogDiff { get; set; }

        public string ChangeTypeText => TxLogDiff == null ? "" : TxLogDiff.ChangeType switch
        {
            TransactionLogChangeType.KeyAdded => "New",
            TransactionLogChangeType.KeyDeleted => "Deleted",
            TransactionLogChangeType.ValuesChanged => "Modified",
            _ => ""
        };

        public string ChangeTypeBrushKey => TxLogDiff == null ? "TextPrimaryBrush" : TxLogDiff.ChangeType switch
        {
            TransactionLogChangeType.KeyAdded => "DiffAddedBrush",
            TransactionLogChangeType.KeyDeleted => "DiffRemovedBrush",
            TransactionLogChangeType.ValuesChanged => "WarningBrush",
            _ => "TextPrimaryBrush"
        };

        public string Details => TxLogDiff?.ChangeSummary ?? "";
    }

    public class TxLogDetailItem
    {
        public string ValueName { get; set; } = "";
        public string ValueType { get; set; } = "";
        public string ChangeTypeText { get; set; } = "";
        public string ChangeTypeBrushKey { get; set; } = "TextPrimaryBrush";
        public string OldData { get; set; } = "";
        public string NewData { get; set; } = "";
    }

    public class TimelineViewModel : ViewModelBase
    {
        private readonly IReadOnlyList<LoadedHiveInfo> _loadedHives;
        private readonly Action<string>? _navigateToKey;
        private LoadedHiveInfo _activeHive;

        private List<TimelineEntry> _allEntries = new();
        private List<TimelineEntry> _filteredEntries = new();
        private List<TransactionLogDiff> _txLogDiffs = new();

        private CancellationTokenSource? _scanCts;
        private bool _isScanning;
        private string _statusText = "Click 'Scan' to analyze registry key timestamps";
        private bool _showProgress;

        private int _selectedTimeFilterIndex;
        private int _selectedLimitIndex = 1;
        private string _searchText = "";
        private bool _txLogFilterChecked;
        private bool _showTxLogFilter;
        private DateTime _customFromDate = DateTime.Now.AddDays(-30);
        private DateTime _customToDate = DateTime.Now;

        private int _currentDisplayCount;
        private int _pageSize = 100;

        private List<string> _manualLogPaths = new();
        private bool _hasTxLogColumns;
        private CancellationTokenSource? _analyzeLogsCts;

        private TimelineEntry? _selectedEntry;
        private bool _showDetailPanel;
        private string _detailHeaderText = "";

        private int _selectedHiveIndex;

        public ObservableCollection<TimelineEntry> DisplayedEntries { get; } = new();
        public ObservableCollection<TxLogDetailItem> DetailItems { get; } = new();
        public ObservableCollection<string> HiveNames { get; } = new();

        public ObservableCollection<string> TimeFilterOptions { get; } = new()
        {
            "All Times", "Last Hour", "Last 24 Hours", "Last 7 Days", "Last 30 Days", "Custom Range..."
        };

        public ObservableCollection<string> LimitOptions { get; } = new()
        {
            "50", "100", "500", "1000", "All"
        };

        public AsyncRelayCommand ScanCommand { get; }
        public RelayCommand ExportCommand { get; }
        public RelayCommand BrowseLogsCommand { get; }
        public RelayCommand LoadMoreCommand { get; }
        public RelayCommand NavigateCommand { get; }
        public RelayCommand CopyPathCommand { get; }

        public bool IsScanning
        {
            get => _isScanning;
            private set { _isScanning = value; OnPropertyChanged(); OnPropertyChanged(nameof(ScanButtonText)); }
        }

        public string ScanButtonText => _isScanning ? "Stop" : "Scan";

        public string StatusText
        {
            get => _statusText;
            set { _statusText = value; OnPropertyChanged(); }
        }

        public bool ShowProgress
        {
            get => _showProgress;
            set { _showProgress = value; OnPropertyChanged(); }
        }

        public int SelectedTimeFilterIndex
        {
            get => _selectedTimeFilterIndex;
            set
            {
                _selectedTimeFilterIndex = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ShowCustomDateRange));
                ApplyFilter();
            }
        }

        public bool ShowCustomDateRange => _selectedTimeFilterIndex == 5;

        public DateTime CustomFromDate
        {
            get => _customFromDate;
            set { _customFromDate = value; OnPropertyChanged(); if (ShowCustomDateRange) ApplyFilter(); }
        }

        public DateTime CustomToDate
        {
            get => _customToDate;
            set { _customToDate = value; OnPropertyChanged(); if (ShowCustomDateRange) ApplyFilter(); }
        }

        public int SelectedLimitIndex
        {
            get => _selectedLimitIndex;
            set
            {
                _selectedLimitIndex = value;
                OnPropertyChanged();
                _pageSize = value switch { 0 => 50, 1 => 100, 2 => 500, 3 => 1000, _ => int.MaxValue };
                ApplyFilter();
            }
        }

        public string SearchText
        {
            get => _searchText;
            set { _searchText = value; OnPropertyChanged(); ApplyFilter(); }
        }

        public bool TxLogFilterChecked
        {
            get => _txLogFilterChecked;
            set { _txLogFilterChecked = value; OnPropertyChanged(); ApplyFilter(); }
        }

        public bool ShowTxLogFilter
        {
            get => _showTxLogFilter;
            private set { _showTxLogFilter = value; OnPropertyChanged(); }
        }

        public bool HasTxLogColumns
        {
            get => _hasTxLogColumns;
            private set { _hasTxLogColumns = value; OnPropertyChanged(); }
        }

        public bool ShowLoadMore => _filteredEntries.Count > _currentDisplayCount;
        public string LoadMoreText => $"Load More ({_filteredEntries.Count - _currentDisplayCount:N0} remaining)";

        public TimelineEntry? SelectedEntry
        {
            get => _selectedEntry;
            set
            {
                _selectedEntry = value;
                OnPropertyChanged();
                UpdateDetailPanel();
            }
        }

        public bool ShowDetailPanel
        {
            get => _showDetailPanel;
            private set { _showDetailPanel = value; OnPropertyChanged(); }
        }

        public string DetailHeaderText
        {
            get => _detailHeaderText;
            private set { _detailHeaderText = value; OnPropertyChanged(); }
        }

        public bool ShowHiveSelector => _loadedHives.Count > 1;

        public int SelectedHiveIndex
        {
            get => _selectedHiveIndex;
            set
            {
                if (_selectedHiveIndex == value) return;
                _selectedHiveIndex = value;
                OnPropertyChanged();
                SwitchHive(value);
            }
        }

        public TimelineViewModel(IReadOnlyList<LoadedHiveInfo> loadedHives, Action<string>? navigateToKey = null)
        {
            _loadedHives = loadedHives;
            _navigateToKey = navigateToKey;
            _activeHive = loadedHives[0];

            foreach (var h in loadedHives)
                HiveNames.Add(h.HiveType.ToString());

            ScanCommand = new AsyncRelayCommand(OnScanAsync);
            ExportCommand = new RelayCommand(OnExport);
            BrowseLogsCommand = new RelayCommand(OnBrowseLogs);
            LoadMoreCommand = new RelayCommand(OnLoadMore);
            NavigateCommand = new RelayCommand(OnNavigate);
            CopyPathCommand = new RelayCommand(OnCopyPath);
        }

        public async Task AutoScanAsync()
        {
            await OnScanAsync(null);
        }

        private async Task OnScanAsync(object? _)
        {
            if (_isScanning)
            {
                _scanCts?.Cancel();
                return;
            }

            IsScanning = true;
            ShowProgress = true;
            _scanCts?.Dispose();
            _scanCts = new CancellationTokenSource();
            var token = _scanCts.Token;

            _allEntries.Clear();
            _filteredEntries.Clear();
            _txLogDiffs.Clear();
            DisplayedEntries.Clear();
            HasTxLogColumns = false;
            ShowTxLogFilter = false;
            ShowDetailPanel = false;

            try
            {
                var rootKey = _activeHive.Parser.GetRootKey();
                if (rootKey == null)
                {
                    StatusText = "Error: Could not read root key from hive";
                    return;
                }

                StatusText = "Scanning registry keys...";
                var scannedCount = 0;

                await Task.Run(() =>
                {
                    ScanKeyRecursive(rootKey, token, ref scannedCount);
                }, token).ConfigureAwait(true);

                _allEntries = _allEntries
                    .OrderByDescending(e => e.LastModified)
                    .ThenBy(e => e.DisplayPath, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                StatusText = $"Scanned {scannedCount:N0} keys. Checking for transaction logs...";

                var hivePath = _activeHive.Parser.FilePath;
                if (!string.IsNullOrEmpty(hivePath))
                {
                    var logPaths = _manualLogPaths.Count > 0
                        ? _manualLogPaths
                        : TransactionLogAnalyzer.DetectLogFiles(hivePath);

                    if (logPaths.Count > 0)
                    {
                        await AnalyzeTransactionLogsAsync(hivePath, logPaths, token);
                    }
                    else
                    {
                        StatusText = $"Scan complete: {scannedCount:N0} keys. No transaction logs found.";
                    }
                }
                else
                {
                    StatusText = $"Scan complete: {scannedCount:N0} keys.";
                }

                ApplyFilter();
            }
            catch (OperationCanceledException)
            {
                StatusText = "Scan cancelled.";
            }
            catch (Exception ex)
            {
                StatusText = $"Scan error: {ex.Message}";
            }
            finally
            {
                IsScanning = false;
                ShowProgress = false;
            }
        }

        private void ScanKeyRecursive(RegistryKey key, CancellationToken token, ref int count)
        {
            token.ThrowIfCancellationRequested();
            count++;

            if (key.LastWriteTime.HasValue)
            {
                _allEntries.Add(new TimelineEntry
                {
                    LastModified = key.LastWriteTime.Value.DateTime,
                    KeyPath = key.KeyPath,
                    DisplayPath = _activeHive.Parser.ConvertRootPath(key.KeyPath)
                });
            }

            if (key.SubKeys != null)
            {
                foreach (var sub in key.SubKeys)
                    ScanKeyRecursive(sub, token, ref count);
            }
        }

        private async Task AnalyzeTransactionLogsAsync(string hivePath, List<string> logPaths, CancellationToken token)
        {
            _analyzeLogsCts?.Cancel();
            _analyzeLogsCts?.Dispose();
            _analyzeLogsCts = new CancellationTokenSource();
            var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(token, _analyzeLogsCts.Token);

            try
            {
                _txLogDiffs = await TransactionLogAnalyzer.AnalyzeAsync(
                    hivePath,
                    logPaths,
                    msg => Application.Current.Dispatcher.Invoke(() => StatusText = msg),
                    linkedCts.Token
                ).ConfigureAwait(true);

                MergeTxLogDiffs();

                if (_txLogDiffs.Count > 0)
                {
                    HasTxLogColumns = true;
                    ShowTxLogFilter = true;
                }

                StatusText = $"Scan complete. Transaction log analysis: {_txLogDiffs.Count:N0} changes found.";
            }
            catch (OperationCanceledException)
            {
                StatusText = "Transaction log analysis cancelled.";
            }
            catch (Exception ex)
            {
                StatusText = $"Transaction log analysis error: {ex.Message}";
            }
        }

        private void MergeTxLogDiffs()
        {
            var entryByPath = new Dictionary<string, TimelineEntry>(StringComparer.OrdinalIgnoreCase);
            foreach (var entry in _allEntries)
                entryByPath.TryAdd(entry.KeyPath, entry);

            foreach (var diff in _txLogDiffs)
            {
                if (entryByPath.TryGetValue(diff.KeyPath, out var existing))
                {
                    existing.TxLogDiff = diff;
                }
                else
                {
                    _allEntries.Add(new TimelineEntry
                    {
                        LastModified = diff.NewTimestamp ?? diff.OldTimestamp ?? DateTime.MinValue,
                        KeyPath = diff.KeyPath,
                        DisplayPath = _activeHive.Parser.ConvertRootPath(diff.KeyPath),
                        TxLogDiff = diff
                    });
                }
            }

            _allEntries = _allEntries
                .OrderByDescending(e => e.LastModified)
                .ThenBy(e => e.DisplayPath, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private void ApplyFilter()
        {
            var now = DateTime.Now;
            DateTime? cutoff = _selectedTimeFilterIndex switch
            {
                1 => now.AddHours(-1),
                2 => now.AddHours(-24),
                3 => now.AddDays(-7),
                4 => now.AddDays(-30),
                _ => null
            };

            IEnumerable<TimelineEntry> filtered = _allEntries;

            if (_selectedTimeFilterIndex == 5)
            {
                filtered = filtered.Where(e => e.LastModified >= _customFromDate && e.LastModified <= _customToDate);
            }
            else if (cutoff.HasValue)
            {
                filtered = filtered.Where(e => e.LastModified >= cutoff.Value);
            }

            if (!string.IsNullOrWhiteSpace(_searchText))
            {
                filtered = filtered.Where(e =>
                    e.DisplayPath.Contains(_searchText, StringComparison.OrdinalIgnoreCase));
            }

            if (_txLogFilterChecked)
            {
                filtered = filtered.Where(e => e.TxLogDiff != null);
            }

            _filteredEntries = filtered.ToList();

            _currentDisplayCount = 0;
            DisplayedEntries.Clear();
            AppendPage();
        }

        private void AppendPage()
        {
            var toAdd = _filteredEntries
                .Skip(_currentDisplayCount)
                .Take(_pageSize)
                .ToList();

            foreach (var entry in toAdd)
                DisplayedEntries.Add(entry);

            _currentDisplayCount += toAdd.Count;
            OnPropertyChanged(nameof(ShowLoadMore));
            OnPropertyChanged(nameof(LoadMoreText));
        }

        private void UpdateDetailPanel()
        {
            DetailItems.Clear();

            if (_selectedEntry?.TxLogDiff == null || _txLogDiffs.Count == 0)
            {
                ShowDetailPanel = false;
                return;
            }

            var diff = _selectedEntry.TxLogDiff;

            var statusText = diff.ChangeType switch
            {
                TransactionLogChangeType.KeyAdded => "New",
                TransactionLogChangeType.KeyDeleted => "Deleted",
                TransactionLogChangeType.ValuesChanged => "Modified",
                _ => ""
            };
            var tsInfo = "";
            if (diff.OldTimestamp.HasValue && diff.NewTimestamp.HasValue)
                tsInfo = $"  |  {diff.OldTimestamp:yyyy-MM-dd HH:mm:ss}  -->  {diff.NewTimestamp:yyyy-MM-dd HH:mm:ss}";
            else if (diff.NewTimestamp.HasValue)
                tsInfo = $"  |  (new) {diff.NewTimestamp:yyyy-MM-dd HH:mm:ss}";

            DetailHeaderText = $"Transaction Log Details: {diff.DisplayPath}  [{statusText}]{tsInfo}";

            foreach (var vc in diff.ValueChanges)
            {
                var changeText = vc.ChangeType switch
                {
                    ValueChangeType.Added => "Added",
                    ValueChangeType.Removed => "Removed",
                    ValueChangeType.Modified => "Modified",
                    _ => ""
                };
                var brushKey = vc.ChangeType switch
                {
                    ValueChangeType.Added => "DiffAddedBrush",
                    ValueChangeType.Removed => "DiffRemovedBrush",
                    ValueChangeType.Modified => "WarningBrush",
                    _ => "TextPrimaryBrush"
                };
                DetailItems.Add(new TxLogDetailItem
                {
                    ValueName = vc.ValueName,
                    ValueType = FormatValueType(vc.ValueType),
                    ChangeTypeText = changeText,
                    ChangeTypeBrushKey = brushKey,
                    OldData = vc.OldData ?? "(not present)",
                    NewData = vc.NewData ?? "(not present)"
                });
            }

            if (diff.ValueChanges.Count == 0 && diff.OldTimestamp.HasValue && diff.NewTimestamp.HasValue)
            {
                DetailItems.Add(new TxLogDetailItem
                {
                    ValueName = "(key timestamp)",
                    ValueType = "",
                    ChangeTypeText = "Modified",
                    ChangeTypeBrushKey = "WarningBrush",
                    OldData = diff.OldTimestamp.Value.ToString("yyyy-MM-dd HH:mm:ss"),
                    NewData = diff.NewTimestamp.Value.ToString("yyyy-MM-dd HH:mm:ss")
                });
            }

            ShowDetailPanel = true;
        }

        private void OnLoadMore()
        {
            AppendPage();
        }

        private void OnNavigate()
        {
            if (_selectedEntry != null && _navigateToKey != null)
            {
                _navigateToKey(_selectedEntry.KeyPath);
                if (Application.Current.MainWindow != null)
                {
                    Application.Current.MainWindow.Activate();
                }
            }
        }

        private void OnCopyPath()
        {
            if (_selectedEntry != null)
            {
                try
                {
                    Clipboard.SetText(_selectedEntry.DisplayPath);
                    StatusText = "Path copied to clipboard";
                }
                catch (System.Runtime.InteropServices.ExternalException)
                {
                    StatusText = "Failed to copy — clipboard may be in use";
                }
            }
        }

        private void OnExport()
        {
            if (_filteredEntries.Count == 0)
            {
                MessageBox.Show("No data to export.", "Export",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Export Timeline to CSV",
                Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*",
                DefaultExt = ".csv",
                FileName = $"registry_timeline_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
            };

            if (dlg.ShowDialog() == true)
            {
                try
                {
                    using var writer = new System.IO.StreamWriter(dlg.FileName, false, System.Text.Encoding.UTF8, 65536);

                    if (_hasTxLogColumns)
                        writer.WriteLine("\"Last Modified\",\"Key Path\",\"Change Type\",\"Details\"");
                    else
                        writer.WriteLine("\"Last Modified\",\"Key Path\"");

                    foreach (var entry in _filteredEntries)
                    {
                        var date = EscapeCsv(entry.LastModified.ToString("yyyy-MM-dd HH:mm:ss"));
                        var path = EscapeCsv(entry.DisplayPath);

                        if (_hasTxLogColumns)
                        {
                            var change = EscapeCsv(entry.ChangeTypeText);
                            var details = EscapeCsv(entry.Details);
                            writer.WriteLine($"\"{date}\",\"{path}\",\"{change}\",\"{details}\"");
                        }
                        else
                        {
                            writer.WriteLine($"\"{date}\",\"{path}\"");
                        }
                    }

                    StatusText = $"Exported {_filteredEntries.Count:N0} entries to {System.IO.Path.GetFileName(dlg.FileName)}";
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Export failed: {ex.Message}", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    StatusText = "Export failed";
                }
            }
        }

        private async void OnBrowseLogs()
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Select Transaction Log Files",
                Filter = "Registry Log Files (*.LOG1;*.LOG2)|*.LOG1;*.LOG2|All Files (*.*)|*.*",
                Multiselect = true
            };

            var hivePath = _activeHive.Parser.FilePath;
            if (!string.IsNullOrEmpty(hivePath))
                dlg.InitialDirectory = System.IO.Path.GetDirectoryName(hivePath) ?? "";

            if (dlg.ShowDialog() != true) return;

            _manualLogPaths = new List<string>(dlg.FileNames);

            if (_allEntries.Count == 0)
            {
                StatusText = $"Selected {_manualLogPaths.Count} log file(s). Click 'Scan' to load timeline with log analysis.";
                return;
            }

            if (string.IsNullOrEmpty(hivePath)) return;

            ShowProgress = true;
            try
            {
                await AnalyzeTransactionLogsAsync(hivePath, _manualLogPaths, CancellationToken.None);
                ApplyFilter();
            }
            finally
            {
                ShowProgress = false;
            }
        }

        private void SwitchHive(int index)
        {
            if (index < 0 || index >= _loadedHives.Count) return;

            _scanCts?.Cancel();
            _analyzeLogsCts?.Cancel();
            _activeHive = _loadedHives[index];
            _allEntries.Clear();
            _filteredEntries.Clear();
            _txLogDiffs.Clear();
            _manualLogPaths.Clear();
            DisplayedEntries.Clear();
            DetailItems.Clear();
            HasTxLogColumns = false;
            ShowTxLogFilter = false;
            ShowDetailPanel = false;
            StatusText = "Hive changed. Click 'Scan' to analyze.";
        }

        private static string FormatValueType(string valueType)
        {
            if (string.IsNullOrEmpty(valueType)) return "";
            return int.TryParse(valueType, out _) ? $"Unknown ({valueType})" : valueType;
        }

        private static string EscapeCsv(string value) => value.Replace("\"", "\"\"");

        public void Cleanup()
        {
            _scanCts?.Cancel();
            _scanCts?.Dispose();
            _analyzeLogsCts?.Cancel();
            _analyzeLogsCts?.Dispose();
            _allEntries.Clear();
            _filteredEntries.Clear();
            _txLogDiffs.Clear();
        }
    }
}

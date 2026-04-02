using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Windows;
using System.Windows.Media;
using RegistryExpert.Core;

namespace RegistryExpert.Wpf.ViewModels
{
    /// <summary>
    /// ViewModel for the Search window. Owns all search state, pagination, and preview logic.
    /// </summary>
    public class SearchViewModel : ViewModelBase, IDisposable
    {
        private const int PageSize = 1000;

        // Known prefixes to strip from raw ValueData.ToString() output
        private static readonly string[] ValueDataPrefixes =
        {
            "RegSzData: ", "RegSzData:", "RegExpandSzData: ", "RegExpandSzData:",
            "RegDwordData: ", "RegDwordData:", "RegQwordData: ", "RegQwordData:",
            "RegBinaryData: ", "RegBinaryData:", "RegMultiSzData: ", "RegMultiSzData:",
            "Name: ", "Name:", "Data: ", "Data:"
        };

        // ── Backing fields ──────────────────────────────────────────────────

        private string _searchTerm = "";
        private bool _matchWholeWord;
        private bool _isSearching;
        private string _statusText = "Enter a search term and press Search";
        private Brush _statusBrush;
        private string _statusBrushKey = "TextSecondaryBrush";
        private SearchResult? _selectedResult;
        private string _previewPath = "Select a result to preview";
        private string _previewText = "";
        private bool _canLoadMore;
        private string _loadMoreText = "";

        // ── Internal state ──────────────────────────────────────────────────

        private readonly MainViewModel _mainViewModel;
        private readonly List<(SearchMatch Match, OfflineRegistryParser Parser)> _allMatches = new();
        private int _displayedCount;
        private CancellationTokenSource? _searchCts;
        private string _currentSearchTerm = "";

        /// <summary>The search term used for the most recent search (for highlight).</summary>
        public string CurrentSearchTerm => _currentSearchTerm;

        // ── Properties ──────────────────────────────────────────────────────

        public string SearchTerm
        {
            get => _searchTerm;
            set => SetProperty(ref _searchTerm, value);
        }

        public bool MatchWholeWord
        {
            get => _matchWholeWord;
            set => SetProperty(ref _matchWholeWord, value);
        }

        public bool IsSearching
        {
            get => _isSearching;
            set => SetProperty(ref _isSearching, value);
        }

        public string StatusText
        {
            get => _statusText;
            set => SetProperty(ref _statusText, value);
        }

        public Brush StatusBrush
        {
            get => _statusBrush;
            set => SetProperty(ref _statusBrush, value);
        }

        public ObservableCollection<SearchResult> Results { get; } = new();

        public SearchResult? SelectedResult
        {
            get => _selectedResult;
            set
            {
                if (SetProperty(ref _selectedResult, value))
                    OnSelectedResultChanged();
            }
        }

        public string PreviewPath
        {
            get => _previewPath;
            set => SetProperty(ref _previewPath, value);
        }

        public string PreviewText
        {
            get => _previewText;
            set => SetProperty(ref _previewText, value);
        }

        public bool CanLoadMore
        {
            get => _canLoadMore;
            set => SetProperty(ref _canLoadMore, value);
        }

        public string LoadMoreText
        {
            get => _loadMoreText;
            set => SetProperty(ref _loadMoreText, value);
        }

        // ── Commands ────────────────────────────────────────────────────────

        public AsyncRelayCommand SearchCommand { get; }
        public RelayCommand CancelCommand { get; }
        public RelayCommand LoadMoreCommand { get; }
        public RelayCommand NavigateCommand { get; }
        public RelayCommand CopyPathCommand { get; }
        public RelayCommand CopyValueCommand { get; }

        // ── Constructor ─────────────────────────────────────────────────────

        public SearchViewModel(MainViewModel mainViewModel)
        {
            _mainViewModel = mainViewModel;
            _statusBrush = GetBrush("TextSecondaryBrush");

            SearchCommand = new AsyncRelayCommand(OnSearchAsync);
            CancelCommand = new RelayCommand(OnCancel, () => IsSearching);
            LoadMoreCommand = new RelayCommand(OnLoadMore, () => CanLoadMore);
            NavigateCommand = new RelayCommand(OnNavigate, () => SelectedResult != null);
            CopyPathCommand = new RelayCommand(OnCopyPath, () => SelectedResult != null);
            CopyValueCommand = new RelayCommand(OnCopyValue, () => SelectedResult != null);
        }

        // ── Command handlers ────────────────────────────────────────────────

        private async Task OnSearchAsync()
        {
            var term = SearchTerm?.Trim() ?? "";
            if (string.IsNullOrEmpty(term))
            {
                MessageBox.Show("Please enter a search term.", "Search",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // Cancel any previous search
            _searchCts?.Cancel();
            _searchCts?.Dispose();
            _searchCts = new CancellationTokenSource();
            var token = _searchCts.Token;

            // Reset state
            _allMatches.Clear();
            Results.Clear();
            _displayedCount = 0;
            CanLoadMore = false;
            _currentSearchTerm = term;

            IsSearching = true;
            SetStatus("Searching...", "WarningBrush");

            try
            {
                foreach (var hiveInfo in _mainViewModel.LoadedHives.Values)
                {
                    token.ThrowIfCancellationRequested();

                    var parser = hiveInfo.Parser;
                    var matches = await Task.Run(
                        () => parser.SearchAll(term, caseSensitive: false,
                            wholeWord: MatchWholeWord, cancellationToken: token),
                        token).ConfigureAwait(true);

                    foreach (var match in matches)
                    {
                        _allMatches.Add((match, parser));
                    }
                }

                DisplayNextPage();
                UpdateStatusAfterSearch();
            }
            catch (OperationCanceledException)
            {
                SetStatus("Search cancelled", "WarningBrush");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Search error: {ex.Message}");
                SetStatus($"Search failed: {ex.Message}", "ErrorBrush");
            }
            finally
            {
                IsSearching = false;
            }
        }

        private void OnCancel()
        {
            _searchCts?.Cancel();
        }

        private void OnLoadMore()
        {
            DisplayNextPage();
            UpdateStatusAfterSearch();
        }

        private void OnNavigate()
        {
            var result = SelectedResult;
            if (result == null) return;

            // Navigate using display path (hive-prefixed)
            // NavigateToKey with optional valueName for auto-selecting values
            string? valueName = result.MatchType != "Key" ? result.ValueName : null;
            _mainViewModel.NavigateToKey(result.KeyPath, valueName);
        }

        private void OnCopyPath()
        {
            var result = SelectedResult;
            if (result == null) return;

            try
            {
                Clipboard.SetText(result.KeyPath);
            }
            catch (System.Runtime.InteropServices.ExternalException)
            {
                // Clipboard locked by another application
            }
        }

        private void OnCopyValue()
        {
            var result = SelectedResult;
            if (result == null) return;

            try
            {
                var text = !string.IsNullOrEmpty(result.FullValue) ? result.FullValue : result.Data;
                if (!string.IsNullOrEmpty(text))
                    Clipboard.SetText(text);
            }
            catch (System.Runtime.InteropServices.ExternalException)
            {
                // Clipboard locked by another application
            }
        }

        // ── Pagination ──────────────────────────────────────────────────────

        private void DisplayNextPage()
        {
            var remaining = _allMatches.Count - _displayedCount;
            var count = Math.Min(remaining, PageSize);
            if (count <= 0) return;

            for (int i = _displayedCount; i < _displayedCount + count; i++)
            {
                var (match, parser) = _allMatches[i];
                var displayPath = parser.ConvertRootPath(match.Key.KeyPath);

                SearchResult result;

                if (match.MatchKind == "Key" || match.MatchedValue == null)
                {
                    result = new SearchResult
                    {
                        KeyPath = displayPath,
                        RawKeyPath = match.Key.KeyPath,
                        MatchType = "Key",
                        Details = match.Key.KeyName ?? "",
                        ValueName = "",
                        Data = "",
                        FullValue = $"Key: {displayPath}",
                        ImageKey = "folder",
                        HiveTypeName = parser.CurrentHiveType.ToString()
                    };
                }
                else
                {
                    var value = match.MatchedValue;
                    var valueType = value.ValueType ?? "Unknown";
                    var valueName = value.ValueName ?? "";
                    var displayName = string.IsNullOrEmpty(valueName) ? "(Default)" : valueName;
                    var rawData = CleanValueData(value.ValueData?.ToString() ?? "");
                    var truncatedData = rawData.Length > 200 ? rawData[..200] + "..." : rawData;

                    result = new SearchResult
                    {
                        KeyPath = displayPath,
                        RawKeyPath = match.Key.KeyPath,
                        MatchType = valueType,
                        Details = displayName,
                        ValueName = valueName,
                        Data = truncatedData,
                        FullValue = $"Name: {displayName}\nType: {valueType}\nData: {rawData}",
                        ImageKey = RegistryValueItem.GetImageKey(valueType),
                        HiveTypeName = parser.CurrentHiveType.ToString()
                    };
                }

                Results.Add(result);
            }

            _displayedCount += count;
            UpdateLoadMoreState();
        }

        private void UpdateLoadMoreState()
        {
            var remaining = _allMatches.Count - _displayedCount;
            CanLoadMore = remaining > 0;
            if (remaining > 0)
            {
                var nextBatch = Math.Min(remaining, PageSize);
                LoadMoreText = $"Load next {nextBatch}";
            }
            else
            {
                LoadMoreText = "";
            }
        }

        // ── Status ──────────────────────────────────────────────────────────

        private void UpdateStatusAfterSearch()
        {
            var total = _allMatches.Count;
            var displayed = _displayedCount;

            if (total == 0)
            {
                SetStatus("No results found", "TextSecondaryBrush");
            }
            else if (displayed < total)
            {
                SetStatus($"Showing {displayed:N0} of {total:N0} results", "SuccessBrush");
            }
            else if (total > PageSize)
            {
                SetStatus($"Showing all {displayed:N0} results", "SuccessBrush");
            }
            else
            {
                SetStatus($"Found {displayed:N0} results", "SuccessBrush");
            }
        }

        private void SetStatus(string text, string brushKey)
        {
            StatusText = text;
            _statusBrushKey = brushKey;
            StatusBrush = GetBrush(brushKey);
        }

        /// <summary>Refreshes the status brush from the current theme resources (call after theme change).</summary>
        public void RefreshStatusBrush()
        {
            StatusBrush = GetBrush(_statusBrushKey);
        }

        // ── Preview ─────────────────────────────────────────────────────────

        private void OnSelectedResultChanged()
        {
            var result = SelectedResult;
            if (result == null)
            {
                PreviewPath = "Select a result to preview";
                PreviewText = "";
                return;
            }

            PreviewPath = result.KeyPath;
            PreviewText = result.FullValue;
        }

        // ── Helpers ─────────────────────────────────────────────────────────

        private static string CleanValueData(string valueData)
        {
            if (string.IsNullOrEmpty(valueData))
                return valueData;

            foreach (var prefix in ValueDataPrefixes)
            {
                if (valueData.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    return valueData[prefix.Length..].Trim();
                }
            }

            return valueData.Trim();
        }

        private static Brush GetBrush(string resourceKey)
        {
            if (Application.Current.Resources[resourceKey] is Brush brush)
                return brush;
            return Brushes.Gray;
        }

        // ── Disposal ────────────────────────────────────────────────────────

        public void Dispose()
        {
            _searchCts?.Cancel();
            _searchCts?.Dispose();
            _allMatches.Clear();
            Results.Clear();
        }
    }
}

using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Windows.Input;
using System.Windows.Media;
using RegistryExpert.Core;
using HiveType = RegistryExpert.Core.OfflineRegistryParser.HiveType;

namespace RegistryExpert.Wpf.ViewModels
{
    /// <summary>Distinguishes which statistic a tree node represents.</summary>
    public enum StatMode { KeyCount, ValueCount, DataSize }

    public class StatisticsViewModel : ViewModelBase
    {
        // ── Inner types ─────────────────────────────────────────────────────

        public class StatCardItem
        {
            public string Label { get; init; } = "";
            public string Value { get; init; } = "";
            public string BrushKey { get; init; } = "AccentBrush";
        }

        public class KeyStatNode : ViewModelBase
        {
            private bool _isExpanded;
            private bool _isLoading;
            private ObservableCollection<KeyStatNode> _children = new();

            public string DisplayName { get; init; } = "";
            public string FullPath { get; init; } = "";
            public long StatValue { get; init; }
            public double BarWidth { get; init; }
            public string FormattedValue { get; init; } = "";
            public string BarBrushKey { get; init; } = "AccentBrush";
            public bool HasChildren { get; init; }
            public bool IsMoreNode { get; init; }
            public bool IsValueNode { get; init; }

            /// <summary>Nesting depth in the tree (0 = root). Used to shrink the name column so bar charts stay left-aligned.</summary>
            public int Depth { get; init; }

            /// <summary>Name column width that compensates for TreeView indentation (~19px per level).</summary>
            public double NameColumnWidth => Math.Max(120, 280 - Depth * 19);

            /// <summary>Registry value type name for tooltip display (e.g., "RegSz").</summary>
            public string? ValueTypeName { get; init; }

            /// <summary>Remaining items for "... and N more" pagination nodes.</summary>
            public List<KeyStatistics>? RemainingItems { get; init; }

            /// <summary>Remaining value items for "... and N more values" pagination nodes.</summary>
            public List<ValueStatistic>? RemainingValueItems { get; init; }

            /// <summary>Reference to the parent ViewModel for lazy loading.</summary>
            public StatisticsViewModel? OwnerVm { get; init; }

            /// <summary>Which statistic this node tracks (key count, value count, or data size).</summary>
            public StatMode StatMode { get; init; }

            /// <summary>Local max for value bar scaling (carried on "more values" pagination nodes).</summary>
            internal long _valueLocalMax;

            public ObservableCollection<KeyStatNode> Children
            {
                get => _children;
                set => SetProperty(ref _children, value);
            }

            public bool IsExpanded
            {
                get => _isExpanded;
                set
                {
                    if (SetProperty(ref _isExpanded, value) && value)
                        OnExpanded();
                }
            }

            public bool IsLoading
            {
                get => _isLoading;
                set => SetProperty(ref _isLoading, value);
            }

            private async void OnExpanded()
            {
                // "More items" nodes are handled by StatisticsWindow code-behind
                // via the TreeViewItem.Expanded routed event, not here.
                if (IsMoreNode)
                    return;

                // Lazy load children: check for dummy placeholder
                if (Children.Count == 1 && Children[0].DisplayName == "Loading...")
                {
                    if (OwnerVm != null)
                    {
                        IsLoading = true;
                        try
                        {
                            await OwnerVm.LoadChildStatsAsync(this);
                        }
                        finally
                        {
                            IsLoading = false;
                        }
                    }
                }
            }
        }

        // ── Constants ───────────────────────────────────────────────────────

        private const double MaxBarWidth = 300.0;
        private const int InitialBatchSize = 50;
        private const int MoreBatchSize = 100;

        // ── Backing fields ──────────────────────────────────────────────────

        private bool _isAnalyzing;
        private string _windowTitle = "Registry Statistics";
        private string? _selectedHiveName;
        private bool _showHiveSelector;
        private int _selectedTabIndex;

        // ── State ───────────────────────────────────────────────────────────

        private readonly IReadOnlyList<LoadedHiveInfo> _loadedHives;
        private OfflineRegistryParser _activeParser = null!;
        private long _keyCountRootMax = 1;
        private long _valueCountRootMax = 1;
        private long _dataSizeRootMax = 1;

        // ── Properties ──────────────────────────────────────────────────────

        public string WindowTitle
        {
            get => _windowTitle;
            set => SetProperty(ref _windowTitle, value);
        }

        public bool IsAnalyzing
        {
            get => _isAnalyzing;
            set => SetProperty(ref _isAnalyzing, value);
        }

        public bool ShowHiveSelector
        {
            get => _showHiveSelector;
            set => SetProperty(ref _showHiveSelector, value);
        }

        public ObservableCollection<string> HiveNames { get; } = new();

        public string? SelectedHiveName
        {
            get => _selectedHiveName;
            set
            {
                if (SetProperty(ref _selectedHiveName, value) && value != null)
                    OnHiveChanged(value);
            }
        }

        public int SelectedTabIndex
        {
            get => _selectedTabIndex;
            set => SetProperty(ref _selectedTabIndex, value);
        }

        public ObservableCollection<StatCardItem> StatCards { get; } = new();
        public ObservableCollection<KeyStatNode> KeyCountNodes { get; } = new();
        public ObservableCollection<KeyStatNode> ValueCountNodes { get; } = new();
        public ObservableCollection<KeyStatNode> DataSizeNodes { get; } = new();

        // ── Constructor ─────────────────────────────────────────────────────

        public StatisticsViewModel(IReadOnlyList<LoadedHiveInfo> loadedHives)
        {
            _loadedHives = loadedHives;

            if (loadedHives.Count > 1)
            {
                ShowHiveSelector = true;
                foreach (var hive in loadedHives)
                    HiveNames.Add(hive.HiveType.ToString());
            }

            // Use the first hive as default
            if (loadedHives.Count > 0)
            {
                _activeParser = loadedHives[0].Parser;
                _selectedHiveName = loadedHives[0].HiveType.ToString();
                WindowTitle = $"Registry Statistics - {_selectedHiveName}";
                LoadStatsForParser(_activeParser);
            }
        }

        // ── Hive switching ──────────────────────────────────────────────────

        private void OnHiveChanged(string hiveName)
        {
            if (!Enum.TryParse<HiveType>(hiveName, out var ht))
                return;

            var hive = _loadedHives.FirstOrDefault(h => h.HiveType == ht);
            if (hive == null) return;

            _activeParser = hive.Parser;
            WindowTitle = $"Registry Statistics - {ht}";
            LoadStatsForParser(_activeParser);
        }

        // ── Stats loading ───────────────────────────────────────────────────

        private async void LoadStatsForParser(OfflineRegistryParser parser)
        {
            IsAnalyzing = true;

            // Update summary cards immediately (fast operation)
            var stats = parser.GetStatistics();
            StatCards.Clear();
            StatCards.Add(new StatCardItem { Label = "File Size", Value = stats.FormattedFileSize, BrushKey = "AccentBrush" });
            StatCards.Add(new StatCardItem { Label = "Total Keys", Value = stats.TotalKeys.ToString("N0"), BrushKey = "SuccessBrush" });
            StatCards.Add(new StatCardItem { Label = "Total Values", Value = stats.TotalValues.ToString("N0"), BrushKey = "WarningBrush" });
            StatCards.Add(new StatCardItem { Label = "Hive Type", Value = stats.HiveType, BrushKey = "InfoBrush" });

            // Clear tree data
            KeyCountNodes.Clear();
            ValueCountNodes.Clear();
            DataSizeNodes.Clear();

            // Analyze in background
            List<KeyStatistics> keyStats;
            try
            {
                keyStats = await Task.Run(() => RegistryStatisticsAnalyzer.AnalyzeTopLevelKeys(parser));
            }
            catch (ObjectDisposedException)
            {
                IsAnalyzing = false;
                return;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"LoadStatsForParser error: {ex.Message}");
                IsAnalyzing = false;
                return;
            }

            // Populate key count tree
            var byCount = keyStats.OrderByDescending(k => k.SubKeyCount).ToList();
            PopulateTree(KeyCountNodes, byCount, k => k.SubKeyCount, "SuccessBrush", StatMode.KeyCount);

            // Populate value count tree
            var byValues = keyStats.OrderByDescending(k => k.ValueCount).ToList();
            PopulateTree(ValueCountNodes, byValues, k => k.ValueCount, "WarningBrush", StatMode.ValueCount);

            // Populate data size tree
            var bySize = keyStats.OrderByDescending(k => k.TotalSize).ToList();
            PopulateTree(DataSizeNodes, bySize, k => k.TotalSize, "AccentBrush", StatMode.DataSize);

            IsAnalyzing = false;
        }

        // ── Tree population ─────────────────────────────────────────────────

        private void PopulateTree(ObservableCollection<KeyStatNode> nodes, List<KeyStatistics> data,
            Func<KeyStatistics, long> valueSelector, string brushKey, StatMode statMode)
        {
            nodes.Clear();
            if (data.Count == 0) return;

            long maxValue = data.Max(d => valueSelector(d));
            if (maxValue == 0) maxValue = 1;

            // Store root-level max so child/more-node loading uses the same scale
            switch (statMode)
            {
                case StatMode.KeyCount: _keyCountRootMax = maxValue; break;
                case StatMode.ValueCount: _valueCountRootMax = maxValue; break;
                case StatMode.DataSize: _dataSizeRootMax = maxValue; break;
            }

            var batch = data.Take(InitialBatchSize).ToList();
            foreach (var item in batch)
            {
                nodes.Add(CreateNode(item, valueSelector(item), maxValue, brushKey, statMode));
            }

            // Add "more" node if needed
            if (data.Count > InitialBatchSize)
            {
                var remaining = data.Skip(InitialBatchSize).ToList();
                nodes.Add(CreateMoreNode(remaining, brushKey, statMode));
            }
        }

        private KeyStatNode CreateNode(KeyStatistics item, long value, long maxValue, string brushKey, StatMode statMode, int depth = 0)
        {
            var displayName = item.KeyPath.Contains('\\')
                ? item.KeyPath.Substring(item.KeyPath.LastIndexOf('\\') + 1)
                : item.KeyPath;

            var barWidth = (double)value / maxValue * MaxBarWidth;
            if (barWidth < 3) barWidth = 3;

            var formattedValue = statMode == StatMode.DataSize
                ? RegistryStatisticsAnalyzer.FormatSize(value)
                : value.ToString("N0");

            var node = new KeyStatNode
            {
                DisplayName = displayName,
                FullPath = item.KeyPath,
                StatValue = value,
                BarWidth = barWidth,
                FormattedValue = formattedValue,
                BarBrushKey = brushKey,
                HasChildren = item.SubKeyCount > 1
                    || (statMode == StatMode.ValueCount && item.ValueCount > 0),
                StatMode = statMode,
                OwnerVm = this,
                Depth = depth
            };

            // Add dummy child for expand button if key has subkeys,
            // or in ValueCount mode if it has direct values to show
            if (item.SubKeyCount > 1
                || (statMode == StatMode.ValueCount && item.ValueCount > 0))
            {
                node.Children.Add(new KeyStatNode { DisplayName = "Loading..." });
            }

            return node;
        }

        private KeyStatNode CreateMoreNode(List<KeyStatistics> remaining, string brushKey, StatMode statMode, int depth = 0)
        {
            var moreNode = new KeyStatNode
            {
                DisplayName = $"... and {remaining.Count} more",
                FullPath = "",
                StatValue = 0,
                BarWidth = 0,
                FormattedValue = "",
                BarBrushKey = brushKey,
                IsMoreNode = true,
                RemainingItems = remaining,
                StatMode = statMode,
                OwnerVm = this,
                Depth = depth
            };
            // Dummy child so it's expandable
            moreNode.Children.Add(new KeyStatNode { DisplayName = "Loading..." });
            return moreNode;
        }

        // ── Lazy child loading ──────────────────────────────────────────────

        internal async Task LoadChildStatsAsync(KeyStatNode parentNode)
        {
            if (string.IsNullOrEmpty(parentNode.FullPath)) return;

            try
            {
                // Expensive registry I/O on background thread
                var childStats = await Task.Run(() =>
                    RegistryStatisticsAnalyzer.GetChildKeyStats(
                        _activeParser, parentNode.FullPath, parentNode.StatMode == StatMode.DataSize));

                List<ValueStatistic>? valueStats = null;
                if (parentNode.StatMode == StatMode.ValueCount)
                {
                    valueStats = await Task.Run(() =>
                        RegistryStatisticsAnalyzer.GetKeyValueStats(_activeParser, parentNode.FullPath));
                }

                // UI thread resumes here (WPF SynchronizationContext)
                parentNode.Children.Clear();

                Func<KeyStatistics, long> valueSelector = parentNode.StatMode switch
                {
                    StatMode.DataSize => k => k.TotalSize,
                    StatMode.ValueCount => k => k.ValueCount,
                    _ => k => (long)k.SubKeyCount
                };

                long maxValue = parentNode.StatMode switch
                {
                    StatMode.DataSize => _dataSizeRootMax,
                    StatMode.ValueCount => _valueCountRootMax,
                    _ => _keyCountRootMax
                };
                if (maxValue == 0) maxValue = 1;

                if (childStats.Count > 0)
                {
                    var sorted = childStats.OrderByDescending(c => valueSelector(c)).ToList();
                    int childDepth = parentNode.Depth + 1;

                    foreach (var child in sorted.Take(InitialBatchSize))
                    {
                        parentNode.Children.Add(CreateNode(child, valueSelector(child), maxValue,
                            parentNode.BarBrushKey, parentNode.StatMode, childDepth));
                    }

                    if (sorted.Count > InitialBatchSize)
                    {
                        var remaining = sorted.Skip(InitialBatchSize).ToList();
                        parentNode.Children.Add(CreateMoreNode(remaining, parentNode.BarBrushKey, parentNode.StatMode, childDepth));
                    }
                }

                // For ValueCount mode, list the key's own direct data values
                if (valueStats != null && valueStats.Count > 0)
                {
                    var sorted = valueStats.OrderBy(v => v.ValueName).ToList();
                    int childDepth = parentNode.Depth + 1;

                    parentNode.Children.Add(new KeyStatNode
                    {
                        DisplayName = $"Values ({sorted.Count:N0})",
                        FormattedValue = "",
                        BarBrushKey = "InfoBrush",
                        IsValueNode = true,
                        Depth = childDepth
                    });

                    foreach (var val in sorted.Take(InitialBatchSize))
                    {
                        parentNode.Children.Add(CreateValueNode(val, childDepth));
                    }

                    if (sorted.Count > InitialBatchSize)
                    {
                        var remaining = sorted.Skip(InitialBatchSize).ToList();
                        parentNode.Children.Add(CreateMoreValueNode(remaining, childDepth));
                    }
                }

                if (parentNode.Children.Count == 0)
                {
                    parentNode.Children.Add(new KeyStatNode
                    {
                        DisplayName = "(No subkeys or values)",
                        FormattedValue = "",
                        Depth = parentNode.Depth + 1
                    });
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading child stats: {ex.Message}");
                parentNode.Children.Clear();
                parentNode.Children.Add(new KeyStatNode
                {
                    DisplayName = "(Error loading)",
                    FormattedValue = "",
                    Depth = parentNode.Depth + 1
                });
            }
        }

        /// <summary>
        /// Expand a "more items" node by inserting its batch into the parent collection.
        /// Called from code-behind since we need to manipulate the parent's collection.
        /// </summary>
        internal void ExpandMoreNode(KeyStatNode moreNode, ObservableCollection<KeyStatNode> parentCollection)
        {
            if (moreNode.RemainingItems == null) return;

            var remaining = moreNode.RemainingItems;
            Func<KeyStatistics, long> valueSelector = moreNode.StatMode switch
            {
                StatMode.DataSize => k => k.TotalSize,
                StatMode.ValueCount => k => k.ValueCount,
                _ => k => (long)k.SubKeyCount
            };

            long maxValue = moreNode.StatMode switch
            {
                StatMode.DataSize => _dataSizeRootMax,
                StatMode.ValueCount => _valueCountRootMax,
                _ => _keyCountRootMax
            };
            if (maxValue == 0) maxValue = 1;

            var batch = remaining.Take(MoreBatchSize).ToList();
            var leftover = remaining.Skip(MoreBatchSize).ToList();

            int moreIndex = parentCollection.IndexOf(moreNode);
            if (moreIndex < 0) return;

            // Remove the "more" node
            parentCollection.RemoveAt(moreIndex);

            // Insert batch items at the same position
            int insertIndex = moreIndex;
            foreach (var child in batch)
            {
                parentCollection.Insert(insertIndex++, CreateNode(child, valueSelector(child), maxValue,
                    moreNode.BarBrushKey, moreNode.StatMode, moreNode.Depth));
            }

            // Add another "more" node if there are still remaining items
            if (leftover.Count > 0)
            {
                parentCollection.Insert(insertIndex, CreateMoreNode(leftover, moreNode.BarBrushKey, moreNode.StatMode, moreNode.Depth));
            }
        }

        // ── Value node helpers (Value Counts tab) ───────────────────────────

        private KeyStatNode CreateValueNode(ValueStatistic val, int depth = 0)
        {
            return new KeyStatNode
            {
                DisplayName = string.IsNullOrEmpty(val.ValueName) ? "(Default)" : val.ValueName,
                FullPath = "",
                StatValue = 1,
                BarWidth = 3,
                FormattedValue = "1",
                BarBrushKey = "InfoBrush",
                HasChildren = false,
                IsValueNode = true,
                ValueTypeName = val.ValueType,
                StatMode = StatMode.ValueCount,
                Depth = depth
            };
        }

        private KeyStatNode CreateMoreValueNode(List<ValueStatistic> remaining, int depth = 0)
        {
            var moreNode = new KeyStatNode
            {
                DisplayName = $"... and {remaining.Count:N0} more values",
                FullPath = "",
                StatValue = 0,
                BarWidth = 0,
                FormattedValue = "",
                BarBrushKey = "InfoBrush",
                IsMoreNode = true,
                IsValueNode = true,
                RemainingValueItems = remaining,
                StatMode = StatMode.ValueCount,
                OwnerVm = this,
                Depth = depth
            };
            // Dummy child so it's expandable
            moreNode.Children.Add(new KeyStatNode { DisplayName = "Loading..." });
            return moreNode;
        }

        /// <summary>
        /// Expand a "more values" pagination node by inserting value items into the parent collection.
        /// Called from code-behind.
        /// </summary>
        internal void ExpandMoreValueNode(KeyStatNode moreNode, ObservableCollection<KeyStatNode> parentCollection)
        {
            if (moreNode.RemainingValueItems == null) return;

            var remaining = moreNode.RemainingValueItems;

            var batch = remaining.Take(MoreBatchSize).ToList();
            var leftover = remaining.Skip(MoreBatchSize).ToList();

            int moreIndex = parentCollection.IndexOf(moreNode);
            if (moreIndex < 0) return;

            // Remove the "more" node
            parentCollection.RemoveAt(moreIndex);

            // Insert batch items at the same position
            int insertIndex = moreIndex;
            foreach (var val in batch)
            {
                parentCollection.Insert(insertIndex++, CreateValueNode(val, moreNode.Depth));
            }

            // Add another "more values" node if there are still remaining items
            if (leftover.Count > 0)
            {
                parentCollection.Insert(insertIndex, CreateMoreValueNode(leftover, moreNode.Depth));
            }
        }
    }
}

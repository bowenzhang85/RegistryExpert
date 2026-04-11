using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using RegistryParser.Abstractions;
using RegistryExpert.Core;
using RegistryExpert.Core.Services;

namespace RegistryExpert.Wpf.ViewModels
{
    /// <summary>
    /// Tree node for the comparison TreeView. Holds diff info for coloring.
    /// Lazy loading: unexpanded nodes have a single dummy child ("Loading...").
    /// </summary>
    public class CompareTreeNode : ViewModelBase
    {
        private bool _isExpanded;
        private bool _isSelected;
        private ObservableCollection<CompareTreeNode> _children = new();
        private string _displayName = "";

        public string DisplayName { get => _displayName; set => SetProperty(ref _displayName, value); }
        public string Path { get; init; } = "";
        public RegistryKey Key { get; init; } = null!;
        public bool HasDifference { get; init; }
        public bool IsUniqueToThisHive { get; init; }
        public bool HasValueDifference { get; init; }
        public bool IsDummy { get; init; }

        /// <summary>
        /// DiffColorKey: "added" for GREEN (unique), "removed" for RED (value diff), "" for no diff.
        /// Used by XAML DataTriggers to set Foreground color.
        /// </summary>
        public string DiffColorKey
        {
            get
            {
                if (!HasDifference) return "";
                return IsUniqueToThisHive ? "added" : "removed";
            }
        }

        public ObservableCollection<CompareTreeNode> Children
        {
            get => _children;
            set => SetProperty(ref _children, value);
        }

        public bool IsExpanded
        {
            get => _isExpanded;
            set => SetProperty(ref _isExpanded, value);
        }

        public bool IsSelected
        {
            get => _isSelected;
            set => SetProperty(ref _isSelected, value);
        }
    }

    /// <summary>
    /// Row item for the values DataGrid with diff coloring.
    /// </summary>
    public class CompareValueItem : ViewModelBase
    {
        public string Name { get; init; } = "";
        public string Type { get; init; } = "";
        public string Data { get; init; } = "";
        public string ImageKey { get; init; } = "reg_str";
        public ValueDiffStatus DiffStatus { get; init; }

        public string DiffColorKey => DiffStatus switch
        {
            ValueDiffStatus.UniqueToThisHive => "added",
            ValueDiffStatus.ValueDiffers => "removed",
            _ => ""
        };
    }

    /// <summary>
    /// Represents a hive that's already loaded in the main window (for the dropdown).
    /// </summary>
    public class AvailableHiveItem
    {
        public string FilePath { get; init; } = "";
        public string DisplayName { get; init; } = "";
    }

    public class CompareViewModel : ViewModelBase
    {
        // ── State ──
        private bool _isLandingView = true;
        private bool _isComparing;
        private string _compareProgressText = "";

        // Left side
        private OfflineRegistryParser? _leftParser;
        private string _leftFilePath = "";
        private string _leftFileName = "";
        private string _leftStatusText = "";
        private bool _leftIsLoaded;
        private bool _leftIsLoading;
        private double _leftLoadProgress;
        private CompareTreeNode? _selectedLeftNode;
        private string _leftPathText = "";

        // Right side
        private OfflineRegistryParser? _rightParser;
        private string _rightFilePath = "";
        private string _rightFileName = "";
        private string _rightStatusText = "Load the first hive first";
        private bool _rightIsLoaded;
        private bool _rightIsLoading;
        private double _rightLoadProgress;
        private CompareTreeNode? _selectedRightNode;
        private string _rightPathText = "";

        // Comparison data (non-null after comparison)
        private Dictionary<string, RegistryKey>? _leftKeyIndex;
        private Dictionary<string, RegistryKey>? _rightKeyIndex;
        private Dictionary<string, RegistryComparer.DiffInfo>? _leftDiffIndex;
        private Dictionary<string, RegistryComparer.DiffInfo>? _rightDiffIndex;

        // Filter
        private bool _showDifferencesOnly;

        // Cancellation
        private CancellationTokenSource? _leftLoadCts;
        private CancellationTokenSource? _rightLoadCts;
        private CancellationTokenSource? _compareCts;

        // Syncing flag (prevents infinite loops during synchronized navigation)
        private bool _isSyncing;

        // ── Properties ──
        public bool IsLandingView { get => _isLandingView; set { if (SetProperty(ref _isLandingView, value)) OnPropertyChanged(nameof(IsComparisonView)); } }
        public bool IsComparisonView => !_isLandingView;
        public bool IsComparing { get => _isComparing; set { if (SetProperty(ref _isComparing, value)) OnPropertyChanged(nameof(CanCompare)); } }
        public string CompareProgressText { get => _compareProgressText; set => SetProperty(ref _compareProgressText, value); }

        // Left
        public string LeftFileName { get => _leftFileName; set => SetProperty(ref _leftFileName, value); }
        public string LeftStatusText { get => _leftStatusText; set => SetProperty(ref _leftStatusText, value); }
        public bool LeftIsLoaded { get => _leftIsLoaded; set { if (SetProperty(ref _leftIsLoaded, value)) { OnPropertyChanged(nameof(CanCompare)); OnPropertyChanged(nameof(RightCanLoad)); } } }
        public bool LeftIsLoading { get => _leftIsLoading; set { if (SetProperty(ref _leftIsLoading, value)) OnPropertyChanged(nameof(CanCompare)); } }
        public double LeftLoadProgress { get => _leftLoadProgress; set => SetProperty(ref _leftLoadProgress, value); }
        public string LeftPathText { get => _leftPathText; set => SetProperty(ref _leftPathText, value); }
        public CompareTreeNode? SelectedLeftNode
        {
            get => _selectedLeftNode;
            set
            {
                if (SetProperty(ref _selectedLeftNode, value) && value != null)
                    OnLeftNodeSelected(value);
            }
        }
        public ObservableCollection<CompareTreeNode> LeftTreeRoots { get; } = new();
        public ObservableCollection<CompareValueItem> LeftValues { get; } = new();

        // Right
        public string RightFileName { get => _rightFileName; set => SetProperty(ref _rightFileName, value); }
        public string RightStatusText { get => _rightStatusText; set => SetProperty(ref _rightStatusText, value); }
        public bool RightIsLoaded { get => _rightIsLoaded; set { if (SetProperty(ref _rightIsLoaded, value)) OnPropertyChanged(nameof(CanCompare)); } }
        public bool RightIsLoading { get => _rightIsLoading; set { if (SetProperty(ref _rightIsLoading, value)) OnPropertyChanged(nameof(CanCompare)); } }
        public double RightLoadProgress { get => _rightLoadProgress; set => SetProperty(ref _rightLoadProgress, value); }
        public string RightPathText { get => _rightPathText; set => SetProperty(ref _rightPathText, value); }
        public CompareTreeNode? SelectedRightNode
        {
            get => _selectedRightNode;
            set
            {
                if (SetProperty(ref _selectedRightNode, value) && value != null)
                    OnRightNodeSelected(value);
            }
        }
        public ObservableCollection<CompareTreeNode> RightTreeRoots { get; } = new();
        public ObservableCollection<CompareValueItem> RightValues { get; } = new();

        // Derived
        public bool CanCompare => LeftIsLoaded && RightIsLoaded && !IsComparing && !LeftIsLoading && !RightIsLoading;
        public bool RightCanLoad => LeftIsLoaded;

        // Filter
        public bool ShowDifferencesOnly
        {
            get => _showDifferencesOnly;
            set => SetProperty(ref _showDifferencesOnly, value);
        }

        // Available hives from main window
        public ObservableCollection<AvailableHiveItem> AvailableHives { get; } = new();
        public bool HasAvailableHives => AvailableHives.Count > 0;

        // ── Commands ──
        public AsyncRelayCommand LoadLeftHiveCommand { get; }
        public AsyncRelayCommand LoadRightHiveCommand { get; }
        public RelayCommand LoadLeftFromAvailableCommand { get; }
        public RelayCommand LoadRightFromAvailableCommand { get; }
        public AsyncRelayCommand CompareCommand { get; }
        public RelayCommand BackToLandingCommand { get; }
        public RelayCommand CancelLoadCommand { get; }

        // ── Constructor ──
        public CompareViewModel(IReadOnlyList<LoadedHiveInfo> loadedHives)
        {
            // Populate available hives from main window
            foreach (var hive in loadedHives)
            {
                AvailableHives.Add(new AvailableHiveItem
                {
                    FilePath = hive.FilePath,
                    DisplayName = $"{hive.Parser.FriendlyName}  ({Path.GetFileName(hive.FilePath)})"
                });
            }
            OnPropertyChanged(nameof(HasAvailableHives));

            // Auto-load if exactly one hive is available
            if (loadedHives.Count == 1)
                _ = LoadHiveFileAsync(loadedHives[0].FilePath, isLeft: true);

            LoadLeftHiveCommand = new AsyncRelayCommand(
                async () => await BrowseAndLoadHiveAsync(isLeft: true));

            LoadRightHiveCommand = new AsyncRelayCommand(
                async () => await BrowseAndLoadHiveAsync(isLeft: false));

            LoadLeftFromAvailableCommand = new RelayCommand(
                (param) =>
                {
                    if (param is AvailableHiveItem item)
                        _ = LoadHiveFileAsync(item.FilePath, isLeft: true);
                });

            LoadRightFromAvailableCommand = new RelayCommand(
                (param) =>
                {
                    if (param is AvailableHiveItem item)
                        _ = LoadHiveFileAsync(item.FilePath, isLeft: false);
                },
                (_) => RightCanLoad);

            CompareCommand = new AsyncRelayCommand(RunComparisonAsync, () => CanCompare);

            BackToLandingCommand = new RelayCommand(ResetToLanding);

            CancelLoadCommand = new RelayCommand(() =>
            {
                _leftLoadCts?.Cancel();
                _rightLoadCts?.Cancel();
            });
        }

        // ── Hive Loading ──

        private async Task BrowseAndLoadHiveAsync(bool isLeft)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = isLeft ? "Open First Registry Hive" : "Open Second Registry Hive",
                Filter = "All Files|*.*|Registry Hives|*.hiv;NTUSER.DAT;SAM;SECURITY;SOFTWARE;SYSTEM;USRCLASS.DAT;DEFAULT;Amcache.hve;BCD",
                FilterIndex = 1
            };

            if (dialog.ShowDialog() == true)
                await LoadHiveFileAsync(dialog.FileName, isLeft);
        }

        public async Task LoadHiveFileAsync(string filePath, bool isLeft)
        {
            // Cancel previous load for this side only
            if (isLeft)
            {
                _leftLoadCts?.Cancel();
                _leftLoadCts?.Dispose();
                _leftLoadCts = new CancellationTokenSource();
            }
            else
            {
                _rightLoadCts?.Cancel();
                _rightLoadCts?.Dispose();
                _rightLoadCts = new CancellationTokenSource();
            }
            var token = isLeft ? _leftLoadCts!.Token : _rightLoadCts!.Token;

            if (isLeft)
            {
                LeftIsLoading = true;
                LeftLoadProgress = 0;
                LeftStatusText = "Loading...";
            }
            else
            {
                RightIsLoading = true;
                RightLoadProgress = 0;
                RightStatusText = "Loading...";
            }

            try
            {
                var parser = new OfflineRegistryParser();

                var progress = new Progress<(string phase, double percent)>(update =>
                {
                    if (isLeft)
                        LeftLoadProgress = update.percent;
                    else
                        RightLoadProgress = update.percent;
                });

                await Task.Run(() => parser.LoadHive(filePath, progress, token), token).ConfigureAwait(true);

                if (isLeft)
                {
                    _leftParser?.Dispose();
                    _leftParser = parser;
                    _leftFilePath = filePath;
                    LeftFileName = $"{parser.FriendlyName}: {System.IO.Path.GetFileName(filePath)}";
                    LeftStatusText = "Loaded successfully";
                    LeftIsLoaded = true;

                    // Update right status hint
                    if (!RightIsLoaded)
                        RightStatusText = $"Select a {parser.FriendlyName} hive to compare";
                }
                else
                {
                    // Validate hive type matches left
                    if (_leftParser != null && _leftParser.CurrentHiveType != parser.CurrentHiveType)
                    {
                        parser.Dispose();
                        MessageBox.Show(
                            $"Hive type mismatch!\n\nFirst hive: {_leftParser.CurrentHiveType}\nSecond hive: {parser.CurrentHiveType}\n\nBoth hives must be the same type for comparison.",
                            "Type Mismatch",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                        RightStatusText = "Hive type mismatch";
                        return;
                    }

                    _rightParser?.Dispose();
                    _rightParser = parser;
                    _rightFilePath = filePath;
                    RightFileName = $"{parser.FriendlyName}: {System.IO.Path.GetFileName(filePath)}";
                    RightStatusText = "Loaded successfully";
                    RightIsLoaded = true;
                }
            }
            catch (OperationCanceledException)
            {
                if (isLeft) LeftStatusText = "Loading cancelled";
                else RightStatusText = "Loading cancelled";
            }
            catch (Exception ex)
            {
                var msg = $"Error loading hive: {ex.Message}";
                if (isLeft) LeftStatusText = msg;
                else RightStatusText = msg;
                MessageBox.Show(msg, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (isLeft) LeftIsLoading = false;
                else RightIsLoading = false;
            }
        }

        // ── Comparison ──

        private async Task RunComparisonAsync()
        {
            if (_leftParser == null || _rightParser == null) return;

            _compareCts?.Cancel();
            _compareCts?.Dispose();
            _compareCts = new CancellationTokenSource();
            var token = _compareCts.Token;

            IsComparing = true;
            CompareProgressText = "Building key index...";

            try
            {
                var leftParser = _leftParser;
                var rightParser = _rightParser;

                // Phase 1: Build key indexes
                await Task.Run(() =>
                {
                    token.ThrowIfCancellationRequested();
                    _leftKeyIndex = RegistryComparer.BuildKeyIndex(leftParser.GetRootKey());
                    token.ThrowIfCancellationRequested();
                    _rightKeyIndex = RegistryComparer.BuildKeyIndex(rightParser.GetRootKey());
                }, token).ConfigureAwait(true);

                CompareProgressText = "Analyzing differences...";

                // Phase 2: Compute diffs
                var leftKeys = _leftKeyIndex!;
                var rightKeys = _rightKeyIndex!;

                await Task.Run(() =>
                {
                    _leftDiffIndex = RegistryComparer.ComputeDiff(leftParser.GetRootKey(), rightKeys, token);
                    _rightDiffIndex = RegistryComparer.ComputeDiff(rightParser.GetRootKey(), leftKeys, token);
                }, token).ConfigureAwait(true);

                // Phase 3: Create root tree nodes
                LeftTreeRoots.Clear();
                RightTreeRoots.Clear();

                var leftRoot = CreateTreeNode(leftParser.GetRootKey()!, RegistryComparer.NormalizedRootName, isLeft: true);
                var rightRoot = CreateTreeNode(rightParser.GetRootKey()!, RegistryComparer.NormalizedRootName, isLeft: false);

                if (leftRoot != null)
                {
                    leftRoot.DisplayName = System.IO.Path.GetFileName(_leftFilePath);
                    LeftTreeRoots.Add(leftRoot);
                    leftRoot.IsExpanded = true;
                }
                if (rightRoot != null)
                {
                    rightRoot.DisplayName = System.IO.Path.GetFileName(_rightFilePath);
                    RightTreeRoots.Add(rightRoot);
                    rightRoot.IsExpanded = true;
                }

                // Switch to comparison view
                IsLandingView = false;
            }
            catch (OperationCanceledException)
            {
                // Cancelled — do nothing
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during comparison:\n{ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsComparing = false;
                CompareProgressText = "";
            }
        }

        // ── Tree Node Creation (lazy loading) ──

        /// <summary>
        /// Create a single CompareTreeNode using pre-computed diff info.
        /// If the key has visible children, adds a dummy sentinel child.
        /// Returns null if the node should be hidden (diff-only filter).
        /// </summary>
        private CompareTreeNode? CreateTreeNode(RegistryKey key, string path, bool isLeft)
        {
            var diffIndex = isLeft ? _leftDiffIndex : _rightDiffIndex;

            RegistryComparer.DiffInfo diffInfo = default;
            diffIndex?.TryGetValue(path, out diffInfo);

            // In differences-only mode, skip nodes with no differences
            if (_showDifferencesOnly && !diffInfo.HasDifference)
                return null;

            var node = new CompareTreeNode
            {
                DisplayName = key.KeyName,
                Path = path,
                Key = key,
                HasDifference = diffInfo.HasDifference,
                IsUniqueToThisHive = diffInfo.IsUniqueToThisHive,
                HasValueDifference = diffInfo.HasValueDifference
            };

            // Add dummy child if this key has visible children
            if (HasVisibleChildren(key, path, isLeft))
            {
                node.Children.Add(new CompareTreeNode { DisplayName = "Loading...", IsDummy = true });
            }

            return node;
        }

        private bool HasVisibleChildren(RegistryKey key, string parentPath, bool isLeft)
        {
            if (key.SubKeys == null || key.SubKeys.Count == 0)
                return false;

            if (!_showDifferencesOnly)
                return true;

            var diffIndex = isLeft ? _leftDiffIndex : _rightDiffIndex;
            if (diffIndex == null) return false;

            foreach (var sub in key.SubKeys)
            {
                var subPath = $"{parentPath}\\{sub.KeyName}";
                if (diffIndex.TryGetValue(subPath, out var diff) && diff.HasDifference)
                    return true;
            }
            return false;
        }

        /// <summary>
        /// Populate children of a node being expanded. Called from code-behind
        /// when TreeViewItem.Expanded fires and the node has a dummy child.
        /// </summary>
        public void PopulateChildren(CompareTreeNode parentNode, bool isLeft)
        {
            if (parentNode.Key.SubKeys == null) return;

            parentNode.Children.Clear();

            foreach (var subKey in parentNode.Key.SubKeys.OrderBy(k => k.KeyName, StringComparer.OrdinalIgnoreCase))
            {
                var path = $"{parentNode.Path}\\{subKey.KeyName}";
                var childNode = CreateTreeNode(subKey, path, isLeft);
                if (childNode != null)
                    parentNode.Children.Add(childNode);
            }
        }

        // ── Synchronized Navigation ──

        private void OnLeftNodeSelected(CompareTreeNode node)
        {
            if (_isSyncing || node.IsDummy) return;
            _isSyncing = true;
            try
            {
                LeftPathText = node.Path;
                LoadValues(LeftValues, node.Key, node.Path, isLeft: true);

                // Find and select matching node in right tree
                var rightNode = FindNodeByPath(RightTreeRoots, node.Path, isLeft: false);
                if (rightNode != null)
                {
                    rightNode.IsSelected = true;
                    RightPathText = rightNode.Path;
                    LoadValues(RightValues, rightNode.Key, rightNode.Path, isLeft: false);
                }
                else
                {
                    RightPathText = node.Path + " (not found)";
                    RightValues.Clear();
                }
            }
            finally
            {
                _isSyncing = false;
            }
        }

        private void OnRightNodeSelected(CompareTreeNode node)
        {
            if (_isSyncing || node.IsDummy) return;
            _isSyncing = true;
            try
            {
                RightPathText = node.Path;
                LoadValues(RightValues, node.Key, node.Path, isLeft: false);

                var leftNode = FindNodeByPath(LeftTreeRoots, node.Path, isLeft: true);
                if (leftNode != null)
                {
                    leftNode.IsSelected = true;
                    LeftPathText = leftNode.Path;
                    LoadValues(LeftValues, leftNode.Key, leftNode.Path, isLeft: true);
                }
                else
                {
                    LeftPathText = node.Path + " (not found)";
                    LeftValues.Clear();
                }
            }
            finally
            {
                _isSyncing = false;
            }
        }

        /// <summary>
        /// Sync expand/collapse from one tree to the other.
        /// Called from code-behind routed event handlers.
        /// </summary>
        public void SyncExpand(CompareTreeNode sourceNode, bool sourceIsLeft)
        {
            if (_isSyncing) return;
            _isSyncing = true;
            try
            {
                var targetRoots = sourceIsLeft ? RightTreeRoots : LeftTreeRoots;
                var targetNode = FindNodeByPath(targetRoots, sourceNode.Path, !sourceIsLeft);
                if (targetNode != null)
                    targetNode.IsExpanded = true;
            }
            finally
            {
                _isSyncing = false;
            }
        }

        public void SyncCollapse(CompareTreeNode sourceNode, bool sourceIsLeft)
        {
            if (_isSyncing) return;
            _isSyncing = true;
            try
            {
                var targetRoots = sourceIsLeft ? RightTreeRoots : LeftTreeRoots;
                var targetNode = FindNodeByPath(targetRoots, sourceNode.Path, !sourceIsLeft);
                if (targetNode != null)
                    targetNode.IsExpanded = false;
            }
            finally
            {
                _isSyncing = false;
            }
        }

        // ── Value Loading ──

        private void LoadValues(ObservableCollection<CompareValueItem> target, RegistryKey key, string keyPath, bool isLeft)
        {
            target.Clear();

            var otherIndex = isLeft ? _rightKeyIndex : _leftKeyIndex;
            if (otherIndex == null) return;

            var valueDiffs = RegistryComparer.GetValueDiffs(key, keyPath, otherIndex, _showDifferencesOnly);
            foreach (var diff in valueDiffs)
            {
                target.Add(new CompareValueItem
                {
                    Name = diff.Name,
                    Type = diff.Type,
                    Data = diff.Data,
                    ImageKey = RegistryValueItem.GetImageKey(diff.Type),
                    DiffStatus = diff.Status
                });
            }
        }

        // ── Value Selection Sync ──

        /// <summary>
        /// Called from code-behind when a value row is selected on one side.
        /// Finds and returns the index of the matching value name on the other side, or -1.
        /// </summary>
        public int FindMatchingValueIndex(string valueName, bool sourceIsLeft)
        {
            var target = sourceIsLeft ? RightValues : LeftValues;
            for (int i = 0; i < target.Count; i++)
            {
                if (string.Equals(target[i].Name, valueName, StringComparison.OrdinalIgnoreCase))
                    return i;
            }
            return -1;
        }

        // ── Tree Node Lookup ──

        /// <summary>
        /// Find a node by path in the tree, expanding lazy nodes as needed.
        /// </summary>
        private CompareTreeNode? FindNodeByPath(ObservableCollection<CompareTreeNode> roots, string path, bool isLeft)
        {
            if (roots.Count == 0) return null;

            var root = roots[0];
            if (string.Equals(path, root.Path, StringComparison.OrdinalIgnoreCase))
                return root;

            if (!path.StartsWith(root.Path + "\\", StringComparison.OrdinalIgnoreCase))
                return null;

            // Check key index to see if the path even exists
            var keyIndex = isLeft ? _leftKeyIndex : _rightKeyIndex;
            if (keyIndex == null || !keyIndex.ContainsKey(path))
                return null;

            // Check diff-only filter
            if (_showDifferencesOnly)
            {
                var diffIndex = isLeft ? _leftDiffIndex : _rightDiffIndex;
                if (diffIndex == null || !diffIndex.TryGetValue(path, out var diff) || !diff.HasDifference)
                    return null;
            }

            var remaining = path.Substring(root.Path.Length + 1);
            var segments = remaining.Split('\\');

            var current = root;
            foreach (var segment in segments)
            {
                // Ensure children are populated (lazy loading)
                if (current.Children.Count == 1 && current.Children[0].IsDummy)
                {
                    PopulateChildren(current, isLeft);
                    current.IsExpanded = true;
                }

                CompareTreeNode? found = null;
                foreach (var child in current.Children)
                {
                    if (child.DisplayName.Equals(segment, StringComparison.OrdinalIgnoreCase))
                    {
                        found = child;
                        break;
                    }
                }

                if (found == null) return null;
                current = found;
            }

            return current;
        }

        // ── Reset ──

        private void ResetToLanding()
        {
            _compareCts?.Cancel();

            LeftTreeRoots.Clear();
            RightTreeRoots.Clear();
            LeftValues.Clear();
            RightValues.Clear();
            LeftPathText = "";
            RightPathText = "";
            _leftKeyIndex = null;
            _rightKeyIndex = null;
            _leftDiffIndex = null;
            _rightDiffIndex = null;
            ShowDifferencesOnly = false;

            IsLandingView = true;
        }

        // ── Cleanup ──

        public void Dispose()
        {
            _compareCts?.Cancel();
            _compareCts?.Dispose();
            _leftLoadCts?.Cancel();
            _leftLoadCts?.Dispose();
            _rightLoadCts?.Cancel();
            _rightLoadCts?.Dispose();
            _leftParser?.Dispose();
            _rightParser?.Dispose();
            _leftKeyIndex?.Clear();
            _rightKeyIndex?.Clear();
            _leftDiffIndex?.Clear();
            _rightDiffIndex?.Clear();
        }
    }
}

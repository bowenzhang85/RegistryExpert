using System.ComponentModel;
using System.Diagnostics;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using RegistryExpert.Wpf.Helpers;
using RegistryExpert.Wpf.ViewModels;

namespace RegistryExpert.Wpf.Views
{
    public partial class MainWindow : Window
    {
        private MainViewModel ViewModel => (MainViewModel)DataContext;
        private SearchWindow? _searchWindow;
        private AnalyzeWindow? _analyzeWindow;
        private StatisticsWindow? _statisticsWindow;
        private CompareWindow? _compareWindow;
        private TimelineWindow? _timelineWindow;

        public MainWindow()
        {
            InitializeComponent();
        }

        // ── Window lifecycle ───────────────────────────────────────────────

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            ThemeManager.ApplyWindowChrome(this);
            ThemeManager.ThemeChanged += OnThemeChanged;
            ViewModel.RequestOpenSearch += OnRequestOpenSearch;
            ViewModel.RequestOpenAnalyze += OnRequestOpenAnalyze;
            ViewModel.RequestOpenStatistics += OnRequestOpenStatistics;
            ViewModel.RequestOpenCompare += OnRequestOpenCompare;
            ViewModel.RequestOpenTimeline += OnRequestOpenTimeline;
        }

        private void Window_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            ThemeManager.ThemeChanged -= OnThemeChanged;
            ViewModel.RequestOpenSearch -= OnRequestOpenSearch;
            ViewModel.RequestOpenAnalyze -= OnRequestOpenAnalyze;
            ViewModel.RequestOpenStatistics -= OnRequestOpenStatistics;
            ViewModel.RequestOpenCompare -= OnRequestOpenCompare;
            ViewModel.RequestOpenTimeline -= OnRequestOpenTimeline;
        }

        private void OnThemeChanged(object? sender, EventArgs e)
        {
            ThemeManager.ApplyWindowChrome(this);
        }

        // ── Search window ─────────────────────────────────────────────────

        private void OnRequestOpenSearch()
        {
            // If already open, just activate it
            if (_searchWindow != null && _searchWindow.IsLoaded)
            {
                _searchWindow.Activate();
                return;
            }

            _searchWindow = new SearchWindow(ViewModel);
            _searchWindow.Owner = this;
            _searchWindow.Closed += (s, e) => _searchWindow = null;
            _searchWindow.Show();
        }

        // ── Analyze window ────────────────────────────────────────────────

        private void OnRequestOpenAnalyze()
        {
            // If already open, just activate it
            if (_analyzeWindow != null && _analyzeWindow.IsLoaded)
            {
                _analyzeWindow.Activate();
                return;
            }

            try
            {
                var hives = ViewModel.LoadedHives.Values.ToList();
                _analyzeWindow = new AnalyzeWindow(hives);
                _analyzeWindow.Closed += (s, e) => _analyzeWindow = null;
                _analyzeWindow.Show();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error opening Analyze window: {ex}");
                System.Windows.MessageBox.Show(
                    $"Failed to open Analyze window:\n\n{ex.Message}",
                    "Error",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
                _analyzeWindow = null;
            }
        }

        // ── Statistics window ─────────────────────────────────────────────

        private void OnRequestOpenStatistics()
        {
            // If already open, just activate it
            if (_statisticsWindow != null && _statisticsWindow.IsLoaded)
            {
                _statisticsWindow.Activate();
                return;
            }

            try
            {
                var hives = ViewModel.LoadedHives.Values.ToList();
                _statisticsWindow = new StatisticsWindow(hives);
                _statisticsWindow.Closed += (s, e) => _statisticsWindow = null;
                _statisticsWindow.Show();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error opening Statistics window: {ex}");
                System.Windows.MessageBox.Show(
                    $"Failed to open Statistics window:\n\n{ex.Message}",
                    "Error",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
                _statisticsWindow = null;
            }
        }

        // ── Compare window ────────────────────────────────────────────────

        private void OnRequestOpenCompare()
        {
            if (_compareWindow != null && _compareWindow.IsLoaded)
            {
                _compareWindow.Activate();
                return;
            }

            try
            {
                var hives = ViewModel.LoadedHives.Values.ToList();
                _compareWindow = new CompareWindow(hives);
                _compareWindow.Closed += (s, e) => _compareWindow = null;
                _compareWindow.Show();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error opening Compare window: {ex}");
                System.Windows.MessageBox.Show(
                    $"Failed to open Compare window:\n\n{ex.Message}",
                    "Error",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
                _compareWindow = null;
            }
        }

        // ── Timeline window ───────────────────────────────────────────────

        private void OnRequestOpenTimeline()
        {
            if (_timelineWindow != null && _timelineWindow.IsLoaded)
            {
                _timelineWindow.Activate();
                return;
            }

            try
            {
                var hives = ViewModel.LoadedHives.Values.ToList();
                _timelineWindow = new TimelineWindow(hives, path => ViewModel.NavigateToKey(path));
                _timelineWindow.Closed += (s, e) => _timelineWindow = null;
                _timelineWindow.Show();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error opening Timeline window: {ex}");
                System.Windows.MessageBox.Show(
                    $"Failed to open Timeline window:\n\n{ex.Message}",
                    "Error",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
                _timelineWindow = null;
            }
        }

        // ── TreeView selection ─────────────────────────────────────────────

        private void RegistryTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (e.NewValue is RegistryKeyNode node)
            {
                ViewModel.SelectedTreeNode = node;
            }
        }

        // ── TreeView first-letter navigation ───────────────────────────────

        private void RegistryTree_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            // Only handle plain letter keys (no Ctrl, Alt, Shift modifiers)
            if (Keyboard.Modifiers != ModifierKeys.None)
                return;

            // Convert Key to a character; only handle A-Z
            var key = e.Key;
            if (key < Key.A || key > Key.Z)
                return;

            char letter = (char)('A' + (key - Key.A));

            var currentNode = ViewModel.SelectedTreeNode;
            if (currentNode == null)
                return;

            // Determine the sibling collection and the parent ItemsControl
            System.Collections.IList siblings;
            ItemsControl parentItemsControl;

            if (currentNode.Parent != null)
            {
                siblings = currentNode.Parent.Children;
                // Find the parent TreeViewItem container
                var parentTvi = FindTreeViewItemForNode(RegistryTree, currentNode.Parent);
                if (parentTvi == null)
                    return;
                parentItemsControl = parentTvi;
            }
            else
            {
                siblings = ViewModel.TreeRoots;
                parentItemsControl = RegistryTree;
            }

            // Build a list of matching siblings (case-insensitive first-letter match)
            var matches = new System.Collections.Generic.List<(RegistryKeyNode node, int index)>();
            for (int i = 0; i < siblings.Count; i++)
            {
                if (siblings[i] is RegistryKeyNode sibling
                    && sibling.DisplayName.Length > 0
                    && char.ToUpperInvariant(sibling.DisplayName[0]) == letter)
                {
                    matches.Add((sibling, i));
                }
            }

            if (matches.Count == 0)
                return;

            // Find current node among matches, then pick the next one (wrap around)
            int currentMatchIdx = -1;
            for (int m = 0; m < matches.Count; m++)
            {
                if (matches[m].node == currentNode)
                {
                    currentMatchIdx = m;
                    break;
                }
            }

            int nextMatchIdx = (currentMatchIdx + 1) % matches.Count;

            // If only match is the current node, nothing to do
            if (matches.Count == 1 && currentMatchIdx == 0)
                return;

            var (target, siblingIndex) = matches[nextMatchIdx];

            // Force the virtualizing panel to realize the target container
            var panel = FindVisualChild<VirtualizingStackPanel>(parentItemsControl);
            if (panel != null)
            {
                panel.BringIndexIntoViewPublic(siblingIndex);
            }

            // Now the container should exist — select it and scroll into view
            target.IsSelected = true;

            if (parentItemsControl.ItemContainerGenerator.ContainerFromItem(target)
                is TreeViewItem tvi)
            {
                tvi.BringIntoView();
                tvi.Focus();
            }

            e.Handled = true;
        }

        /// <summary>Find the TreeViewItem container for a given data node.</summary>
        private static TreeViewItem? FindTreeViewItemForNode(ItemsControl parent, RegistryKeyNode node)
        {
            // Direct lookup (works if container is realized)
            if (parent.ItemContainerGenerator.ContainerFromItem(node) is TreeViewItem tvi)
                return tvi;

            // Walk realized containers looking for the node in their subtree
            for (int i = 0; i < parent.Items.Count; i++)
            {
                if (parent.ItemContainerGenerator.ContainerFromIndex(i) is TreeViewItem child)
                {
                    if (child.DataContext == node)
                        return child;

                    if (child.IsExpanded)
                    {
                        var found = FindTreeViewItemForNode(child, node);
                        if (found != null)
                            return found;
                    }
                }
            }

            return null;
        }

        /// <summary>Find the first visual child of the given type in the visual tree.</summary>
        private static T? FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            int count = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < count; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T found)
                    return found;

                var result = FindVisualChild<T>(child);
                if (result != null)
                    return result;
            }
            return null;
        }

        // ── Drag and Drop ──────────────────────────────────────────────────

        private void Window_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true;
        }

        private async void Window_Drop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
                return;

            if (e.Data.GetData(DataFormats.FileDrop) is string[] files)
            {
                foreach (var file in files)
                {
                    await ViewModel.LoadHiveFileAsync(file);
                }
            }
        }

        // ── Unload Hive submenu ────────────────────────────────────────────

        private void UnloadHiveMenu_SubmenuOpened(object sender, RoutedEventArgs e)
        {
            if (sender is not MenuItem menuItem)
                return;

            menuItem.Items.Clear();

            var loadedHives = ViewModel.LoadedHives;
            if (loadedHives.Count == 0)
            {
                menuItem.Items.Add(new MenuItem
                {
                    Header = "(no hives loaded)",
                    IsEnabled = false
                });
                return;
            }

            foreach (var kvp in loadedHives.OrderBy(h => h.Key.ToString()))
            {
                var hiveType = kvp.Key;
                var hiveInfo = kvp.Value;
                var item = new MenuItem
                {
                    Header = $"{hiveType} — {System.IO.Path.GetFileName(hiveInfo.FilePath)}"
                };
                item.Click += (s, args) => ViewModel.CloseHiveCommand.Execute(hiveType);
                menuItem.Items.Add(item);
            }
        }

        // ── Values grid double-click ───────────────────────────────────────

        private void ValuesGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var val = ViewModel.SelectedValue;
            if (val == null) return;

            // Build detail text with hex dump
            var sb = new StringBuilder();
            sb.AppendLine($"Name:    {val.Name}");
            sb.AppendLine($"Type:    {val.Type}");
            sb.AppendLine($"Slack:   {val.SlackSize} bytes");
            sb.AppendLine();
            sb.AppendLine("Data:");
            sb.AppendLine(val.Data);

            if (val.RawBytes.Length > 0)
            {
                sb.AppendLine();
                sb.AppendLine("Hex Dump:");
                sb.Append(MainViewModel.FormatHexDump(val.RawBytes));
            }

            // Show modal dialog
            var dialog = new Window
            {
                Title = $"Value: {val.Name}",
                Width = 640,
                Height = 480,
                MinWidth = 400,
                MinHeight = 300,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this,
                Style = (Style)FindResource("ModernWindowStyle")
            };

            var textBox = new TextBox
            {
                Text = sb.ToString(),
                IsReadOnly = true,
                FontFamily = new System.Windows.Media.FontFamily("Consolas"),
                FontSize = 12,
                TextWrapping = TextWrapping.Wrap,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                Background = (System.Windows.Media.Brush)FindResource("SurfaceBrush"),
                Foreground = (System.Windows.Media.Brush)FindResource("TextPrimaryBrush"),
                CaretBrush = (System.Windows.Media.Brush)FindResource("TextPrimaryBrush"),
                BorderThickness = new Thickness(0),
                Padding = new Thickness(12, 8, 12, 8)
            };

            dialog.Content = textBox;

            // Apply dark title bar to the dialog
            dialog.Loaded += (s, ev) => ThemeManager.ApplyWindowChrome(dialog);

            dialog.ShowDialog();
        }

        // ── Bookmark interactions ──────────────────────────────────────────

        private void BookmarkCollapsedBar_MouseDown(object sender, MouseButtonEventArgs e)
        {
            ViewModel.ToggleBookmarksCommand.Execute(null);
        }

        private void BookmarkItem_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is FrameworkElement fe && fe.DataContext is MainViewModel.BookmarkItem bookmark)
            {
                ViewModel.NavigateToKeyCommand.Execute(bookmark.Path);
            }
        }
    }
}

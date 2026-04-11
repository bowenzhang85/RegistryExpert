using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using RegistryExpert.Core;
using RegistryExpert.Core.Models;
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

            // Restore saved theme
            var settings = ViewModel.Settings;
            if (settings.Theme == "Light")
                ThemeManager.SetTheme(ThemeManager.Theme.Light);

            ThemeManager.ApplyWindowChrome(this);
            ThemeManager.ThemeChanged += OnThemeChanged;

            ViewModel.RequestOpenSearch += OnRequestOpenSearch;
            ViewModel.RequestOpenAnalyze += OnRequestOpenAnalyze;
            ViewModel.RequestOpenStatistics += OnRequestOpenStatistics;
            ViewModel.RequestOpenCompare += OnRequestOpenCompare;
            ViewModel.RequestOpenTimeline += OnRequestOpenTimeline;
            ViewModel.RequestOpenAbout += OnRequestOpenAbout;
            ViewModel.RequestShowUpdateResult += OnRequestShowUpdateResult;
            ViewModel.RequestScrollToNode += OnRequestScrollToNode;
            ViewModel.RequestShowHivePicker += OnRequestShowHivePicker;
            ViewModel.RequestShowRecentBundles += OnRequestShowRecentBundles;

            // Auto-check for updates on startup
            _ = CheckForUpdatesOnStartupAsync();
        }

        private void Window_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            // Save settings
            var settings = ViewModel.Settings;
            settings.Save();

            ThemeManager.ThemeChanged -= OnThemeChanged;
            ViewModel.RequestOpenSearch -= OnRequestOpenSearch;
            ViewModel.RequestOpenAnalyze -= OnRequestOpenAnalyze;
            ViewModel.RequestOpenStatistics -= OnRequestOpenStatistics;
            ViewModel.RequestOpenCompare -= OnRequestOpenCompare;
            ViewModel.RequestOpenTimeline -= OnRequestOpenTimeline;
            ViewModel.RequestOpenAbout -= OnRequestOpenAbout;
            ViewModel.RequestShowUpdateResult -= OnRequestShowUpdateResult;
            ViewModel.RequestScrollToNode -= OnRequestScrollToNode;
            ViewModel.RequestShowHivePicker -= OnRequestShowHivePicker;
            ViewModel.RequestShowRecentBundles -= OnRequestShowRecentBundles;
        }

        private void OnThemeChanged(object? sender, EventArgs e)
        {
            ThemeManager.ApplyWindowChrome(this);
        }

        // ── Hive picker dialog ────────────────────────────────────────────

        private List<DiscoveredHive>? OnRequestShowHivePicker(List<DiscoveredHive> discovered)
        {
            var items = discovered.Select(h => new HivePickerItem(h)).ToList();
            var picker = new HivePickerWindow(items)
            {
                Owner = this
            };

            if (picker.ShowDialog() == true)
            {
                return picker.SelectedItems
                    .Select(i => i.Hive)
                    .ToList();
            }

            return null;
        }

        private (BundleInfo? Selected, bool BrowseRequested) OnRequestShowRecentBundles(List<BundleInfo> bundles)
        {
            var dialog = new RecentBundlesWindow(bundles)
            {
                Owner = this
            };

            if (dialog.ShowDialog() == true)
                return (dialog.SelectedBundle, false);

            return (null, dialog.BrowseRequested);
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
            _searchWindow.Closed += (s, e) => _searchWindow = null;
            _searchWindow.Show();
        }

        // ── Scroll-to-node (search navigation) ───────────────────────────

        private void OnRequestScrollToNode(RegistryKeyNode node, string? valueName)
        {
            // Build the ancestor chain from root to target node
            var chain = new System.Collections.Generic.List<RegistryKeyNode>();
            var current = node;
            while (current != null)
            {
                chain.Add(current);
                current = current.Parent;
            }
            chain.Reverse(); // root -> ... -> target

            // Defer until layout is complete, then force-realize each container along the path
            Dispatcher.BeginInvoke(DispatcherPriority.Loaded, () =>
            {
                ItemsControl container = RegistryTree;

                foreach (var pathNode in chain)
                {
                    // Force the virtualizing panel to realize this item's container
                    int index = container.Items.IndexOf(pathNode);
                    if (index >= 0)
                    {
                        var panel = FindVisualChild<VirtualizingStackPanel>(container);
                        panel?.BringIndexIntoViewPublic(index);
                    }

                    // Now get the realized TreeViewItem
                    if (container.ItemContainerGenerator.ContainerFromItem(pathNode) is not TreeViewItem tvi)
                        break;

                    tvi.IsExpanded = true;
                    tvi.UpdateLayout(); // force child containers to be generated

                    if (pathNode == node)
                    {
                        // Final node — select, scroll into view, and focus
                        tvi.IsSelected = true;
                        tvi.BringIntoView();
                        tvi.Focus();
                    }

                    container = tvi;
                }

                // Scroll the ValuesGrid to the selected value
                if (valueName != null && ViewModel.SelectedValue != null)
                {
                    ValuesGrid.ScrollIntoView(ViewModel.SelectedValue);
                }
            });
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

        // ── About window ─────────────────────────────────────────────────

        private void OnRequestOpenAbout()
        {
            var aboutWindow = new AboutWindow();
            aboutWindow.Owner = this;
            aboutWindow.ShowDialog();
        }

        // ── Update check ──────────────────────────────────────────────────

        private void OnRequestShowUpdateResult(UpdateInfo? info, bool isManualCheck)
        {
            if (info == null)
            {
                if (isManualCheck)
                    MessageBox.Show("Unable to check for updates. Please check your internet connection.",
                        "Update Check Failed", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!info.UpdateAvailable)
            {
                if (isManualCheck)
                    MessageBox.Show($"You're up to date!\n\nRegistry Expert {info.CurrentVersion} is the latest version.",
                        "No Updates Available", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            ShowUpdateAvailableDialog(info);
        }

        private void ShowUpdateAvailableDialog(UpdateInfo info)
        {
            var dialog = new Window
            {
                Title = "Update Available",
                Width = 500,
                Height = 450,
                ResizeMode = ResizeMode.NoResize,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this,
                ShowInTaskbar = false,
                Style = (Style)FindResource("ModernWindowStyle")
            };

            var mainPanel = new StackPanel { Margin = new Thickness(24) };

            // Title
            mainPanel.Children.Add(new TextBlock
            {
                Text = "A new version is available!",
                FontFamily = new FontFamily("Segoe UI"),
                FontWeight = FontWeights.SemiBold,
                FontSize = 14,
                Foreground = (Brush)FindResource("AccentBrush"),
                Margin = new Thickness(0, 0, 0, 16)
            });

            // Version panel
            var versionBorder = new Border
            {
                Background = (Brush)FindResource("SurfaceBrush"),
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(16, 12, 16, 12),
                Margin = new Thickness(0, 0, 0, 16)
            };
            var versionPanel = new StackPanel();
            versionPanel.Children.Add(new TextBlock
            {
                Text = $"Current version: {info.CurrentVersion}",
                Foreground = (Brush)FindResource("TextSecondaryBrush"),
                FontSize = 13,
                Margin = new Thickness(0, 0, 0, 4)
            });
            versionPanel.Children.Add(new TextBlock
            {
                Text = $"Latest version: {info.LatestVersion}",
                Foreground = (Brush)FindResource("TextPrimaryBrush"),
                FontWeight = FontWeights.Bold,
                FontSize = 13
            });
            versionBorder.Child = versionPanel;
            mainPanel.Children.Add(versionBorder);

            // Release notes header
            mainPanel.Children.Add(new TextBlock
            {
                Text = "New Features",
                FontWeight = FontWeights.Bold,
                Foreground = (Brush)FindResource("TextPrimaryBrush"),
                FontSize = 13,
                Margin = new Thickness(0, 0, 0, 8)
            });

            // Release notes
            var notesBox = new TextBox
            {
                Text = info.ReleaseNotes,
                IsReadOnly = true,
                TextWrapping = TextWrapping.Wrap,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Background = (Brush)FindResource("SurfaceBrush"),
                Foreground = (Brush)FindResource("TextPrimaryBrush"),
                CaretBrush = (Brush)FindResource("TextPrimaryBrush"),
                BorderThickness = new Thickness(1),
                BorderBrush = (Brush)FindResource("BorderBrush"),
                Padding = new Thickness(12, 8, 12, 8),
                FontSize = 12,
                Height = 140
            };
            mainPanel.Children.Add(notesBox);

            // Button panel
            var buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 16, 0, 0)
            };

            var downloadBtn = new Button
            {
                Content = "Download",
                Style = (Style)FindResource("AccentButtonStyle"),
                Width = 100,
                Margin = new Thickness(0, 0, 8, 0)
            };
            downloadBtn.Click += (s, e) =>
            {
                if (!string.IsNullOrEmpty(info.ReleaseUrl) && info.ReleaseUrl.StartsWith("https://github.com/"))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = info.ReleaseUrl,
                        UseShellExecute = true
                    });
                }
                dialog.Close();
            };
            buttonPanel.Children.Add(downloadBtn);

            var laterBtn = new Button
            {
                Content = "Later",
                Style = (Style)FindResource("SecondaryButtonStyle"),
                Width = 100
            };
            laterBtn.Click += (s, e) => dialog.Close();
            buttonPanel.Children.Add(laterBtn);

            mainPanel.Children.Add(buttonPanel);

            dialog.Content = mainPanel;
            dialog.Loaded += (s, e) => ThemeManager.ApplyWindowChrome(dialog);
            dialog.ShowDialog();
        }

        private async Task CheckForUpdatesOnStartupAsync()
        {
            try
            {
                await Task.Delay(2000);
                var info = await UpdateChecker.CheckForUpdatesAsync();
                if (info?.UpdateAvailable == true)
                    ShowUpdateAvailableDialog(info);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Startup update check failed: {ex.Message}");
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

            if (e.Data.GetData(DataFormats.FileDrop) is string[] paths)
            {
                foreach (var path in paths)
                {
                    if (Directory.Exists(path))
                    {
                        // It's a folder — scan it for hive files
                        await ViewModel.LoadHivesFromFolderAsync(path);
                    }
                    else
                    {
                        // It's a file — load it directly
                        await ViewModel.LoadHiveFileAsync(path);
                    }
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

            foreach (var kvp in loadedHives.OrderBy(h => h.Value.RootNode.DisplayName))
            {
                var hiveKey = kvp.Key;
                var hiveInfo = kvp.Value;
                var displayName = hiveInfo.RootNode.DisplayName;
                var item = new MenuItem
                {
                    Header = $"{displayName} — {System.IO.Path.GetFileName(hiveInfo.FilePath)}"
                };
                item.Click += (s, args) => ViewModel.CloseHiveCommand.Execute(hiveKey);
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

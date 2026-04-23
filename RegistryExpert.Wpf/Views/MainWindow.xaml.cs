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

            // If launched after an auto-update (or if version changed since last
            // run), show the green "Updated successfully" banner at the top.
            CheckAndShowUpdatedBanner();
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

        // Once a downloaded update is ready (silent background flow), we cache
        // the local path here so the "Restart & Update" button can launch it.
        private string? _pendingUpdateLocalPath;
        private UpdateInfo? _pendingUpdateInfo;

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

            // If we already silently downloaded this version, jump to the
            // "Update Ready to Install" dialog directly.
            if (AutoUpdater.IsUpdateAlreadyDownloaded(info))
            {
                _pendingUpdateInfo = info;
                _pendingUpdateLocalPath = AutoUpdater.GetDownloadCachePath(info.LatestVersion);
                ShowUpdateReadyDialog(info);
            }
            else
            {
                ShowUpdateAvailableDialog(info);
            }
        }

        // Dialog used when the update has NOT yet been downloaded.
        // Provides "Download & Install" (in-app) and "View on GitHub" (browser).
        private void ShowUpdateAvailableDialog(UpdateInfo info)
        {
            var dialog = BuildUpdateDialogShell(
                title: "Update Available",
                heading: "A new version is available!",
                info: info,
                out var buttonPanel,
                out var progressHost);

            // Progress UI (hidden until Download clicked)
            var progressBar = new ProgressBar
            {
                Height = 8,
                Minimum = 0,
                Maximum = 1,
                Value = 0,
                Foreground = (Brush)FindResource("AccentBrush"),
                Background = (Brush)FindResource("SurfaceBrush"),
                BorderThickness = new Thickness(0),
                Margin = new Thickness(0, 0, 0, 6),
                Visibility = Visibility.Collapsed
            };
            var progressLabel = new TextBlock
            {
                Foreground = (Brush)FindResource("TextSecondaryBrush"),
                FontSize = 12,
                Visibility = Visibility.Collapsed
            };
            progressHost.Children.Add(progressBar);
            progressHost.Children.Add(progressLabel);

            var downloadBtn = new Button
            {
                Content = "Download & Install",
                Style = (Style)FindResource("AccentButtonStyle"),
                Width = 150,
                Margin = new Thickness(0, 0, 8, 0)
            };
            var githubBtn = new Button
            {
                Content = "View on GitHub",
                Style = (Style)FindResource("SecondaryButtonStyle"),
                Width = 130,
                Margin = new Thickness(0, 0, 8, 0)
            };
            var laterBtn = new Button
            {
                Content = "Later",
                Style = (Style)FindResource("SecondaryButtonStyle"),
                Width = 80
            };

            CancellationTokenSource? cts = null;

            downloadBtn.Click += async (s, e) =>
            {
                if (string.IsNullOrEmpty(info.DownloadUrl))
                {
                    MessageBox.Show(dialog,
                        "This release does not include a downloadable RegistryExpert.exe asset.\nPlease use 'View on GitHub' instead.",
                        "Download Unavailable", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                downloadBtn.IsEnabled = false;
                githubBtn.IsEnabled = false;
                laterBtn.Content = "Cancel";
                progressBar.Visibility = Visibility.Visible;
                progressLabel.Visibility = Visibility.Visible;
                progressLabel.Text = "Starting download...";

                cts = new CancellationTokenSource();
                var progress = new Progress<double>(p =>
                {
                    progressBar.Value = p;
                    progressLabel.Text = $"Downloading... {p * 100:0}%";
                });

                laterBtn.Click -= LaterCloseHandler; // detach close
                EventHandler? cancelHandler = null;
                cancelHandler = (cs, ce) => cts?.Cancel();
                laterBtn.Click += (cs, ce) => cancelHandler?.Invoke(cs, EventArgs.Empty);

                var localPath = await AutoUpdater.DownloadUpdateAsync(info, progress, cts.Token);

                if (localPath == null)
                {
                    if (cts.IsCancellationRequested)
                    {
                        dialog.Close();
                        return;
                    }
                    MessageBox.Show(dialog,
                        "The update could not be downloaded or failed verification.\nPlease try again later or use 'View on GitHub'.",
                        "Download Failed", MessageBoxButton.OK, MessageBoxImage.Error);
                    downloadBtn.IsEnabled = true;
                    githubBtn.IsEnabled = true;
                    laterBtn.Content = "Later";
                    progressBar.Visibility = Visibility.Collapsed;
                    progressLabel.Visibility = Visibility.Collapsed;
                    return;
                }

                _pendingUpdateInfo = info;
                _pendingUpdateLocalPath = localPath;

                dialog.Close();
                ShowUpdateReadyDialog(info);
            };

            githubBtn.Click += (s, e) =>
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

            void LaterCloseHandler(object s, RoutedEventArgs e) => dialog.Close();
            laterBtn.Click += LaterCloseHandler;

            buttonPanel.Children.Add(downloadBtn);
            buttonPanel.Children.Add(githubBtn);
            buttonPanel.Children.Add(laterBtn);

            dialog.Loaded += (s, e) => ThemeManager.ApplyWindowChrome(dialog);
            dialog.ShowDialog();
        }

        // Dialog used once the update has been downloaded & verified.
        // Provides "Restart & Update" and "Later".
        private void ShowUpdateReadyDialog(UpdateInfo info)
        {
            var dialog = BuildUpdateDialogShell(
                title: "Update Ready to Install",
                heading: "Update downloaded \u2014 ready to install",
                info: info,
                out var buttonPanel,
                out _);

            var restartBtn = new Button
            {
                Content = "Restart & Update",
                Style = (Style)FindResource("AccentButtonStyle"),
                Width = 150,
                Margin = new Thickness(0, 0, 8, 0)
            };
            var laterBtn = new Button
            {
                Content = "Later",
                Style = (Style)FindResource("SecondaryButtonStyle"),
                Width = 100
            };

            restartBtn.Click += (s, e) =>
            {
                dialog.Close();
                ApplyPendingUpdate();
            };
            laterBtn.Click += (s, e) => dialog.Close();

            buttonPanel.Children.Add(restartBtn);
            buttonPanel.Children.Add(laterBtn);

            dialog.Loaded += (s, e) => ThemeManager.ApplyWindowChrome(dialog);
            dialog.ShowDialog();
        }

        // Shared layout used by both update dialogs above.
        // Returns the dialog window plus the empty button panel and an empty
        // host StackPanel where progress UI can be inserted (above the buttons).
        private Window BuildUpdateDialogShell(
            string title,
            string heading,
            UpdateInfo info,
            out StackPanel buttonPanel,
            out StackPanel progressHost)
        {
            var dialog = new Window
            {
                Title = title,
                Width = 500,
                Height = 470,
                ResizeMode = ResizeMode.NoResize,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this,
                ShowInTaskbar = false,
                Style = (Style)FindResource("ModernWindowStyle")
            };

            var mainPanel = new StackPanel { Margin = new Thickness(24) };

            mainPanel.Children.Add(new TextBlock
            {
                Text = heading,
                FontFamily = new FontFamily("Segoe UI"),
                FontWeight = FontWeights.SemiBold,
                FontSize = 14,
                Foreground = (Brush)FindResource("AccentBrush"),
                Margin = new Thickness(0, 0, 0, 16)
            });

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
            if (UpdateChecker.IsUsingOverrideUrl)
            {
                versionPanel.Children.Add(new TextBlock
                {
                    Text = "(Update source overridden via REGEXPERT_UPDATE_URL)",
                    Foreground = (Brush)FindResource("TextSecondaryBrush"),
                    FontStyle = FontStyles.Italic,
                    FontSize = 11,
                    Margin = new Thickness(0, 6, 0, 0)
                });
            }
            versionBorder.Child = versionPanel;
            mainPanel.Children.Add(versionBorder);

            mainPanel.Children.Add(new TextBlock
            {
                Text = "Release Notes",
                FontWeight = FontWeights.Bold,
                Foreground = (Brush)FindResource("TextPrimaryBrush"),
                FontSize = 13,
                Margin = new Thickness(0, 0, 0, 8)
            });

            mainPanel.Children.Add(new TextBox
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
                Height = 130
            });

            // Host for optional progress UI (used by ShowUpdateAvailableDialog)
            progressHost = new StackPanel { Margin = new Thickness(0, 12, 0, 0) };
            mainPanel.Children.Add(progressHost);

            buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 16, 0, 0)
            };
            mainPanel.Children.Add(buttonPanel);

            dialog.Content = mainPanel;
            return dialog;
        }

        // Launches the updater script and shuts down the application.
        // Detects whether elevation is required and prompts the user accordingly.
        // Shows a brief modal "Installing update..." dialog before exiting so
        // the user gets visual feedback during the swap.
        private void ApplyPendingUpdate()
        {
            if (string.IsNullOrEmpty(_pendingUpdateLocalPath) || !File.Exists(_pendingUpdateLocalPath))
            {
                MessageBox.Show("The downloaded update file could not be found. Please try again.",
                    "Update Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var currentExe = AutoUpdater.GetCurrentExecutablePath();
            var canWrite = AutoUpdater.CanWriteToInstallLocation(currentExe);
            bool elevated = false;

            if (!canWrite)
            {
                var resp = MessageBox.Show(this,
                    "RegistryExpert needs administrator rights to install this update because it is " +
                    "located in a protected folder.\n\nClick Yes to continue (a UAC prompt will appear), " +
                    "or No to cancel.",
                    "Administrator Rights Required",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);
                if (resp != MessageBoxResult.Yes) return;
                elevated = true;
            }

            // Show modal "Installing update..." with spinner; after a brief delay
            // (so the user sees the message), launch the updater and shut down.
            var installingDialog = BuildInstallingDialog();
            var localExe = _pendingUpdateLocalPath!;
            var fromVersion = UpdateChecker.GetCurrentVersion();

            installingDialog.Loaded += async (s, e) =>
            {
                // Allow the user ~700ms to see the "installing" message
                await Task.Delay(700);

                var ok = AutoUpdater.LaunchUpdaterAndExit(localExe, currentExe, elevated, fromVersion);
                if (!ok)
                {
                    installingDialog.Close();
                    MessageBox.Show(this,
                        "Failed to launch the update process. The update was not applied.",
                        "Update Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // The updater script waits for our process to exit before swapping the file.
                Application.Current.Shutdown();
            };

            installingDialog.ShowDialog();
        }

        // Compact "Installing update..." modal with indeterminate progress bar.
        // No close button, no chrome buttons -- displayed briefly before exit.
        private Window BuildInstallingDialog()
        {
            var dialog = new Window
            {
                Title = "Installing Update",
                Width = 380,
                Height = 200,
                ResizeMode = ResizeMode.NoResize,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this,
                ShowInTaskbar = false,
                Style = (Style)FindResource("ModernWindowStyle")
            };

            var stack = new StackPanel
            {
                Margin = new Thickness(28, 24, 28, 24),
                VerticalAlignment = VerticalAlignment.Center
            };

            stack.Children.Add(new TextBlock
            {
                Text = "Installing update\u2026",
                FontFamily = new FontFamily("Segoe UI"),
                FontWeight = FontWeights.SemiBold,
                FontSize = 15,
                Foreground = (Brush)FindResource("AccentBrush"),
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 0, 0, 14)
            });

            stack.Children.Add(new ProgressBar
            {
                IsIndeterminate = true,
                Height = 4,
                Foreground = (Brush)FindResource("AccentBrush"),
                Background = (Brush)FindResource("SurfaceBrush"),
                BorderThickness = new Thickness(0),
                Margin = new Thickness(0, 0, 0, 14)
            });

            stack.Children.Add(new TextBlock
            {
                Text = "Registry Expert will restart automatically in a moment.",
                Foreground = (Brush)FindResource("TextSecondaryBrush"),
                FontSize = 12,
                TextWrapping = TextWrapping.Wrap,
                HorizontalAlignment = HorizontalAlignment.Center,
                TextAlignment = TextAlignment.Center
            });

            dialog.Content = stack;
            dialog.Loaded += (s, e) => ThemeManager.ApplyWindowChrome(dialog);
            return dialog;
        }

        // ── Post-update success banner ────────────────────────────────────

        private DispatcherTimer? _bannerHideTimer;

        // Decides whether to show the banner based on:
        //   1. App.UpgradedFromVersion (set when launched with --just-updated arg)
        //   2. Settings.LastSeenVersion vs current version (sideload fallback)
        private void CheckAndShowUpdatedBanner()
        {
            try
            {
                var settings = ViewModel.Settings;
                var currentVersion = UpdateChecker.GetCurrentVersion();
                var upgradedFrom = App.UpgradedFromVersion;

                if (!string.IsNullOrEmpty(upgradedFrom))
                {
                    // Auto-update detected: full banner with release notes button
                    ShowUpdatedBanner(
                        $"Updated to Registry Expert {currentVersion} (from {upgradedFrom})",
                        showReleaseNotesButton: true);
                }
                else if (!string.IsNullOrEmpty(settings.LastSeenVersion)
                    && !string.Equals(settings.LastSeenVersion, currentVersion, StringComparison.Ordinal))
                {
                    // Manual / sideload upgrade: simpler welcome banner.
                    // Only show on UPGRADE (current > last); suppress on downgrade or unparseable.
                    if (Version.TryParse(currentVersion, out var curV)
                        && Version.TryParse(settings.LastSeenVersion, out var lastV)
                        && curV > lastV)
                    {
                        ShowUpdatedBanner(
                            $"Welcome to Registry Expert {currentVersion}",
                            showReleaseNotesButton: false);
                    }
                }

                // Always remember current version so we don't repeat the banner next launch
                if (!string.Equals(settings.LastSeenVersion, currentVersion, StringComparison.Ordinal))
                {
                    settings.LastSeenVersion = currentVersion;
                    settings.Save();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"CheckAndShowUpdatedBanner failed: {ex.Message}");
            }
        }

        private void ShowUpdatedBanner(string text, bool showReleaseNotesButton)
        {
            UpdatedBannerText.Text = "\u2713  " + text;
            UpdatedBannerReleaseNotesBtn.Visibility = showReleaseNotesButton
                ? Visibility.Visible : Visibility.Collapsed;
            UpdatedBanner.Visibility = Visibility.Visible;

            // Auto-hide after 12 seconds
            _bannerHideTimer?.Stop();
            _bannerHideTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(12) };
            _bannerHideTimer.Tick += (s, e) =>
            {
                _bannerHideTimer?.Stop();
                HideUpdatedBanner();
            };
            _bannerHideTimer.Start();
        }

        private void HideUpdatedBanner()
        {
            UpdatedBanner.Visibility = Visibility.Collapsed;
            _bannerHideTimer?.Stop();
            _bannerHideTimer = null;
        }

        private void UpdatedBannerCloseBtn_Click(object sender, RoutedEventArgs e)
            => HideUpdatedBanner();

        private void UpdatedBannerReleaseNotesBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var version = UpdateChecker.GetCurrentVersion();

                // Smart URL: if REGEXPERT_UPDATE_URL is set, derive the test repo
                // release URL from it; otherwise use the production repo URL.
                string url;
                if (UpdateChecker.IsUsingOverrideUrl)
                {
                    // Try to derive owner/repo from the override URL.
                    // Expected shape: https://api.github.com/repos/<owner>/<repo>/releases/latest
                    var apiUrl = UpdateChecker.GitHubApiUrl;
                    var marker = "/repos/";
                    var ownerStart = apiUrl.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
                    if (ownerStart >= 0)
                    {
                        var rest = apiUrl.Substring(ownerStart + marker.Length);
                        var slash = rest.IndexOf('/');
                        if (slash > 0)
                        {
                            var owner = rest.Substring(0, slash);
                            var afterOwner = rest.Substring(slash + 1);
                            var nextSlash = afterOwner.IndexOf('/');
                            var repo = nextSlash > 0 ? afterOwner.Substring(0, nextSlash) : afterOwner;
                            url = $"https://github.com/{owner}/{repo}/releases/tag/v{version}";
                        }
                        else
                        {
                            url = $"https://github.com/bowenzhang85/RegistryExpert/releases/tag/v{version}";
                        }
                    }
                    else
                    {
                        url = $"https://github.com/bowenzhang85/RegistryExpert/releases/tag/v{version}";
                    }
                }
                else
                {
                    url = $"https://github.com/bowenzhang85/RegistryExpert/releases/tag/v{version}";
                }

                Process.Start(new ProcessStartInfo { FileName = url, UseShellExecute = true });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Release notes link failed: {ex.Message}");
            }
        }

        private async Task CheckForUpdatesOnStartupAsync()
        {
            try
            {
                await Task.Delay(2000);
                var info = await UpdateChecker.CheckForUpdatesAsync();
                if (info == null || !info.UpdateAvailable) return;

                var settings = ViewModel.Settings;

                // Best-effort: clean up cache folders for older versions
                AutoUpdater.CleanupOldCaches(info.LatestVersion);

                if (settings.AutoDownloadUpdates && !string.IsNullOrEmpty(info.DownloadUrl))
                {
                    string? localPath;
                    if (AutoUpdater.IsUpdateAlreadyDownloaded(info))
                    {
                        localPath = AutoUpdater.GetDownloadCachePath(info.LatestVersion);
                    }
                    else
                    {
                        // Silent background download (no progress UI on startup)
                        localPath = await AutoUpdater.DownloadUpdateAsync(info, progress: null, cancellationToken: default);
                    }

                    if (localPath != null)
                    {
                        _pendingUpdateInfo = info;
                        _pendingUpdateLocalPath = localPath;
                        ShowUpdateReadyDialog(info);
                        return;
                    }
                    // If silent download failed, fall through to the manual prompt
                }

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

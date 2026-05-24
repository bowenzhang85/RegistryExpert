using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using RegistryExpert.Core;
using RegistryExpert.Core.Models;
using RegistryExpert.Core.Services;
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
            ViewModel.RequestShowReleaseNotes += OnRequestShowReleaseNotes;
            ViewModel.RequestScrollToNode += OnRequestScrollToNode;
            ViewModel.RequestShowHivePicker += OnRequestShowHivePicker;
            ViewModel.RequestShowRecentBundles += OnRequestShowRecentBundles;

            // Subscribe to remote-instance events (a second invocation forwarded
            // file paths and/or an activate request to us via the named pipe).
            App.RemoteOpenRequested += OnRemoteOpenRequested;
            App.RemoteActivateRequested += OnRemoteActivateRequested;

            // Auto-check for updates on startup
            _ = CheckForUpdatesOnStartupAsync();

            // If launched after an auto-update (or if version changed since last
            // run), show the green "Updated successfully" banner at the top.
            CheckAndShowUpdatedBanner();

            // Load any files supplied as CLI args (e.g. from the shell verb).
            _ = ProcessStartupFilesAsync();

            // Bridge-release one-shot silent migration: if this build was just upgraded
            // from a legacy portable v2.2.1 via the legacy batch-script swap, silently
            // download the matching installer and migrate to %LOCALAPPDATA%\Programs\.
            // Best-effort; no UI on failure, user just keeps the bridge portable.
            _ = TryMigrateToInstallerAsync();
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
            ViewModel.RequestShowReleaseNotes -= OnRequestShowReleaseNotes;
            ViewModel.RequestScrollToNode -= OnRequestScrollToNode;
            ViewModel.RequestShowHivePicker -= OnRequestShowHivePicker;
            ViewModel.RequestShowRecentBundles -= OnRequestShowRecentBundles;

            App.RemoteOpenRequested -= OnRemoteOpenRequested;
            App.RemoteActivateRequested -= OnRemoteActivateRequested;
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

        // ── Release notes window ──────────────────────────────────────────

        private void OnRequestShowReleaseNotes(UpdateInfo info)
        {
            ShowReleaseNotesWindow(info);
        }

        private void ShowReleaseNotesWindow(UpdateInfo info)
        {
            try
            {
                var window = new ReleaseNotesWindow(
                    version: info.LatestVersion,
                    publishedAt: info.PublishedAt,
                    markdownBody: info.ReleaseNotes,
                    githubUrl: string.IsNullOrWhiteSpace(info.ReleaseUrl) ? null : info.ReleaseUrl)
                {
                    Owner = this
                };
                window.ShowDialog();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to open ReleaseNotesWindow: {ex.Message}");
                MessageBox.Show(this,
                    "Could not open the release notes window.\n\n" + ex.Message,
                    "Release Notes", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // ── Shell integration: CLI args, remote-instance handling, registration refresh ──

        /// <summary>
        /// Load any hive file paths that were supplied on the command line
        /// (e.g. when launched from the Explorer "Open with Registry Expert" verb).
        /// </summary>
        private async Task ProcessStartupFilesAsync()
        {
            var files = App.StartupFilePaths;
            if (files == null || files.Count == 0) return;

            foreach (var path in files)
            {
                try
                {
                    await ViewModel.LoadHiveFileAsync(path);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Startup file load failed for '{path}': {ex.Message}");
                    ViewModel.StatusText = $"Failed to load {System.IO.Path.GetFileName(path)}: {ex.Message}";
                }
            }
        }

        /// <summary>
        /// Called by App when a second instance forwards file paths through the named pipe.
        /// Activate ourselves and load each hive sequentially.
        /// </summary>
        private async void OnRemoteOpenRequested(IReadOnlyList<string> paths)
        {
            try
            {
                BringSelfToForeground();
                foreach (var path in paths)
                {
                    try
                    {
                        await ViewModel.LoadHiveFileAsync(path);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Remote-open load failed for '{path}': {ex.Message}");
                        ViewModel.StatusText = $"Failed to load {System.IO.Path.GetFileName(path)}: {ex.Message}";
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"OnRemoteOpenRequested failed: {ex.Message}");
            }
        }

        /// <summary>Called when a second instance only asked us to come to the foreground.</summary>
        private void OnRemoteActivateRequested() => BringSelfToForeground();

        private void BringSelfToForeground()
        {
            try
            {
                if (WindowState == WindowState.Minimized)
                    WindowState = WindowState.Normal;

                Activate();
                Topmost = true;
                Topmost = false;

                var handle = new System.Windows.Interop.WindowInteropHelper(this).Handle;
                if (handle != IntPtr.Zero)
                    SetForegroundWindow(handle);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"BringSelfToForeground failed: {ex.Message}");
            }
        }

        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

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
                _pendingUpdateLocalPath = AutoUpdater.GetDownloadCachePath(info.LatestVersion, info.DownloadKind);
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
        // Dialog used once the update has been downloaded & verified.
        // Reuses the polished ReleaseNotesWindow (full markdown rendering, themed
        // header band) in its install-mode constructor. Footer shows Install +
        // Later buttons; the dialog result drives the next action.
        private void ShowUpdateReadyDialog(UpdateInfo info)
        {
            var window = new ReleaseNotesWindow(
                version: info.LatestVersion,
                publishedAt: info.PublishedAt,
                markdownBody: info.ReleaseNotes,
                githubUrl: string.IsNullOrWhiteSpace(info.ReleaseUrl) ? null : info.ReleaseUrl,
                titleOverride: "Update Ready to Install",
                installButtonText: "Install and update",
                secondaryButtonText: "Later")
            {
                Owner = this
            };

            var result = window.ShowDialog();
            if (result == true)
            {
                ApplyPendingUpdate();
            }
            else
            {
                // User clicked Later or closed via title-bar X -> show the
                // persistent bottom-right reminder for the rest of this session.
                ShowUpdateToastForPendingUpdate(info);
            }
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
                Height = 410,
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

            // Short preview of the release body (first 3 lines or ~240 chars).
            // The full notes are available via the "View full release notes" link below.
            var previewBorder = new Border
            {
                Background = (Brush)FindResource("SurfaceBrush"),
                BorderBrush = (Brush)FindResource("BorderBrush"),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(12, 10, 12, 10),
                Margin = new Thickness(0, 0, 0, 8)
            };
            previewBorder.Child = new TextBlock
            {
                Text = BuildReleaseNotesPreview(info.ReleaseNotes),
                Foreground = (Brush)FindResource("TextPrimaryBrush"),
                FontSize = 12,
                TextWrapping = TextWrapping.Wrap,
                LineHeight = 18,
                MaxHeight = 80,
                TextTrimming = TextTrimming.CharacterEllipsis
            };
            mainPanel.Children.Add(previewBorder);

            // "View full release notes…" link → opens the pretty in-app window.
            var fullLinkText = new TextBlock
            {
                FontSize = 12,
                Margin = new Thickness(0, 0, 0, 12)
            };
            var fullLink = new Hyperlink(new Run("View full release notes \u2192"))
            {
                Foreground = (Brush)FindResource("AccentBrush"),
                TextDecorations = TextDecorations.Underline
            };
            fullLink.Click += (s, e) => ShowReleaseNotesWindow(info);
            fullLinkText.Inlines.Add(fullLink);
            mainPanel.Children.Add(fullLinkText);

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

            // Branch on payload kind: installer (silent in-place upgrade) vs portable
            // (legacy batch-script swap). Installer path is preferred for any release
            // that ships RegistryExpert-Setup.exe; the portable batch path remains for
            // backward compatibility with releases that only ship the single-file exe.
            var info = _pendingUpdateInfo;
            var isInstaller = info?.DownloadKind == DownloadKind.Installer;

            if (isInstaller)
            {
                ApplyPendingUpdateViaInstaller();
            }
            else
            {
                ApplyPendingUpdateViaBatchScript();
            }
        }

        /// <summary>
        /// Installer path: silently launches RegistryExpert-Setup.exe with
        /// /VERYSILENT /CLOSEAPPLICATIONS — the installer waits for our process
        /// to exit, swaps the files, and relaunches the new exe with
        /// --just-updated &lt;version&gt;. No UAC prompt for per-user installs.
        /// </summary>
        private void ApplyPendingUpdateViaInstaller()
        {
            var installingDialog = BuildInstallingDialog();
            var setupExe = _pendingUpdateLocalPath!;
            var fromVersion = UpdateChecker.GetCurrentVersion();

            installingDialog.Loaded += async (s, e) =>
            {
                // Allow the user ~700ms to see the "installing" message
                await Task.Delay(700);

                var ok = AutoUpdater.LaunchInstallerAndExit(setupExe, fromVersion);
                if (!ok)
                {
                    installingDialog.Close();
                    MessageBox.Show(this,
                        "Failed to launch the installer. The update was not applied.",
                        "Update Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // The installer's /CLOSEAPPLICATIONS waits for our process to exit,
                // then writes the new files and runs the [Run] section to relaunch.
                Application.Current.Shutdown();
            };

            installingDialog.ShowDialog();
        }

        /// <summary>
        /// Legacy portable path: writes a batch script to %TEMP% that waits for
        /// our process to exit, then renames the downloaded exe over the running
        /// one and relaunches. Used when the release does not include the installer.
        /// </summary>
        private void ApplyPendingUpdateViaBatchScript()
        {
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

        // ── Update toast (bottom-right) ──────────────────────────────────

        /// <summary>
        /// Show the bottom-right "Install and update" reminder toast. Called from
        /// the "Later" button on the Update Ready dialog and from the on-launch
        /// path when a cached update is still pending.
        /// </summary>
        private void ShowUpdateToastForPendingUpdate(UpdateInfo info)
        {
            ViewModel.PendingUpdateVersion = info.LatestVersion;
            ViewModel.IsUpdateToastVisible = true;
        }

        /// <summary>Toast "Install and update" button: launches the update flow.</summary>
        private void InstallUpdateNow_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.IsUpdateToastVisible = false;
            ApplyPendingUpdate();
        }

        /// <summary>Toast ✕: hides the toast for THIS session only.</summary>
        private void DismissUpdateToast_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.IsUpdateToastVisible = false;
            // Note: don't clear PendingUpdateVersion — keep VM state so we can
            // re-show on next launch if the cached download is still applicable.
        }

        /// <summary>
        /// Build a short, plain-text preview of release notes for use in the update dialogs.
        /// Strips markdown syntax (headings, list markers, bold/italic, code ticks) and
        /// truncates to ~240 chars with an ellipsis. Falls back to an em-dash when empty.
        /// </summary>
        private static string BuildReleaseNotesPreview(string? markdown)
        {
            if (string.IsNullOrWhiteSpace(markdown))
                return "\u2014  (no release notes provided)";

            var sb = new StringBuilder(256);
            var lines = markdown.Replace("\r\n", "\n").Split('\n');

            foreach (var raw in lines)
            {
                var line = raw.TrimEnd();
                if (string.IsNullOrWhiteSpace(line)) continue;

                // Skip fenced code block delimiters entirely
                if (line.TrimStart().StartsWith("```", StringComparison.Ordinal)) continue;

                // Strip leading heading marks
                var t = line.TrimStart();
                while (t.StartsWith("#", StringComparison.Ordinal)) t = t.Substring(1);
                t = t.TrimStart();

                // Strip leading list markers ("- ", "* ", "+ ", "1. ")
                if (t.Length >= 2 && (t[0] == '-' || t[0] == '*' || t[0] == '+') && t[1] == ' ')
                    t = t.Substring(2);
                else
                {
                    int j = 0;
                    while (j < t.Length && char.IsDigit(t[j])) j++;
                    if (j > 0 && j < t.Length - 1 && t[j] == '.' && t[j + 1] == ' ')
                        t = t.Substring(j + 2);
                }

                // Strip inline markdown: backticks, ** and * pairs (best-effort)
                t = t.Replace("`", "").Replace("**", "").Replace("__", "");
                t = t.Replace('*', ' ').Replace('_', ' ');

                if (sb.Length > 0) sb.Append("  \u2022  ");
                sb.Append(t);

                if (sb.Length > 240) break;
            }

            var preview = sb.ToString().Trim();
            if (preview.Length > 240)
                preview = preview.Substring(0, 237).TrimEnd() + "\u2026";

            return preview.Length > 0 ? preview : "\u2014  (no release notes provided)";
        }

        // ── Bridge-release one-shot silent migration ─────────────────────

        /// <summary>
        /// When a legacy v2.2.1 portable user just got upgraded via the legacy
        /// batch-script swap to this bridge release, silently download the matching
        /// installer and let it migrate the user to %LOCALAPPDATA%\Programs\.
        /// One-click migration from the user's perspective: they clicked
        /// "Restart &amp; Update" in the legacy dialog and this finishes the job.
        ///
        /// Trigger conditions (all must be true):
        ///   1. App.UpgradedFromVersion is set (we were launched with --just-updated &lt;v&gt;)
        ///   2. We are NOT already running from the installed location
        ///   3. The release for our current version on GitHub exposes an installer asset
        ///
        /// Best-effort: any failure (network, missing asset, install location detection)
        /// leaves the user on the bridge portable. No error UI; they can continue using it
        /// and will eventually update again via the normal auto-update flow.
        /// </summary>
        private async Task TryMigrateToInstallerAsync()
        {
            // Trigger 1: must have been launched with --just-updated
            if (string.IsNullOrEmpty(App.UpgradedFromVersion)) return;

            // Trigger 2: must NOT already be running from the installer location
            var currentExe = Environment.ProcessPath
                ?? System.Reflection.Assembly.GetEntryAssembly()?.Location
                ?? "";
            if (string.IsNullOrEmpty(currentExe)) return;

            var installedExe = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Programs", "RegistryExpert", "RegistryExpert.exe");

            if (string.Equals(currentExe, installedExe, StringComparison.OrdinalIgnoreCase))
                return; // already installed; nothing to migrate

            try
            {
                // Trigger 3: installer for our current version must be downloadable
                var currentVersion = UpdateChecker.GetCurrentVersion();
                var info = await UpdateChecker.GetReleaseByTagAsync("v" + currentVersion);
                if (info == null
                    || info.DownloadKind != DownloadKind.Installer
                    || string.IsNullOrEmpty(info.DownloadUrl))
                {
                    System.Diagnostics.Debug.WriteLine(
                        "Migration skipped: installer asset not available for current version.");
                    return;
                }

                // Show a small modal while we silently install. Cannot be closed by the
                // user; will close when this process exits via Application.Shutdown().
                var dlg = BuildInstallingDialog(
                    title: "Setting up Registry Expert",
                    heading: "Setting up Registry Expert\u2026",
                    subtitle: "Moving to the installed location. The app will restart in a moment.");

                dlg.Loaded += async (s, e) =>
                {
                    try
                    {
                        // Reuse cached download if we have it; otherwise fetch from the release.
                        string? localPath;
                        if (AutoUpdater.IsUpdateAlreadyDownloaded(info))
                        {
                            localPath = AutoUpdater.GetDownloadCachePath(info.LatestVersion, info.DownloadKind);
                        }
                        else
                        {
                            localPath = await AutoUpdater.DownloadUpdateAsync(info);
                        }

                        if (string.IsNullOrEmpty(localPath))
                        {
                            System.Diagnostics.Debug.WriteLine(
                                "Migration aborted: installer download failed; continuing as portable.");
                            dlg.Close();
                            return;
                        }

                        // Hand off to the installer; this exits our process.
                        // Inno Setup's [Run] section will relaunch the new exe from the
                        // installed location with --just-updated <prev> so the post-update
                        // banner fires there.
                        var ok = AutoUpdater.LaunchInstallerAndExit(
                            setupExePath: localPath,
                            currentVersionForArg: App.UpgradedFromVersion!);
                        if (ok)
                        {
                            Application.Current.Shutdown();
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine(
                                "Migration aborted: installer launch failed; continuing as portable.");
                            dlg.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"TryMigrateToInstallerAsync (download/launch) failed: {ex.Message}");
                        dlg.Close();
                    }
                };

                dlg.ShowDialog();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"TryMigrateToInstallerAsync failed: {ex.Message}");
                // Swallow — best-effort migration, user keeps the bridge portable.
            }
        }

        // Compact modal with indeterminate progress bar. No close button,
        // no chrome buttons -- displayed briefly before exit.
        // Default heading is "Installing update..."; callers can override for
        // contexts like the legacy-portable -> installer silent migration.
        private Window BuildInstallingDialog(
            string title = "Installing Update",
            string heading = "Installing update\u2026",
            string subtitle = "Registry Expert will restart automatically in a moment.")
        {
            var dialog = new Window
            {
                Title = title,
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
                Text = heading,
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
                Text = subtitle,
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

        private async void UpdatedBannerReleaseNotesBtn_Click(object sender, RoutedEventArgs e)
        {
            // Prefer the in-app Release Notes window. If we cannot fetch the release
            // data (offline, API failure), fall back to opening the GitHub URL in
            // the browser so the user still gets to read the notes.
            try
            {
                var version = UpdateChecker.GetCurrentVersion();
                var tag = "v" + version;

                UpdatedBannerReleaseNotesBtn.IsEnabled = false;
                try
                {
                    var info = await UpdateChecker.GetReleaseByTagAsync(tag);
                    if (info != null)
                    {
                        ShowReleaseNotesWindow(info);
                        return;
                    }
                }
                finally
                {
                    UpdatedBannerReleaseNotesBtn.IsEnabled = true;
                }

                // Fallback — open the release URL in the user's browser.
                var url = BuildGitHubReleaseTagUrl(version);
                Process.Start(new ProcessStartInfo { FileName = url, UseShellExecute = true });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Release notes link failed: {ex.Message}");
            }
        }

        /// <summary>
        /// Build a GitHub releases-tag URL that respects REGEXPERT_UPDATE_URL when set,
        /// so a private/test repo can be targeted without rebuilding.
        /// </summary>
        private static string BuildGitHubReleaseTagUrl(string version)
        {
            string owner = "bowenzhang85";
            string repo = "RegistryExpert";

            if (UpdateChecker.IsUsingOverrideUrl)
            {
                var apiUrl = UpdateChecker.GitHubApiUrl;
                const string marker = "/repos/";
                var ownerStart = apiUrl.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
                if (ownerStart >= 0)
                {
                    var rest = apiUrl.Substring(ownerStart + marker.Length);
                    var slash = rest.IndexOf('/');
                    if (slash > 0)
                    {
                        owner = rest.Substring(0, slash);
                        var afterOwner = rest.Substring(slash + 1);
                        var nextSlash = afterOwner.IndexOf('/');
                        repo = nextSlash > 0 ? afterOwner.Substring(0, nextSlash) : afterOwner;
                    }
                }
            }

            return $"https://github.com/{owner}/{repo}/releases/tag/v{version}";
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
                        localPath = AutoUpdater.GetDownloadCachePath(info.LatestVersion, info.DownloadKind);
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

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using RegistryParser.Abstractions;

namespace RegistryExpert
{
    public class TimelineForm : Form
    {
        private OfflineRegistryParser _parser;
        private readonly IReadOnlyList<(OfflineRegistryParser Parser, string HiveTypeName)> _parsers;
        private readonly MainForm _mainForm;
        
        private DataGridView _timelineGrid = null!;
        private ComboBox _filterCombo = null!;
        private ComboBox _limitCombo = null!;
        private ComboBox? _hiveSelector;
        private TextBox _searchBox = null!;
        private System.Windows.Forms.Timer _searchDebounce = null!;
        private bool _isUpdatingPlaceholder;
        private const string SearchPlaceholder = "Search key path...";
        private DateTimePicker _fromDatePicker = null!;
        private DateTimePicker _toDatePicker = null!;
        private FlowLayoutPanel _customDatePanel = null!;
        private Button _refreshButton = null!;
        private Button _exportButton = null!;
        private Panel _statusPanel = null!;
        private string _statusText = "Click 'Scan' to analyze registry key timestamps";
        private Color _statusForeColor;
        private ProgressBar _progressBar = null!;
        private List<Panel> _separators = new();
        private Panel _filterPanel = null!;
        private Panel _searchRow = null!;
        
        private List<TimelineEntry> _allEntries = new();
        private List<TimelineEntry> _filteredEntries = new();
        private CancellationTokenSource? _scanCts;
        private bool _isScanning;
        
        // Pagination state
        private int _currentDisplayCount = 0;
        private int _pageSize = 100;
        private Panel _loadMorePanel = null!;
        private Button _loadMoreButton = null!;
        
        // Transaction log analysis state
        private Button _browseLogsButton = null!;
        private CheckBox _txLogFilterCheckbox = null!;
        private List<string> _detectedLogPaths = new();
        private List<string> _manualLogPaths = new();
        private List<TransactionLogDiff> _txLogDiffs = new();
        private CancellationTokenSource? _analyzeLogsCts;

        // Transaction log detail panel
        private SplitContainer _mainSplitter = null!;
        private Panel _detailPanel = null!;
        private DataGridView _detailGrid = null!;
        private Label _detailHeader = null!;

        public TimelineForm(IReadOnlyList<(OfflineRegistryParser Parser, string HiveTypeName)> parsers, MainForm mainForm)
        {
            _parsers = parsers;
            _parser = parsers[0].Parser;
            _mainForm = mainForm;
            InitializeComponent();
            this.Icon = mainForm.Icon;
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            ModernTheme.ApplyWindowStyle(this);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _scanCts?.Cancel();
                _scanCts?.Dispose();
                _scanCts = null;
                _analyzeLogsCts?.Cancel();
                _analyzeLogsCts?.Dispose();
                _analyzeLogsCts = null;
                _searchDebounce?.Stop();
                _searchDebounce?.Dispose();

                // Clear large data collections to release references before GC
                _allEntries.Clear();
                _filteredEntries.Clear();
                _txLogDiffs.Clear();
                _timelineGrid.Rows.Clear();
            }
            base.Dispose(disposing);
        }

        /// <summary>
        /// Handle DPI changes when moving between monitors with different DPI settings.
        /// </summary>
        protected override void OnDpiChanged(DpiChangedEventArgs e)
        {
            // Reset the cached DPI scale factor so it gets recalculated
            DpiHelper.ResetScaleFactor();
            
            base.OnDpiChanged(e);
            
            // Reapply theme and custom timeline padding overrides
            ModernTheme.ApplyTo(_timelineGrid);
            _timelineGrid.DefaultCellStyle.Padding = DpiHelper.ScalePadding(8, 4, 8, 4);
            _timelineGrid.ColumnHeadersDefaultCellStyle.Padding = DpiHelper.ScalePadding(8, 4, 8, 4);

            // Rescale button sizes for new DPI
            foreach (var btn in new[] { _refreshButton, _exportButton })
            {
                btn.Size = DpiHelper.ScaleSize(90, 30);
                btn.Margin = DpiHelper.ScalePadding(2);
            }
            foreach (var btn in new[] { _browseLogsButton })
            {
                btn.Size = DpiHelper.ScaleSize(110, 30);
                btn.Margin = DpiHelper.ScalePadding(2);
            }
            
            // Rescale separators
            foreach (var sep in _separators)
            {
                sep.Width = DpiHelper.Scale(1);
                sep.Height = DpiHelper.Scale(16);
                sep.Margin = DpiHelper.ScalePadding(4, 7, 4, 7);
            }
            
            // Rescale status panel
            _statusPanel.Height = DpiHelper.Scale(28);
            _statusPanel.Padding = DpiHelper.ScalePadding(10, 0, 10, 0);

            // Rescale filter panel and search row
            _filterPanel.Height = DpiHelper.Scale(100);
            _searchRow.Height = DpiHelper.Scale(32);
        }

        private void InitializeComponent()
        {
            this.Text = "Timeline View - Registry Keys by Last Modified";
            this.Size = new Size(1000, 700);
            this.StartPosition = FormStartPosition.CenterParent;
            this.MinimumSize = new Size(800, 500);
            this.AutoScaleMode = AutoScaleMode.Dpi;
            ModernTheme.ApplyTo(this);

            // Top filter panel
            _filterPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = DpiHelper.Scale(100),
                BackColor = ModernTheme.Surface,
                Padding = new Padding(12, 12, 12, 0)
            };

            // First row - use FlowLayoutPanel for proper DPI scaling
            var filterFlow = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = DpiHelper.Scale(40),
                BackColor = ModernTheme.Surface,
                WrapContents = false,
                AutoSize = false,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };

            var filterLabel = new Label
            {
                Text = "Time Filter:",
                ForeColor = ModernTheme.TextSecondary,
                Font = ModernTheme.RegularFont,
                AutoSize = true,
                Margin = new Padding(0, 8, 5, 0)
            };

            _filterCombo = new ComboBox
            {
                Size = DpiHelper.ScaleSize(150, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = ModernTheme.RegularFont,
                Margin = new Padding(0, 2, 10, 0)
            };
            _filterCombo.Items.AddRange(new object[] { 
                "All Times", 
                "Last Hour", 
                "Last 24 Hours", 
                "Last 7 Days", 
                "Last 30 Days",
                "Custom Range..." 
            });
            _filterCombo.SelectedIndex = 0;
            _filterCombo.SelectedIndexChanged += FilterCombo_SelectedIndexChanged;
            _filterCombo.BackColor = ModernTheme.Surface;
            _filterCombo.ForeColor = ModernTheme.TextPrimary;

            var limitLabel = new Label
            {
                Text = "Show:",
                ForeColor = ModernTheme.TextSecondary,
                Font = ModernTheme.RegularFont,
                AutoSize = true,
                Margin = new Padding(0, 8, 5, 0)
            };

            _limitCombo = new ComboBox
            {
                Size = DpiHelper.ScaleSize(80, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = ModernTheme.RegularFont,
                Margin = new Padding(0, 2, 5, 0)
            };
            _limitCombo.Items.AddRange(new object[] { "50", "100", "500", "1000", "All" });
            _limitCombo.SelectedIndex = 1; // Default to 100
            _limitCombo.SelectedIndexChanged += LimitCombo_SelectedIndexChanged;
            _limitCombo.BackColor = ModernTheme.Surface;
            _limitCombo.ForeColor = ModernTheme.TextPrimary;

            var keysLabel = new Label
            {
                Text = "keys",
                ForeColor = ModernTheme.TextSecondary,
                Font = ModernTheme.RegularFont,
                AutoSize = true,
                Margin = new Padding(0, 8, 0, 0)
            };

            // Create pill-shaped action buttons (matching MainForm toolbar style)
            _refreshButton = CreatePillButton("Scan", RefreshButton_Click);
            _exportButton = CreatePillButton("Export", ExportButton_Click);
            _exportButton.Enabled = false;

            // Search box with placeholder
            _searchBox = new TextBox
            {
                Size = DpiHelper.ScaleSize(200, 28),
                Font = ModernTheme.RegularFont,
                Text = SearchPlaceholder,
                ForeColor = ModernTheme.TextSecondary,
                BackColor = ModernTheme.Surface,
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(0, 2, 0, 0),
                AccessibleName = "Search key path"
            };
            _searchBox.GotFocus += (s, e) =>
            {
                if (_searchBox.Text == SearchPlaceholder)
                {
                    _isUpdatingPlaceholder = true;
                    _searchBox.Text = "";
                    _searchBox.ForeColor = ModernTheme.TextPrimary;
                    _isUpdatingPlaceholder = false;
                }
            };
            _searchBox.LostFocus += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(_searchBox.Text))
                {
                    _isUpdatingPlaceholder = true;
                    _searchBox.Text = SearchPlaceholder;
                    _searchBox.ForeColor = ModernTheme.TextSecondary;
                    _isUpdatingPlaceholder = false;
                }
            };

            // Debounce timer for search (300ms)
            _searchDebounce = new System.Windows.Forms.Timer { Interval = 300 };
            _searchDebounce.Tick += (s, e) =>
            {
                _searchDebounce.Stop();
                ApplyFilter();
            };
            _searchBox.TextChanged += (s, e) =>
            {
                // Don't trigger search on placeholder text changes
                if (_isUpdatingPlaceholder) return;
                if (_searchBox.Text == SearchPlaceholder) return;
                _searchDebounce.Stop();
                _searchDebounce.Start();
            };

            // Hive selector (only visible when multiple hives are loaded)
            if (_parsers.Count > 1)
            {
                var hiveSelectorLabel = new Label
                {
                    Text = "Hive:",
                    ForeColor = ModernTheme.TextSecondary,
                    Font = ModernTheme.RegularFont,
                    AutoSize = true,
                    Margin = new Padding(0, 8, 5, 0)
                };
                _hiveSelector = new ComboBox
                {
                    Size = DpiHelper.ScaleSize(130, 28),
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    Font = ModernTheme.RegularFont,
                    Margin = new Padding(0, 2, 10, 0),
                    BackColor = ModernTheme.Surface,
                    ForeColor = ModernTheme.TextPrimary
                };
                foreach (var (_, name) in _parsers)
                    _hiveSelector.Items.Add(name);
                _hiveSelector.SelectedIndex = 0;
                _hiveSelector.SelectedIndexChanged += (s, e) =>
                {
                    if (_hiveSelector.SelectedIndex >= 0 && _hiveSelector.SelectedIndex < _parsers.Count)
                    {
                        // Cancel any in-progress scan before switching parsers
                        _scanCts?.Cancel();

                        _parser = _parsers[_hiveSelector.SelectedIndex].Parser;
                        // Clear existing data so next scan uses the new parser
                        _allEntries.Clear();
                        _filteredEntries.Clear();
                        _timelineGrid.Rows.Clear();
                        _currentDisplayCount = 0;

                        UpdateStatus($"Selected {_parsers[_hiveSelector.SelectedIndex].HiveTypeName} hive. Press Scan to load timeline.");
                        _exportButton.Enabled = false;
                    }
                };
                filterFlow.Controls.AddRange(new Control[] { hiveSelectorLabel, _hiveSelector });
            }

            // Build filter flow: combos, separators, buttons
            _separators.Clear();
            var sep1 = CreateFilterSeparator();
            var sep2 = CreateFilterSeparator();

            // Transaction log toolbar controls
            var sep3 = CreateFilterSeparator();

            _browseLogsButton = CreatePillButton("Browse Logs...", BrowseLogsButton_Click);
            _browseLogsButton.Size = DpiHelper.ScaleSize(110, 30);

            _txLogFilterCheckbox = new CheckBox
            {
                Text = "TxLog only",
                ForeColor = ModernTheme.TextSecondary,
                Font = ModernTheme.RegularFont,
                AutoSize = true,
                Margin = new Padding(8, 8, 0, 0),
                Visible = false  // Only shown after analysis
            };
            _txLogFilterCheckbox.CheckedChanged += (s, e) => ApplyFilter();

            filterFlow.Controls.AddRange(new Control[] {
                filterLabel, _filterCombo, limitLabel, _limitCombo, keysLabel,
                sep1, _refreshButton, sep2, _exportButton,
                sep3, _browseLogsButton, _txLogFilterCheckbox
            });

            // Search row - full-width search box on second row
            _searchRow = new Panel
            {
                Dock = DockStyle.Top,
                Height = DpiHelper.Scale(32),
                BackColor = ModernTheme.Surface,
                Padding = new Padding(0, 2, 0, 4)
            };
            _searchBox.Dock = DockStyle.Fill;
            _searchBox.Size = default; // Clear fixed size; Dock.Fill handles sizing
            _searchBox.Margin = new Padding(0);
            _searchRow.Controls.Add(_searchBox);

            // Custom date range panel (hidden by default) - also use FlowLayoutPanel
            _customDatePanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = DpiHelper.Scale(36),
                Visible = false,
                BackColor = ModernTheme.Surface,
                WrapContents = false,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };

            var fromLabel = new Label
            {
                Text = "From:",
                ForeColor = ModernTheme.TextSecondary,
                Font = ModernTheme.RegularFont,
                AutoSize = true,
                Margin = new Padding(0, 8, 5, 0)
            };

            _fromDatePicker = new DateTimePicker
            {
                Size = DpiHelper.ScaleSize(180, 28),
                Format = DateTimePickerFormat.Custom,
                CustomFormat = "yyyy-MM-dd HH:mm",
                Font = ModernTheme.RegularFont,
                Margin = new Padding(0, 2, 15, 0)
            };
            _fromDatePicker.Value = DateTime.Now.AddDays(-7);
            _fromDatePicker.ValueChanged += (s, e) => ApplyFilter();

            var toLabel = new Label
            {
                Text = "To:",
                ForeColor = ModernTheme.TextSecondary,
                Font = ModernTheme.RegularFont,
                AutoSize = true,
                Margin = new Padding(0, 8, 5, 0)
            };

            _toDatePicker = new DateTimePicker
            {
                Size = DpiHelper.ScaleSize(180, 28),
                Format = DateTimePickerFormat.Custom,
                CustomFormat = "yyyy-MM-dd HH:mm",
                Font = ModernTheme.RegularFont,
                Margin = new Padding(0, 2, 0, 0)
            };
            _toDatePicker.Value = DateTime.Now;
            _toDatePicker.ValueChanged += (s, e) => ApplyFilter();

            _customDatePanel.Controls.AddRange(new Control[] { fromLabel, _fromDatePicker, toLabel, _toDatePicker });

            // Bottom separator line for filter panel
            var filterSeparator = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 1,
                BackColor = ModernTheme.Border
            };
            filterSeparator.Tag = "filterBottomSep";

            // Add flow panels to filter panel (order matters - bottom to top for Dock.Top)
            _filterPanel.Controls.Add(filterSeparator);
            _filterPanel.Controls.Add(_customDatePanel);
            _filterPanel.Controls.Add(_searchRow);
            _filterPanel.Controls.Add(filterFlow);

            // Status bar (matching MainForm: 28px, owner-drawn)
            _statusPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = DpiHelper.Scale(28),
                BackColor = ModernTheme.Surface,
                Padding = DpiHelper.ScalePadding(10, 0, 10, 0)
            };
            _statusForeColor = ModernTheme.TextSecondary;

            // Owner-draw status text (matches MainForm pattern)
            typeof(Panel).InvokeMember("DoubleBuffered",
                System.Reflection.BindingFlags.SetProperty | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
                null, _statusPanel, new object[] { true });
            _statusPanel.Paint += StatusPanel_Paint;

            _progressBar = new ProgressBar
            {
                Dock = DockStyle.Bottom,
                Height = DpiHelper.Scale(3),
                Style = ProgressBarStyle.Marquee,
                Visible = false
            };

            _statusPanel.Controls.Add(_progressBar);

            // Timeline grid
            _timelineGrid = new DataGridView { Dock = DockStyle.Fill, AccessibleName = "Timeline Results" };
            _timelineGrid.MultiSelect = false;
            ModernTheme.ApplyTo(_timelineGrid);
            _timelineGrid.BackgroundColor = ModernTheme.Background;  // Override: use Background instead of Surface
            _timelineGrid.DefaultCellStyle.Padding = DpiHelper.ScalePadding(8, 4, 8, 4);  // Override default padding
            _timelineGrid.ColumnHeadersDefaultCellStyle.Padding = DpiHelper.ScalePadding(8, 4, 8, 4);

            // Add columns
            _timelineGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "LastModified",
                HeaderText = "Last Modified",
                FillWeight = 25,
                DefaultCellStyle = { Format = "yyyy-MM-dd HH:mm:ss" }
            });
            _timelineGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "KeyPath",
                HeaderText = "Key Path",
                FillWeight = 75
            });
            _timelineGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "ChangeType",
                HeaderText = "Change Type",
                FillWeight = 15,
                Visible = false  // Hidden until analysis is run
            });
            _timelineGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Details",
                HeaderText = "Details",
                FillWeight = 20,
                Visible = false
            });

            // Enable sorting
            foreach (DataGridViewColumn col in _timelineGrid.Columns)
            {
                col.SortMode = DataGridViewColumnSortMode.Automatic;
            }

            _timelineGrid.CellDoubleClick += TimelineGrid_CellDoubleClick;
            _timelineGrid.KeyDown += TimelineGrid_KeyDown;
            _timelineGrid.SelectionChanged += TimelineGrid_SelectionChanged;

            // Context menu
            var contextMenu = new ContextMenuStrip();
            contextMenu.Renderer = new ToolStripProfessionalRenderer(new ModernColorTable());
            
            var navigateItem = new ToolStripMenuItem("Navigate to Key", null, (s, e) => NavigateToSelectedKey());
            navigateItem.Font = ModernTheme.RegularFont;
            
            var copyPathItem = new ToolStripMenuItem("Copy Path", null, (s, e) => CopySelectedPath());
            copyPathItem.ShortcutKeys = Keys.Control | Keys.C;
            copyPathItem.Font = ModernTheme.RegularFont;
            
            contextMenu.Items.Add(navigateItem);
            contextMenu.Items.Add(copyPathItem);
            
            contextMenu.Opening += (s, e) =>
            {
                contextMenu.BackColor = ModernTheme.Surface;
                contextMenu.ForeColor = ModernTheme.TextPrimary;
            };
            
            _timelineGrid.ContextMenuStrip = contextMenu;

            // Load More panel (between grid and status)
            _loadMorePanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = DpiHelper.Scale(50),
                BackColor = ModernTheme.Background,
                Visible = false,
                Padding = new Padding(0, 8, 0, 8)
            };

            _loadMoreButton = new Button
            {
                Text = "Load More",
                FlatStyle = FlatStyle.Flat,
                BackColor = ModernTheme.Surface,
                ForeColor = ModernTheme.TextPrimary,
                Font = ModernTheme.RegularFont,
                Size = DpiHelper.ScaleSize(250, 34),
                Cursor = Cursors.Hand
            };
            _loadMoreButton.FlatAppearance.BorderColor = ModernTheme.Border;
            _loadMoreButton.FlatAppearance.MouseOverBackColor = ModernTheme.Selection;
            _loadMoreButton.Click += LoadMore_Click;

            // Center the button in the panel
            _loadMorePanel.Resize += (s, e) =>
            {
                _loadMoreButton.Location = new Point(
                    (_loadMorePanel.Width - _loadMoreButton.Width) / 2,
                    (_loadMorePanel.Height - _loadMoreButton.Height) / 2
                );
            };
            _loadMorePanel.Controls.Add(_loadMoreButton);

            // Detail panel for transaction log value-level changes
            _detailPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ModernTheme.Surface,
                Visible = true
            };

            _detailHeader = new Label
            {
                Dock = DockStyle.Top,
                Height = DpiHelper.Scale(28),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                ForeColor = ModernTheme.TextPrimary,
                BackColor = ModernTheme.Surface,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = DpiHelper.ScalePadding(8, 0, 0, 0),
                Text = "Transaction Log Details"
            };

            _detailGrid = new DataGridView { Dock = DockStyle.Fill, AccessibleName = "Transaction Log Value Changes" };
            _detailGrid.MultiSelect = false;
            ModernTheme.ApplyTo(_detailGrid);
            _detailGrid.BackgroundColor = ModernTheme.Background;
            _detailGrid.DefaultCellStyle.Padding = DpiHelper.ScalePadding(6, 3, 6, 3);

            _detailGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "DetailValueName", HeaderText = "Value Name", FillWeight = 20 });
            _detailGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "DetailValueType", HeaderText = "Type", FillWeight = 10 });
            _detailGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "DetailChangeType", HeaderText = "Change", FillWeight = 10 });
            _detailGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "DetailOldData", HeaderText = "Before (Current Hive)", FillWeight = 30 });
            _detailGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "DetailNewData", HeaderText = "After (With Logs)", FillWeight = 30 });

            var detailSep = new Panel { Dock = DockStyle.Top, Height = 1, BackColor = ModernTheme.Border };

            _detailPanel.Controls.Add(_detailGrid);
            _detailPanel.Controls.Add(detailSep);
            _detailPanel.Controls.Add(_detailHeader);

            // SplitContainer: top = timeline grid + load-more, bottom = detail panel
            // NOTE: SplitterDistance is NOT set here — when Panel2 is later
            // uncollapsed, the distance is set proportionally to the actual height.
            _mainSplitter = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                Panel2Collapsed = true,  // Detail hidden by default
                SplitterWidth = DpiHelper.Scale(4),
                BackColor = ModernTheme.Border
            };

            _mainSplitter.Panel1.Controls.Add(_timelineGrid);
            _mainSplitter.Panel1.Controls.Add(_loadMorePanel);
            _mainSplitter.Panel2.Controls.Add(_detailPanel);

            // Add controls (order matters for docking)
            this.Controls.Add(_mainSplitter);
            this.Controls.Add(_filterPanel);
            this.Controls.Add(_statusPanel);

            // Handle form closing
            this.FormClosing += (s, e) =>
            {
                _scanCts?.Cancel();
            };

            // Auto-scan on load
            this.Shown += async (s, e) => await ScanRegistryAsync();
        }

        private void FilterCombo_SelectedIndexChanged(object? sender, EventArgs e)
        {
            // Show/hide custom date panel
            _customDatePanel.Visible = _filterCombo.SelectedItem?.ToString() == "Custom Range...";
            ApplyFilter();
        }

        private async void RefreshButton_Click(object? sender, EventArgs e)
        {
            if (_isScanning)
            {
                _scanCts?.Cancel();
                return;
            }
            await ScanRegistryAsync().ConfigureAwait(true);
        }

        private async Task ScanRegistryAsync()
        {
            if (_isScanning) return;

            _isScanning = true;
            _scanCts?.Cancel();
            _scanCts?.Dispose();
            _scanCts = new CancellationTokenSource();
            var token = _scanCts.Token;

            _refreshButton.Text = "Stop";
            _progressBar.Visible = true;
            UpdateStatus("Scanning registry keys...");
            _exportButton.Enabled = false;
            _timelineGrid.Rows.Clear();
            _allEntries.Clear();

            try
            {
                var rootKey = _parser.GetRootKey();
                if (rootKey == null)
                {
                    UpdateStatus("No registry hive loaded");
                    return;
                }

                var entries = new List<TimelineEntry>();
                int scannedCount = 0;

                await Task.Run(() =>
                {
                    ScanKeyRecursive(rootKey, entries, ref scannedCount, token);
                }, token).ConfigureAwait(true);

                if (token.IsCancellationRequested)
                {
                    UpdateStatus($"Scan cancelled - found {entries.Count:N0} keys");
                }
                else
                {
                    _allEntries = entries.OrderByDescending(e => e.LastModified)
                        .ThenBy(e => e.DisplayPath, StringComparer.OrdinalIgnoreCase).ToList();
                    UpdateStatus($"Scanned {scannedCount:N0} keys - {_allEntries.Count:N0} have timestamps");
                }

                // Auto-detect transaction log files
                _detectedLogPaths = TransactionLogAnalyzer.DetectLogFiles(_parser.FilePath ?? "");
                var logPaths = _manualLogPaths.Count > 0 ? _manualLogPaths : _detectedLogPaths;

                // Auto-analyze transaction logs if available
                if (logPaths.Count > 0 && !string.IsNullOrEmpty(_parser.FilePath) && !token.IsCancellationRequested)
                {
                    try
                    {
                        UpdateStatus("Analyzing transaction logs...");
                        _txLogDiffs = await TransactionLogAnalyzer.AnalyzeAsync(
                            _parser.FilePath,
                            logPaths,
                            msg => BeginInvoke(() => UpdateStatus(msg)),
                            token
                        );

                        MergeTxLogDiffs();

                        // Show the TxLog columns if changes were found
                        if (_txLogDiffs.Count > 0)
                        {
                            _timelineGrid.Columns["ChangeType"]!.Visible = true;
                            _timelineGrid.Columns["Details"]!.Visible = true;
                            _txLogFilterCheckbox.Visible = true;
                        }
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Transaction log analysis failed: {ex.Message}");
                        // Continue without transaction log data — normal timeline still displays
                    }
                }

                ApplyFilter();
                _exportButton.Enabled = _filteredEntries.Count > 0;
            }
            catch (OperationCanceledException)
            {
                UpdateStatus($"Scan cancelled - found {_allEntries.Count:N0} keys");
            }
            catch (Exception ex)
            {
                UpdateStatus($"Error: {ex.Message}");
                _statusForeColor = ModernTheme.Error;
                _statusPanel.Invalidate();
            }
            finally
            {
                _isScanning = false;
                _refreshButton.Text = "Scan";
                _progressBar.Visible = false;
            }
        }

        private void ScanKeyRecursive(RegistryKey key, List<TimelineEntry> entries, ref int scannedCount, CancellationToken token)
        {
            if (token.IsCancellationRequested) return;

            scannedCount++;

            // Update status periodically
            if (scannedCount % 1000 == 0)
            {
                var count = scannedCount; // Capture for lambda
                if (!this.IsDisposed && this.IsHandleCreated)
                {
                    this.BeginInvoke(() =>
                    {
                        if (!this.IsDisposed)
                            UpdateStatus($"Scanning... {count:N0} keys processed");
                    });
                }
            }

            if (key.LastWriteTime.HasValue)
            {
                var displayPath = ConvertRootPath(key.KeyPath);
                entries.Add(new TimelineEntry
                {
                    LastModified = key.LastWriteTime.Value.DateTime,
                    KeyPath = key.KeyPath,
                    DisplayPath = displayPath
                });
            }

            if (key.SubKeys != null)
            {
                foreach (var subKey in key.SubKeys)
                {
                    if (token.IsCancellationRequested) break;
                    ScanKeyRecursive(subKey, entries, ref scannedCount, token);
                }
            }
        }

        private void LimitCombo_SelectedIndexChanged(object? sender, EventArgs e)
        {
            var limitText = _limitCombo.SelectedItem?.ToString() ?? "100";
            if (limitText == "All")
            {
                _pageSize = _filteredEntries.Count > 0 ? _filteredEntries.Count : int.MaxValue / 2; // Avoid overflow when adding to _currentDisplayCount
            }
            else if (int.TryParse(limitText, out int limit))
            {
                _pageSize = limit;
            }
            ApplyFilter();
        }

        private void ApplyFilter()
        {
            if (_allEntries.Count == 0)
            {
                _filteredEntries.Clear();
                _currentDisplayCount = 0;
                UpdateGridDisplay();
                return;
            }

            // Apply time filter
            DateTime? fromDate = null;
            DateTime? toDate = DateTime.Now;

            switch (_filterCombo.SelectedIndex)
            {
                case 0: // All Times
                    fromDate = null;
                    toDate = null;
                    break;
                case 1: // Last Hour
                    fromDate = DateTime.Now.AddHours(-1);
                    break;
                case 2: // Last 24 Hours
                    fromDate = DateTime.Now.AddHours(-24);
                    break;
                case 3: // Last 7 Days
                    fromDate = DateTime.Now.AddDays(-7);
                    break;
                case 4: // Last 30 Days
                    fromDate = DateTime.Now.AddDays(-30);
                    break;
                case 5: // Custom Range
                    fromDate = _fromDatePicker.Value;
                    toDate = _toDatePicker.Value;
                    // Ensure fromDate <= toDate by swapping if necessary
                    if (fromDate.HasValue && toDate.HasValue && fromDate.Value > toDate.Value)
                    {
                        (fromDate, toDate) = (toDate, fromDate);
                    }
                    break;
            }

            var filtered = _allEntries.AsEnumerable();

            if (fromDate.HasValue)
                filtered = filtered.Where(e => e.LastModified >= fromDate.Value);
            if (toDate.HasValue)
                filtered = filtered.Where(e => e.LastModified <= toDate.Value);

            // Apply search filter (case-insensitive)
            var searchText = _searchBox.Text.Trim();
            if (!string.IsNullOrEmpty(searchText) && searchText != SearchPlaceholder)
                filtered = filtered.Where(e => e.DisplayPath.Contains(searchText, StringComparison.OrdinalIgnoreCase));

            // Apply TxLog-only filter
            if (_txLogFilterCheckbox != null && _txLogFilterCheckbox.Visible && _txLogFilterCheckbox.Checked)
                filtered = filtered.Where(e => e.TxLogDiff != null);

            _filteredEntries = filtered.ToList();

            // Reset pagination - show first page
            _currentDisplayCount = Math.Min(_pageSize, _filteredEntries.Count);

            // When "All" is selected, show all entries
            if (_limitCombo.SelectedItem?.ToString() == "All")
                _currentDisplayCount = _filteredEntries.Count;
            
            UpdateGridDisplay();
            _exportButton.Enabled = _filteredEntries.Count > 0;
        }

        private void UpdateGridDisplay()
        {
            _timelineGrid.SuspendLayout();
            try
            {
                _timelineGrid.Rows.Clear();

                // Display only up to _currentDisplayCount entries
                foreach (var entry in _filteredEntries.Take(_currentDisplayCount))
                {
                    var rowIndex = _timelineGrid.Rows.Add(
                        entry.LastModified,
                        entry.DisplayPath,
                        entry.TxLogDiff != null ? GetStatusText(entry.TxLogDiff.ChangeType) : "",
                        entry.TxLogDiff?.ChangeSummary ?? ""
                    );
                    // Store the entry reference in the row's Tag for reliable retrieval after sorting
                    _timelineGrid.Rows[rowIndex].Tag = entry;

                    // Highlight rows with transaction log changes
                    ApplyChangeTypeColor(_timelineGrid.Rows[rowIndex], entry);
                }
            }
            finally
            {
                _timelineGrid.ResumeLayout();
            }

            // Update Load More button visibility and text
            int remaining = _filteredEntries.Count - _currentDisplayCount;
            if (remaining > 0 && _pageSize != int.MaxValue)
            {
                _loadMorePanel.Visible = true;
                int nextBatch = Math.Min(_pageSize, remaining);
                _loadMoreButton.Text = $"Load More ({remaining:N0} remaining)";
            }
            else
            {
                _loadMorePanel.Visible = false;
            }

            // Update status
            UpdateStatusLabel();
        }

        private void UpdateStatusLabel()
        {
            var filterText = _filterCombo.SelectedItem?.ToString() ?? "All";
            var txLogSuffix = _txLogDiffs.Count > 0 ? $" | {_txLogDiffs.Count:N0} transaction log changes" : "";
            
            if (_filteredEntries.Count == 0)
            {
                UpdateStatus($"No keys found ({filterText}){txLogSuffix}");
            }
            else if (_currentDisplayCount >= _filteredEntries.Count)
            {
                UpdateStatus($"Showing all {_filteredEntries.Count:N0} of {_allEntries.Count:N0} keys ({filterText}){txLogSuffix}");
            }
            else
            {
                UpdateStatus($"Showing {_currentDisplayCount:N0} of {_filteredEntries.Count:N0} keys ({filterText}){txLogSuffix}");
            }
        }

        private void LoadMore_Click(object? sender, EventArgs e)
        {
            // Add another page of results
            int previousCount = _currentDisplayCount;
            int remaining = _filteredEntries.Count - _currentDisplayCount;
            int toAdd = Math.Min(_pageSize, remaining);
            _currentDisplayCount = _currentDisplayCount + toAdd;
            
            // Append new rows to the grid (more efficient than rebuilding)
            var newEntries = _filteredEntries.Skip(previousCount).Take(_currentDisplayCount - previousCount);
            foreach (var entry in newEntries)
            {
                var rowIndex = _timelineGrid.Rows.Add(
                    entry.LastModified,
                    entry.DisplayPath,
                    entry.TxLogDiff != null ? GetStatusText(entry.TxLogDiff.ChangeType) : "",
                    entry.TxLogDiff?.ChangeSummary ?? ""
                );
                // Store the entry reference in the row's Tag for reliable retrieval after sorting
                _timelineGrid.Rows[rowIndex].Tag = entry;

                // Highlight rows with transaction log changes
                ApplyChangeTypeColor(_timelineGrid.Rows[rowIndex], entry);
            }

            // Update Load More button visibility and text
            remaining = _filteredEntries.Count - _currentDisplayCount;
            if (remaining > 0)
            {
                _loadMoreButton.Text = $"Load More ({remaining:N0} remaining)";
            }
            else
            {
                _loadMorePanel.Visible = false;
            }

            // Update status
            UpdateStatusLabel();
            
            // Scroll to show some of the new content
            if (_timelineGrid.Rows.Count > 0)
            {
                _timelineGrid.FirstDisplayedScrollingRowIndex = previousCount;
            }
        }

        private void TimelineGrid_CellDoubleClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                NavigateToSelectedKey();
            }
        }

        private void TimelineGrid_KeyDown(object? sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                NavigateToSelectedKey();
                e.Handled = true;
            }
            else if (e.Control && e.KeyCode == Keys.C)
            {
                CopySelectedPath();
                e.Handled = true;
            }
        }

        private void NavigateToSelectedKey()
        {
            if (_timelineGrid.SelectedRows.Count > 0)
            {
                var selectedRow = _timelineGrid.SelectedRows[0];
                if (selectedRow.Tag is TimelineEntry entry)
                {
                    _mainForm.NavigateToKey(entry.KeyPath);
                    _mainForm.BringToFront();
                    _mainForm.Activate();
                }
            }
        }

        private void CopySelectedPath()
        {
            if (_timelineGrid.SelectedRows.Count > 0)
            {
                var selectedRow = _timelineGrid.SelectedRows[0];
                if (selectedRow.Tag is TimelineEntry entry)
                {
                    try
                    {
                        Clipboard.SetText(entry.DisplayPath);
                        UpdateStatus("Path copied to clipboard");
                    }
                    catch (System.Runtime.InteropServices.ExternalException)
                    {
                        UpdateStatus("Failed to copy - clipboard may be in use");
                    }
                }
            }
        }

        private void ExportButton_Click(object? sender, EventArgs e)
        {
            if (_filteredEntries.Count == 0)
            {
                MessageBox.Show("No data to export.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using var dialog = new SaveFileDialog
            {
                Title = "Export Timeline to CSV",
                Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*",
                DefaultExt = ".csv",
                FileName = $"registry_timeline_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
            };

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    // Use 64KB buffer for better I/O performance with large exports
                    using var writer = new StreamWriter(dialog.FileName, false, System.Text.Encoding.UTF8, 65536);
                    
                    // Write header
                    var hasTxLog = _timelineGrid.Columns["ChangeType"]?.Visible == true;
                    writer.Write("\"Last Modified\",\"Key Path\"");
                    if (hasTxLog)
                        writer.Write(",\"Change Type\",\"Details\"");
                    writer.WriteLine();
                    
                    // Write data
                    foreach (var entry in _filteredEntries)
                    {
                        var path = EscapeCsv(entry.DisplayPath);
                        writer.Write($"\"{entry.LastModified:yyyy-MM-dd HH:mm:ss}\",\"{path}\"");
                        if (hasTxLog)
                        {
                            var diff = entry.TxLogDiff;
                            var status = diff != null ? GetStatusText(diff.ChangeType) : "";
                            var summary = diff?.ChangeSummary ?? "";
                            writer.Write($",\"{EscapeCsv(status)}\",\"{EscapeCsv(summary)}\"");
                        }
                        writer.WriteLine();
                    }

                    UpdateStatus($"Exported {_filteredEntries.Count:N0} entries to {Path.GetFileName(dialog.FileName)}");
                    _statusForeColor = ModernTheme.Success;
                    _statusPanel.Invalidate();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Export failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UpdateStatus("Export failed");
                    _statusForeColor = ModernTheme.Error;
                    _statusPanel.Invalidate();
                }
            }
        }

        private static string GetStatusText(TransactionLogChangeType changeType) => changeType switch
        {
            TransactionLogChangeType.KeyAdded => "New",
            TransactionLogChangeType.KeyDeleted => "Deleted",
            TransactionLogChangeType.ValuesChanged => "Modified",
            _ => ""
        };

        private void ApplyChangeTypeColor(DataGridViewRow row, TimelineEntry entry)
        {
            if (entry.TxLogDiff != null)
            {
                row.Cells["ChangeType"].Style.ForeColor = entry.TxLogDiff.ChangeType switch
                {
                    TransactionLogChangeType.KeyAdded => ModernTheme.DiffAdded,
                    TransactionLogChangeType.KeyDeleted => ModernTheme.DiffRemoved,
                    _ => ModernTheme.Warning
                };
            }
        }

        /// <summary>
        /// Formats a registry value type for display. Named enum values (e.g. "RegSz") pass through;
        /// raw numeric strings (e.g. "18") are displayed as "Unknown (18)".
        /// </summary>
        private static string FormatValueType(string valueType)
        {
            if (string.IsNullOrEmpty(valueType)) return "";
            return int.TryParse(valueType, out _) ? $"Unknown ({valueType})" : valueType;
        }

        private static string EscapeCsv(string value) => value.Replace("\"", "\"\"");

        private void TimelineGrid_SelectionChanged(object? sender, EventArgs e)
        {
            if (_timelineGrid.SelectedRows.Count == 0 || _txLogDiffs.Count == 0)
            {
                if (!_mainSplitter.Panel2Collapsed)
                    _mainSplitter.Panel2Collapsed = true;
                return;
            }

            var row = _timelineGrid.SelectedRows[0];
            if (row.Tag is not TimelineEntry entry || entry.TxLogDiff == null)
            {
                if (!_mainSplitter.Panel2Collapsed)
                    _mainSplitter.Panel2Collapsed = true;
                return;
            }

            UpdateDetailPanel(entry.TxLogDiff);
        }

        private void UpdateDetailPanel(TransactionLogDiff diff)
        {
            _detailGrid.Rows.Clear();

            var statusText = GetStatusText(diff.ChangeType);
            var tsInfo = "";
            if (diff.OldTimestamp.HasValue && diff.NewTimestamp.HasValue)
                tsInfo = $"  |  {diff.OldTimestamp:yyyy-MM-dd HH:mm:ss}  -->  {diff.NewTimestamp:yyyy-MM-dd HH:mm:ss}";
            else if (diff.NewTimestamp.HasValue)
                tsInfo = $"  |  (new) {diff.NewTimestamp:yyyy-MM-dd HH:mm:ss}";

            _detailHeader.Text = $"Transaction Log Details: {diff.DisplayPath}  [{statusText}]{tsInfo}";

            foreach (var vc in diff.ValueChanges)
            {
                var changeText = vc.ChangeType switch
                {
                    ValueChangeType.Added => "Added",
                    ValueChangeType.Removed => "Removed",
                    ValueChangeType.Modified => "Modified",
                    _ => ""
                };
                var rowIdx = _detailGrid.Rows.Add(
                    vc.ValueName,
                    FormatValueType(vc.ValueType),
                    changeText,
                    vc.OldData ?? "(not present)",
                    vc.NewData ?? "(not present)"
                );

                // Color-code the change type
                var changeCell = _detailGrid.Rows[rowIdx].Cells["DetailChangeType"];
                changeCell.Style.ForeColor = vc.ChangeType switch
                {
                    ValueChangeType.Added => ModernTheme.DiffAdded,
                    ValueChangeType.Removed => ModernTheme.DiffRemoved,
                    ValueChangeType.Modified => ModernTheme.Warning,
                    _ => ModernTheme.TextPrimary
                };
            }

            if (diff.ValueChanges.Count == 0 && diff.OldTimestamp.HasValue && diff.NewTimestamp.HasValue)
            {
                // Only timestamp changed - show that
                _detailGrid.Rows.Add("(key timestamp)", "", "Modified",
                    diff.OldTimestamp.Value.ToString("yyyy-MM-dd HH:mm:ss"),
                    diff.NewTimestamp.Value.ToString("yyyy-MM-dd HH:mm:ss"));
            }

            // Show the detail panel
            if (_mainSplitter.Panel2Collapsed)
            {
                _mainSplitter.Panel2Collapsed = false;
                try { _mainSplitter.SplitterDistance = (int)(_mainSplitter.Height * 0.6); }
                catch (InvalidOperationException) { }
            }
        }

        private void MergeTxLogDiffs()
        {
            // Build a lookup from existing entries by key path
            var entryByPath = new Dictionary<string, TimelineEntry>(StringComparer.OrdinalIgnoreCase);
            foreach (var entry in _allEntries)
            {
                entryByPath.TryAdd(entry.KeyPath, entry);
            }

            foreach (var diff in _txLogDiffs)
            {
                if (entryByPath.TryGetValue(diff.KeyPath, out var existing))
                {
                    // Attach diff to existing entry
                    existing.TxLogDiff = diff;
                }
                else
                {
                    // New key from transaction logs - add as a new entry
                    var newEntry = new TimelineEntry
                    {
                        LastModified = diff.NewTimestamp ?? diff.OldTimestamp ?? DateTime.MinValue,
                        KeyPath = diff.KeyPath,
                        DisplayPath = ConvertRootPath(diff.KeyPath),
                        TxLogDiff = diff
                    };
                    _allEntries.Add(newEntry);
                }
            }

            // Re-sort
            _allEntries = _allEntries.OrderByDescending(e => e.LastModified)
                .ThenBy(e => e.DisplayPath, StringComparer.OrdinalIgnoreCase).ToList();
        }

        private async void BrowseLogsButton_Click(object? sender, EventArgs e)
        {
            using var ofd = new OpenFileDialog
            {
                Title = "Select Transaction Log Files",
                Filter = "Registry Log Files (*.LOG1;*.LOG2)|*.LOG1;*.LOG2|All Files (*.*)|*.*",
                Multiselect = true
            };

            // Start in the hive's directory if available
            var hivePath = _parser.FilePath;
            if (!string.IsNullOrEmpty(hivePath))
                ofd.InitialDirectory = System.IO.Path.GetDirectoryName(hivePath) ?? "";

            if (ofd.ShowDialog(this) != DialogResult.OK) return;

            _manualLogPaths = new List<string>(ofd.FileNames);

            // If no scan has been done yet, just store the paths for the next scan
            if (_allEntries.Count == 0)
            {
                UpdateStatus($"Selected {_manualLogPaths.Count} log file(s). Click 'Scan' to load timeline with log analysis.");
                return;
            }

            // Run analysis immediately on the already-scanned data
            if (string.IsNullOrEmpty(hivePath)) return;

            _analyzeLogsCts?.Cancel();
            _analyzeLogsCts?.Dispose();
            _analyzeLogsCts = new CancellationTokenSource();
            var token = _analyzeLogsCts.Token;

            _progressBar.Visible = true;

            try
            {
                _txLogDiffs = await TransactionLogAnalyzer.AnalyzeAsync(
                    hivePath,
                    _manualLogPaths,
                    msg => BeginInvoke(() => UpdateStatus(msg)),
                    token
                );

                MergeTxLogDiffs();

                if (_txLogDiffs.Count > 0)
                {
                    _timelineGrid.Columns["ChangeType"]!.Visible = true;
                    _timelineGrid.Columns["Details"]!.Visible = true;
                    _txLogFilterCheckbox.Visible = true;
                }

                ApplyFilter();
                UpdateStatus($"Transaction log analysis complete: {_txLogDiffs.Count:N0} changes found across {_manualLogPaths.Count} log file(s).");
            }
            catch (OperationCanceledException)
            {
                UpdateStatus("Transaction log analysis cancelled.");
            }
            catch (Exception ex)
            {
                UpdateStatus($"Transaction log analysis error: {ex.Message}");
                _statusForeColor = ModernTheme.Error;
                _statusPanel.Invalidate();
            }
            finally
            {
                _progressBar.Visible = false;
            }
        }

        private string ConvertRootPath(string path) => _parser.ConvertRootPath(path);

        /// <summary>
        /// Create a pill-shaped button matching MainForm toolbar style.
        /// Transparent background with owner-drawn rounded-rect hover/press effects.
        /// </summary>
        private Button CreatePillButton(string text, EventHandler onClick)
        {
            var btn = new Button
            {
                Text = text,
                AccessibleName = text,
                AccessibleRole = AccessibleRole.PushButton,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.Transparent,
                ForeColor = Color.Transparent,  // Hide framework-drawn text; we draw our own
                Font = ModernTheme.RegularFont,
                Size = DpiHelper.ScaleSize(90, 30),
                Margin = DpiHelper.ScalePadding(2),
                Cursor = Cursors.Hand
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.FlatAppearance.MouseOverBackColor = Color.Transparent;
            btn.FlatAppearance.MouseDownBackColor = Color.Transparent;
            btn.Click += onClick;

            bool isHovered = false;
            bool isPressed = false;
            btn.MouseEnter += (s, e) => { isHovered = true; btn.Invalidate(); };
            btn.MouseLeave += (s, e) => { isHovered = false; isPressed = false; btn.Invalidate(); };
            btn.MouseDown += (s, e) => { isPressed = true; btn.Invalidate(); };
            btn.MouseUp += (s, e) => { isPressed = false; btn.Invalidate(); };

            btn.Paint += (s, e) =>
            {
                var g = e.Graphics;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                // Clear background to prevent ghosting when text changes
                using (var bgBrush = new SolidBrush(btn.Parent?.BackColor ?? ModernTheme.Surface))
                    g.FillRectangle(bgBrush, btn.ClientRectangle);

                // Draw pill-shaped hover/press background
                if (isHovered || isPressed)
                {
                    var hoverColor = isPressed ? ModernTheme.AccentDark : ModernTheme.SurfaceHover;
                    using var hoverBrush = new SolidBrush(hoverColor);
                    var rect = new Rectangle(1, 1, btn.Width - 2, btn.Height - 2);
                    var radius = DpiHelper.Scale(6);
                    using var path = new System.Drawing.Drawing2D.GraphicsPath();
                    int d = radius * 2;
                    path.AddArc(rect.X, rect.Y, d, d, 180, 90);
                    path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
                    path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
                    path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
                    path.CloseFigure();
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    g.FillPath(hoverBrush, path);
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.Default;
                }

                // Draw centered text using btn.Text so it tracks dynamic changes
                var displayText = btn.Text;
                using var textBrush = new SolidBrush(isPressed ? Color.White : ModernTheme.TextPrimary);
                var textSize = g.MeasureString(displayText, ModernTheme.RegularFont);
                var textX = (btn.Width - textSize.Width) / 2;
                var textY = (btn.Height - textSize.Height) / 2;
                g.DrawString(displayText, ModernTheme.RegularFont, textBrush, textX, textY);
            };

            return btn;
        }

        /// <summary>
        /// Create a vertical separator panel for the filter flow, matching MainForm toolbar separators.
        /// </summary>
        private Panel CreateFilterSeparator()
        {
            var sep = new Panel
            {
                Width = DpiHelper.Scale(1),
                Height = DpiHelper.Scale(16),
                Margin = DpiHelper.ScalePadding(4, 7, 4, 7),
                BackColor = ModernTheme.Border
            };
            _separators.Add(sep);
            return sep;
        }

        /// <summary>
        /// Owner-draw the status panel text (matches MainForm pattern).
        /// </summary>
        private void StatusPanel_Paint(object? sender, PaintEventArgs e)
        {
            var padding = _statusPanel.Padding;
            var textRect = new Rectangle(
                padding.Left,
                padding.Top,
                _statusPanel.ClientSize.Width - padding.Left - padding.Right,
                _statusPanel.ClientSize.Height - padding.Top - padding.Bottom);

            TextRenderer.DrawText(
                e.Graphics,
                _statusText,
                ModernTheme.RegularFont,
                textRect,
                _statusForeColor,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPadding);
        }

        /// <summary>
        /// Update the status bar text and reset foreground color to default.
        /// </summary>
        private void UpdateStatus(string text)
        {
            _statusText = text;
            _statusForeColor = ModernTheme.TextSecondary;
            _statusPanel.Invalidate();
        }

        /// <summary>
        /// Refresh theme colors
        /// </summary>
        public void RefreshTheme()
        {
            ModernTheme.ApplyWindowStyle(this);
            this.BackColor = ModernTheme.Background;
            
            // Refresh grid
            _timelineGrid.BackgroundColor = ModernTheme.Background;
            _timelineGrid.DefaultCellStyle.BackColor = ModernTheme.Surface;
            _timelineGrid.DefaultCellStyle.ForeColor = ModernTheme.TextPrimary;
            _timelineGrid.DefaultCellStyle.SelectionBackColor = ModernTheme.Selection;
            _timelineGrid.AlternatingRowsDefaultCellStyle.BackColor = ModernTheme.ListViewAltRow;
            _timelineGrid.ColumnHeadersDefaultCellStyle.BackColor = ModernTheme.TreeViewBack;
            _timelineGrid.ColumnHeadersDefaultCellStyle.ForeColor = ModernTheme.TextSecondary;
            _timelineGrid.ColumnHeadersDefaultCellStyle.SelectionBackColor = ModernTheme.TreeViewBack;
            _timelineGrid.ColumnHeadersDefaultCellStyle.SelectionForeColor = ModernTheme.TextSecondary;
            _timelineGrid.ColumnHeadersDefaultCellStyle.Font = ModernTheme.DataBoldFont;
            _timelineGrid.GridColor = ModernTheme.Border;

            // Refresh Load More panel and button
            _loadMorePanel.BackColor = ModernTheme.Background;
            _loadMoreButton.BackColor = ModernTheme.Surface;
            _loadMoreButton.ForeColor = ModernTheme.TextPrimary;
            _loadMoreButton.FlatAppearance.BorderColor = ModernTheme.Border;
            _loadMoreButton.FlatAppearance.MouseOverBackColor = ModernTheme.Selection;

            // Refresh other controls
            RefreshControlTheme(this);

            // Refresh search box
            _searchBox.BackColor = ModernTheme.Surface;
            _searchBox.ForeColor = _searchBox.Text == SearchPlaceholder ? ModernTheme.TextSecondary : ModernTheme.TextPrimary;

            // Refresh pill buttons (transparent forecolor hides framework text; we owner-draw)
            foreach (var btn in new[] { _refreshButton, _exportButton, _browseLogsButton })
            {
                btn.BackColor = Color.Transparent;
                btn.ForeColor = Color.Transparent;
                btn.Invalidate();
            }

            // Refresh separators
            foreach (var sep in _separators)
            {
                sep.BackColor = ModernTheme.Border;
            }

            // Refresh transaction log detail panel
            if (_detailPanel != null)
            {
                _detailPanel.BackColor = ModernTheme.Surface;
                _detailHeader.ForeColor = ModernTheme.TextPrimary;
                _detailHeader.BackColor = ModernTheme.Surface;
                ModernTheme.ApplyTo(_detailGrid);
                _detailGrid.BackgroundColor = ModernTheme.Background;
            }
            if (_mainSplitter != null)
            {
                _mainSplitter.BackColor = ModernTheme.Border;
            }
            if (_txLogFilterCheckbox != null)
            {
                _txLogFilterCheckbox.ForeColor = ModernTheme.TextSecondary;
            }

            // Refresh status bar
            _statusPanel.BackColor = ModernTheme.Surface;
            _statusForeColor = ModernTheme.TextSecondary;
            _statusPanel.Invalidate();

            this.Refresh();
        }

        private static void RefreshControlTheme(Control parent)
        {
            foreach (Control ctrl in parent.Controls)
            {
                switch (ctrl)
                {
                    case DataGridView:
                        // Handled separately in RefreshTheme
                        break;
                    case TextBox:
                        // Search box handled separately
                        break;
                    case Panel panel:
                        panel.BackColor = ModernTheme.Surface;
                        RefreshControlTheme(panel);
                        break;
                    case Label label:
                        label.ForeColor = ModernTheme.TextSecondary;
                        break;
                    case ComboBox combo:
                        combo.BackColor = ModernTheme.Surface;
                        combo.ForeColor = ModernTheme.TextPrimary;
                        break;
                    case DateTimePicker picker:
                        picker.BackColor = ModernTheme.Surface;
                        picker.ForeColor = ModernTheme.TextPrimary;
                        break;
                    case Button:
                        // Buttons handled separately in RefreshTheme
                        break;
                    default:
                        if (ctrl.HasChildren)
                            RefreshControlTheme(ctrl);
                        break;
                }
            }
        }

        private class TimelineEntry
        {
            public DateTime LastModified { get; set; }
            public string KeyPath { get; set; } = "";
            public string DisplayPath { get; set; } = "";
            // Transaction log analysis results (null if no analysis done)
            public TransactionLogDiff? TxLogDiff { get; set; }
        }

    }
}

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
        
        // Collapse/Expand state
        private Button _collapseButton = null!;
        private bool _isCollapsed;
        private List<CollapsedGroup> _collapsedGroups = new();
        private HashSet<int> _expandedGroupIndices = new();
        private int _currentGroupDisplayCount;

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
                _searchDebounce?.Stop();
                _searchDebounce?.Dispose();

                // Clear large data collections to release references before GC
                _allEntries.Clear();
                _filteredEntries.Clear();
                _collapsedGroups.Clear();
                _expandedGroupIndices.Clear();
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
            foreach (var btn in new[] { _refreshButton, _exportButton, _collapseButton })
            {
                btn.Size = DpiHelper.ScaleSize(90, 30);
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
            _collapseButton = CreatePillButton("Collapse View", CollapseButton_Click);
            _collapseButton.Size = DpiHelper.ScaleSize(110, 30);
            _collapseButton.Enabled = false;

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
                    _searchBox.Text = "";
                    _searchBox.ForeColor = ModernTheme.TextPrimary;
                }
            };
            _searchBox.LostFocus += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(_searchBox.Text))
                {
                    _searchBox.Text = SearchPlaceholder;
                    _searchBox.ForeColor = ModernTheme.TextSecondary;
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
                // Don't trigger search on placeholder text
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
                        _collapsedGroups.Clear();
                        _expandedGroupIndices.Clear();
                        _timelineGrid.Rows.Clear();
                        _currentDisplayCount = 0;

                        // Reset collapse state so UI is consistent for the new hive
                        if (_isCollapsed)
                        {
                            _isCollapsed = false;
                            _collapseButton.Text = "Collapse View";
                            foreach (DataGridViewColumn col in _timelineGrid.Columns)
                                col.SortMode = DataGridViewColumnSortMode.Automatic;
                        }

                        UpdateStatus($"Selected {_parsers[_hiveSelector.SelectedIndex].HiveTypeName} hive. Press Scan to load timeline.");
                        _exportButton.Enabled = false;
                        _collapseButton.Enabled = false;
                    }
                };
                filterFlow.Controls.AddRange(new Control[] { hiveSelectorLabel, _hiveSelector });
            }

            // Build filter flow: combos, separators, buttons
            _separators.Clear();
            var sep1 = CreateFilterSeparator();
            var sep2 = CreateFilterSeparator();
            var sep3 = CreateFilterSeparator();

            filterFlow.Controls.AddRange(new Control[] {
                filterLabel, _filterCombo, limitLabel, _limitCombo, keysLabel,
                sep1, _refreshButton, sep2, _exportButton, sep3, _collapseButton
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

            // Enable sorting
            foreach (DataGridViewColumn col in _timelineGrid.Columns)
            {
                col.SortMode = DataGridViewColumnSortMode.Automatic;
            }

            _timelineGrid.CellDoubleClick += TimelineGrid_CellDoubleClick;
            _timelineGrid.CellClick += TimelineGrid_CellClick;
            _timelineGrid.KeyDown += TimelineGrid_KeyDown;

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

            // Add controls (order matters for docking)
            this.Controls.Add(_timelineGrid);
            this.Controls.Add(_loadMorePanel);
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
            _collapseButton.Enabled = false;
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

                ApplyFilter();
                _exportButton.Enabled = _filteredEntries.Count > 0;
                _collapseButton.Enabled = _filteredEntries.Count > 0;
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

            _filteredEntries = filtered.ToList();

            // Rebuild collapsed groups if in collapsed mode
            if (_isCollapsed)
            {
                BuildCollapsedGroups();
                _expandedGroupIndices.Clear();
                _currentGroupDisplayCount = _limitCombo.SelectedItem?.ToString() == "All" 
                    ? _collapsedGroups.Count 
                    : _pageSize;
            }

            // Reset pagination - show first page
            _currentDisplayCount = Math.Min(_pageSize, _filteredEntries.Count);

            // When "All" is selected, show all entries
            if (_limitCombo.SelectedItem?.ToString() == "All")
                _currentDisplayCount = _filteredEntries.Count;
            
            UpdateGridDisplay();
            _exportButton.Enabled = _filteredEntries.Count > 0;
            _collapseButton.Enabled = _filteredEntries.Count > 0;
        }

        private void UpdateGridDisplay()
        {
            if (_isCollapsed)
            {
                UpdateCollapsedGridDisplay();
                return;
            }

            _timelineGrid.SuspendLayout();
            try
            {
                _timelineGrid.Rows.Clear();

                // Display only up to _currentDisplayCount entries
                foreach (var entry in _filteredEntries.Take(_currentDisplayCount))
                {
                    var rowIndex = _timelineGrid.Rows.Add(
                        entry.LastModified,
                        entry.DisplayPath
                    );
                    // Store the entry reference in the row's Tag for reliable retrieval after sorting
                    _timelineGrid.Rows[rowIndex].Tag = entry;
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

        private void CollapseButton_Click(object? sender, EventArgs e)
        {
            _isCollapsed = !_isCollapsed;
            _collapseButton.Text = _isCollapsed ? "Expand View" : "Collapse View";
            
            // Disable column sorting in collapsed mode (would break group ordering)
            foreach (DataGridViewColumn col in _timelineGrid.Columns)
                col.SortMode = _isCollapsed ? DataGridViewColumnSortMode.NotSortable : DataGridViewColumnSortMode.Automatic;
            
            if (_isCollapsed)
            {
                BuildCollapsedGroups();
                _expandedGroupIndices.Clear();
                _currentGroupDisplayCount = _pageSize;
            }
            else
            {
                _collapsedGroups.Clear();
                _expandedGroupIndices.Clear();
            }
            
            UpdateGridDisplay();
        }

        private void BuildCollapsedGroups()
        {
            _collapsedGroups.Clear();
            
            if (_filteredEntries.Count == 0) return;
            
            // Sort by timestamp (truncated to seconds) descending, then by path to ensure
            // related entries are adjacent for grouping. Sub-second precision is not displayed
            // and would cause interleaving of paths that appear to share the same timestamp.
            _filteredEntries = _filteredEntries
                .OrderByDescending(e => e.LastModified.Ticks / TimeSpan.TicksPerSecond)
                .ThenBy(e => e.DisplayPath, StringComparer.OrdinalIgnoreCase)
                .ToList();
            
            // Minimum number of path segments required for a common ancestor to be considered "related"
            // e.g., 3 means SYSTEM\ControlSet001\Control is the minimum grouping depth
            const int minAncestorSegments = 3;
            
            CollapsedGroup? currentGroup = null;
            bool groupEstablished = false; // True after first merge sets the actual group root
            
            foreach (var entry in _filteredEntries)
            {
                if (currentGroup == null)
                {
                    currentGroup = new CollapsedGroup { RootPath = entry.DisplayPath };
                    currentGroup.Entries.Add(entry);
                    groupEstablished = false;
                }
                else
                {
                    // Find the common ancestor between this entry and the group's root
                    var commonAncestor = GetCommonAncestor(entry.DisplayPath, currentGroup.RootPath);
                    int ancestorSegments = string.IsNullOrEmpty(commonAncestor) ? 0 : commonAncestor.Split('\\').Length;
                    
                    if (ancestorSegments < minAncestorSegments)
                    {
                        // Not related enough - close current group and start a new one
                        _collapsedGroups.Add(currentGroup);
                        currentGroup = new CollapsedGroup { RootPath = entry.DisplayPath };
                        currentGroup.Entries.Add(entry);
                        groupEstablished = false;
                    }
                    else if (!groupEstablished)
                    {
                        // First merge - establish the group's deepest common root
                        currentGroup.Entries.Add(entry);
                        currentGroup.RootPath = commonAncestor;
                        groupEstablished = true;
                    }
                    else
                    {
                        // Group established - only merge if ancestor doesn't shrink the root
                        int currentRootSegments = currentGroup.RootPath.Split('\\').Length;
                        if (ancestorSegments >= currentRootSegments)
                        {
                            currentGroup.Entries.Add(entry);
                        }
                        else
                        {
                            // Would shrink the root - start new group
                            _collapsedGroups.Add(currentGroup);
                            currentGroup = new CollapsedGroup { RootPath = entry.DisplayPath };
                            currentGroup.Entries.Add(entry);
                            groupEstablished = false;
                        }
                    }
                }
            }
            
            if (currentGroup != null)
                _collapsedGroups.Add(currentGroup);
        }

        private static string GetCommonAncestor(string path1, string path2)
        {
            var segments1 = path1.Split('\\');
            var segments2 = path2.Split('\\');
            int minLen = Math.Min(segments1.Length, segments2.Length);
            
            int commonCount = 0;
            for (int i = 0; i < minLen; i++)
            {
                if (segments1[i].Equals(segments2[i], StringComparison.OrdinalIgnoreCase))
                    commonCount++;
                else
                    break;
            }
            
            if (commonCount == 0) return "";
            return string.Join("\\", segments1.Take(commonCount));
        }

        private void UpdateCollapsedGridDisplay()
        {
            _timelineGrid.SuspendLayout();
            try
            {
                _timelineGrid.Rows.Clear();
                
                int groupsToShow = Math.Min(_currentGroupDisplayCount, _collapsedGroups.Count);
                
                for (int i = 0; i < groupsToShow; i++)
                {
                    var group = _collapsedGroups[i];
                    bool isExpanded = _expandedGroupIndices.Contains(i);
                    
                    if (group.Entries.Count == 1)
                    {
                        // Single entry - show as normal row
                        var entry = group.Entries[0];
                        var rowIndex = _timelineGrid.Rows.Add(
                            entry.LastModified,
                            entry.DisplayPath
                        );
                        _timelineGrid.Rows[rowIndex].Tag = entry;
                    }
                    else
                    {
                        // Multi-entry group - show collapsed summary or expanded
                        var minTime = group.Entries.Min(e => e.LastModified);
                        var maxTime = group.Entries.Max(e => e.LastModified);
                        
                        // Format time for display - single timestamp if all entries share the same second
                        string timeDisplay;
                        bool sameSecond = minTime.Year == maxTime.Year && minTime.Month == maxTime.Month
                            && minTime.Day == maxTime.Day && minTime.Hour == maxTime.Hour
                            && minTime.Minute == maxTime.Minute && minTime.Second == maxTime.Second;
                        if (sameSecond)
                            timeDisplay = $"{minTime:yyyy-MM-dd HH:mm:ss}";
                        else if (minTime.Date == maxTime.Date)
                            timeDisplay = $"{minTime:yyyy-MM-dd HH:mm:ss} - {maxTime:HH:mm:ss}";
                        else
                            timeDisplay = $"{minTime:yyyy-MM-dd HH:mm:ss} - {maxTime:yyyy-MM-dd HH:mm:ss}";
                        
                        string indicator = isExpanded ? "[-]" : "[+]";
                        string pathDisplay = $"{indicator} {group.RootPath} ({group.Entries.Count} keys)";
                        
                        var groupRowIndex = _timelineGrid.Rows.Add(
                            timeDisplay,
                            pathDisplay
                        );
                        _timelineGrid.Rows[groupRowIndex].Tag = group;
                        _timelineGrid.Rows[groupRowIndex].DefaultCellStyle.ForeColor = ModernTheme.Accent;
                        _timelineGrid.Rows[groupRowIndex].DefaultCellStyle.SelectionForeColor = ModernTheme.Accent;
                        
                        if (isExpanded)
                        {
                            // Show individual entries indented below the group header
                            foreach (var entry in group.Entries.OrderByDescending(e => e.LastModified))
                            {
                                var childRowIndex = _timelineGrid.Rows.Add(
                                    entry.LastModified,
                                    $"    {entry.DisplayPath}"
                                );
                                _timelineGrid.Rows[childRowIndex].Tag = entry;
                            }
                        }
                    }
                }
            }
            finally
            {
                _timelineGrid.ResumeLayout();
            }
            
            // Update Load More button visibility for collapsed mode
            int remainingGroups = _collapsedGroups.Count - Math.Min(_currentGroupDisplayCount, _collapsedGroups.Count);
            if (remainingGroups > 0)
            {
                _loadMorePanel.Visible = true;
                _loadMoreButton.Text = $"Load More ({remainingGroups:N0} groups remaining)";
            }
            else
            {
                _loadMorePanel.Visible = false;
            }
            
            // Update status
            int shownGroups = Math.Min(_currentGroupDisplayCount, _collapsedGroups.Count);
            int totalEntries = _collapsedGroups.Take(shownGroups).Sum(g => g.Entries.Count);
            int allEntries = _collapsedGroups.Sum(g => g.Entries.Count);
            int groupCount = _collapsedGroups.Take(shownGroups).Count(g => g.Entries.Count > 1);
            var filterText = _filterCombo.SelectedItem?.ToString() ?? "All";
            if (shownGroups < _collapsedGroups.Count)
                UpdateStatus($"Showing {shownGroups:N0} of {_collapsedGroups.Count:N0} rows ({groupCount:N0} groups from {totalEntries:N0} of {allEntries:N0} keys) ({filterText})");
            else
                UpdateStatus($"Collapsed: {_collapsedGroups.Count:N0} rows ({groupCount:N0} groups from {totalEntries:N0} keys) ({filterText})");
        }

        private void TimelineGrid_CellClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (!_isCollapsed || e.RowIndex < 0) return;
            
            var row = _timelineGrid.Rows[e.RowIndex];
            if (row.Tag is CollapsedGroup group)
            {
                int groupIndex = FindGroupIndex(e.RowIndex);
                if (groupIndex < 0) return;
                
                _timelineGrid.SuspendLayout();
                try
                {
                    if (_expandedGroupIndices.Contains(groupIndex))
                    {
                        // Collapse: remove child rows below this header
                        _expandedGroupIndices.Remove(groupIndex);
                        int childCount = group.Entries.Count;
                        for (int i = 0; i < childCount; i++)
                            _timelineGrid.Rows.RemoveAt(e.RowIndex + 1);
                        
                        // Update header text to [+]
                        row.Cells[1].Value = $"[+] {group.RootPath} ({group.Entries.Count} keys)";
                    }
                    else
                    {
                        // Expand: insert child rows below this header
                        _expandedGroupIndices.Add(groupIndex);
                        int insertAt = e.RowIndex + 1;
                        foreach (var entry in group.Entries.OrderByDescending(en => en.LastModified))
                        {
                            var newRow = new DataGridViewRow();
                            newRow.CreateCells(_timelineGrid, entry.LastModified, $"    {entry.DisplayPath}");
                            newRow.Tag = entry;
                            _timelineGrid.Rows.Insert(insertAt, newRow);
                            insertAt++;
                        }
                        
                        // Update header text to [-]
                        row.Cells[1].Value = $"[-] {group.RootPath} ({group.Entries.Count} keys)";
                    }
                }
                finally
                {
                    _timelineGrid.ResumeLayout();
                }
                
                // Update status to reflect expand/collapse change
                int shownGroups = Math.Min(_currentGroupDisplayCount, _collapsedGroups.Count);
                int totalEntries = _collapsedGroups.Take(shownGroups).Sum(g => g.Entries.Count);
                int allEntries = _collapsedGroups.Sum(g => g.Entries.Count);
                int groupCount = _collapsedGroups.Take(shownGroups).Count(g => g.Entries.Count > 1);
                var filterText = _filterCombo.SelectedItem?.ToString() ?? "All";
                if (shownGroups < _collapsedGroups.Count)
                    UpdateStatus($"Showing {shownGroups:N0} of {_collapsedGroups.Count:N0} rows ({groupCount:N0} groups from {totalEntries:N0} of {allEntries:N0} keys) ({filterText})");
                else
                    UpdateStatus($"Collapsed: {_collapsedGroups.Count:N0} rows ({groupCount:N0} groups from {totalEntries:N0} keys) ({filterText})");
            }
        }

        private int FindGroupIndex(int rowIndex)
        {
            // Walk through displayed groups to find which group the clicked row belongs to
            int currentRow = 0;
            int groupsToShow = Math.Min(_currentGroupDisplayCount, _collapsedGroups.Count);
            for (int i = 0; i < groupsToShow; i++)
            {
                var group = _collapsedGroups[i];
                if (currentRow == rowIndex)
                    return i;
                
                currentRow++; // The group/entry header row
                
                if (group.Entries.Count > 1 && _expandedGroupIndices.Contains(i))
                    currentRow += group.Entries.Count; // Expanded child rows
            }
            return -1;
        }

        private void UpdateStatusLabel()
        {
            var filterText = _filterCombo.SelectedItem?.ToString() ?? "All";
            
            if (_filteredEntries.Count == 0)
            {
                UpdateStatus($"No keys found ({filterText})");
            }
            else if (_currentDisplayCount >= _filteredEntries.Count)
            {
                UpdateStatus($"Showing all {_filteredEntries.Count:N0} of {_allEntries.Count:N0} keys ({filterText})");
            }
            else
            {
                UpdateStatus($"Showing {_currentDisplayCount:N0} of {_filteredEntries.Count:N0} keys ({filterText})");
            }
        }

        private void LoadMore_Click(object? sender, EventArgs e)
        {
            if (_isCollapsed)
            {
                // Load more groups in collapsed mode
                int previousGroupCount = _currentGroupDisplayCount;
                int remainingGroups = _collapsedGroups.Count - _currentGroupDisplayCount;
                int groupsToAdd = Math.Min(_pageSize, remainingGroups);
                _currentGroupDisplayCount += groupsToAdd;
                
                // Rebuild the full display (collapsed mode doesn't support incremental append)
                UpdateCollapsedGridDisplay();
                
                // Scroll to show some of the new content
                if (_timelineGrid.Rows.Count > 0)
                {
                    // Find the row index where new groups start
                    int rowIndex = 0;
                    for (int i = 0; i < previousGroupCount && i < _collapsedGroups.Count; i++)
                    {
                        rowIndex++; // Group header or single-entry row
                        if (_collapsedGroups[i].Entries.Count > 1 && _expandedGroupIndices.Contains(i))
                            rowIndex += _collapsedGroups[i].Entries.Count;
                    }
                    if (rowIndex < _timelineGrid.Rows.Count)
                        _timelineGrid.FirstDisplayedScrollingRowIndex = rowIndex;
                }
                return;
            }
            
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
                    entry.DisplayPath
                );
                // Store the entry reference in the row's Tag for reliable retrieval after sorting
                _timelineGrid.Rows[rowIndex].Tag = entry;
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
                var row = _timelineGrid.Rows[e.RowIndex];
                // In collapsed mode, don't navigate on group header rows (toggle expand instead)
                if (row.Tag is CollapsedGroup) return;
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
                    writer.WriteLine("Last Modified,Key Path");
                    
                    // Write data
                    foreach (var entry in _filteredEntries)
                    {
                        var path = entry.DisplayPath.Replace("\"", "\"\""); // Escape quotes
                        writer.WriteLine($"\"{entry.LastModified:yyyy-MM-dd HH:mm:ss}\",\"{path}\"");
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
            foreach (var btn in new[] { _refreshButton, _exportButton, _collapseButton })
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
        }

        private class CollapsedGroup
        {
            public List<TimelineEntry> Entries { get; } = new();
            public string RootPath { get; set; } = "";
        }
    }
}

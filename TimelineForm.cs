using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Registry.Abstractions;

namespace RegistryExpert
{
    public class TimelineForm : Form
    {
        private readonly OfflineRegistryParser _parser;
        private readonly MainForm _mainForm;
        
        private DataGridView _timelineGrid = null!;
        private ComboBox _filterCombo = null!;
        private ComboBox _limitCombo = null!;
        private DateTimePicker _fromDatePicker = null!;
        private DateTimePicker _toDatePicker = null!;
        private FlowLayoutPanel _customDatePanel = null!;
        private Button _refreshButton = null!;
        private Button _exportButton = null!;
        private Label _statusLabel = null!;
        private ProgressBar _progressBar = null!;
        
        private List<TimelineEntry> _allEntries = new();
        private List<TimelineEntry> _filteredEntries = new();
        private CancellationTokenSource? _scanCts;
        private bool _isScanning;
        
        // Pagination state
        private int _currentDisplayCount = 0;
        private int _pageSize = 100;
        private Panel _loadMorePanel = null!;
        private Button _loadMoreButton = null!;

        public TimelineForm(OfflineRegistryParser parser, MainForm mainForm)
        {
            _parser = parser;
            _mainForm = mainForm;
            InitializeComponent();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _scanCts?.Cancel();
                _scanCts?.Dispose();
                _scanCts = null;
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
            
            // Update DataGridView row height
            _timelineGrid.RowTemplate.Height = DpiHelper.Scale(28);
            _timelineGrid.ColumnHeadersHeight = DpiHelper.Scale(32);
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
            var filterPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = DpiHelper.Scale(100),
                BackColor = ModernTheme.Surface,
                Padding = new Padding(16)
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
                Margin = new Padding(0, 6, 5, 0)
            };

            _filterCombo = new ComboBox
            {
                Size = DpiHelper.ScaleSize(150, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = ModernTheme.RegularFont,
                Margin = new Padding(0, 0, 15, 0)
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
                Margin = new Padding(0, 6, 5, 0)
            };

            _limitCombo = new ComboBox
            {
                Size = DpiHelper.ScaleSize(80, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = ModernTheme.RegularFont,
                Margin = new Padding(0, 0, 5, 0)
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
                Margin = new Padding(0, 6, 20, 0)
            };

            _refreshButton = ModernTheme.CreateButton("Scan", RefreshButton_Click);
            _refreshButton.AutoSize = true;
            _refreshButton.Padding = new Padding(15, 5, 15, 5);
            _refreshButton.Margin = new Padding(0, 0, 10, 0);

            _exportButton = ModernTheme.CreateButton("Export", ExportButton_Click);
            _exportButton.AutoSize = true;
            _exportButton.Padding = new Padding(15, 5, 15, 5);
            _exportButton.Margin = new Padding(0, 0, 0, 0);
            _exportButton.Enabled = false;

            filterFlow.Controls.AddRange(new Control[] { 
                filterLabel, _filterCombo, limitLabel, _limitCombo, keysLabel,
                _refreshButton, _exportButton
            });

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
                Margin = new Padding(0, 6, 5, 0)
            };

            _fromDatePicker = new DateTimePicker
            {
                Size = DpiHelper.ScaleSize(180, 28),
                Format = DateTimePickerFormat.Custom,
                CustomFormat = "yyyy-MM-dd HH:mm",
                Font = ModernTheme.RegularFont,
                Margin = new Padding(0, 0, 15, 0)
            };
            _fromDatePicker.Value = DateTime.Now.AddDays(-7);
            _fromDatePicker.ValueChanged += (s, e) => ApplyFilter();

            var toLabel = new Label
            {
                Text = "To:",
                ForeColor = ModernTheme.TextSecondary,
                Font = ModernTheme.RegularFont,
                AutoSize = true,
                Margin = new Padding(0, 6, 5, 0)
            };

            _toDatePicker = new DateTimePicker
            {
                Size = DpiHelper.ScaleSize(180, 28),
                Format = DateTimePickerFormat.Custom,
                CustomFormat = "yyyy-MM-dd HH:mm",
                Font = ModernTheme.RegularFont,
                Margin = new Padding(0, 0, 0, 0)
            };
            _toDatePicker.Value = DateTime.Now;
            _toDatePicker.ValueChanged += (s, e) => ApplyFilter();

            _customDatePanel.Controls.AddRange(new Control[] { fromLabel, _fromDatePicker, toLabel, _toDatePicker });

            // Add flow panels to filter panel (order matters - bottom to top for Dock.Top)
            filterPanel.Controls.Add(_customDatePanel);
            filterPanel.Controls.Add(filterFlow);

            // Status bar
            var statusPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = DpiHelper.Scale(36),
                BackColor = ModernTheme.Surface,
                Padding = new Padding(16, 0, 16, 0)
            };

            _statusLabel = new Label
            {
                Text = "Click 'Scan' to analyze registry key timestamps",
                ForeColor = ModernTheme.TextSecondary,
                Font = ModernTheme.RegularFont,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };

            _progressBar = new ProgressBar
            {
                Dock = DockStyle.Bottom,
                Height = DpiHelper.Scale(3),
                Style = ProgressBarStyle.Marquee,
                Visible = false
            };

            statusPanel.Controls.Add(_statusLabel);
            statusPanel.Controls.Add(_progressBar);

            // Timeline grid
            _timelineGrid = new DataGridView { Dock = DockStyle.Fill };
            ModernTheme.ApplyTo(_timelineGrid);
            _timelineGrid.BackgroundColor = ModernTheme.Background;  // Override: use Background instead of Surface
            _timelineGrid.DefaultCellStyle.Padding = new Padding(8, 4, 8, 4);  // Override default padding
            _timelineGrid.ColumnHeadersDefaultCellStyle.Padding = new Padding(8, 4, 8, 4);
            _timelineGrid.ColumnHeadersHeight = DpiHelper.Scale(36);  // Override default

            // Add columns
            _timelineGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "LastModified",
                HeaderText = "Last Modified",
                FillWeight = 20,
                DefaultCellStyle = { Format = "yyyy-MM-dd HH:mm:ss" }
            });
            _timelineGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "KeyPath",
                HeaderText = "Key Path",
                FillWeight = 60
            });
            _timelineGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "ValuesCount",
                HeaderText = "Values",
                FillWeight = 10,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            _timelineGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "SubkeysCount",
                HeaderText = "Subkeys",
                FillWeight = 10,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });

            // Enable sorting
            foreach (DataGridViewColumn col in _timelineGrid.Columns)
            {
                col.SortMode = DataGridViewColumnSortMode.Automatic;
            }

            _timelineGrid.CellDoubleClick += TimelineGrid_CellDoubleClick;
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
            this.Controls.Add(filterPanel);
            this.Controls.Add(statusPanel);

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
            _customDatePanel.Visible = _filterCombo.SelectedIndex == 5; // "Custom Range..."
            ApplyFilter();
        }

        private async void RefreshButton_Click(object? sender, EventArgs e)
        {
            if (_isScanning)
            {
                _scanCts?.Cancel();
                return;
            }
            await ScanRegistryAsync();
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
            _statusLabel.Text = "Scanning registry keys...";
            _exportButton.Enabled = false;
            _timelineGrid.Rows.Clear();
            _allEntries.Clear();

            try
            {
                var rootKey = _parser.GetRootKey();
                if (rootKey == null)
                {
                    _statusLabel.Text = "No registry hive loaded";
                    return;
                }

                var entries = new List<TimelineEntry>();
                int scannedCount = 0;

                await Task.Run(() =>
                {
                    ScanKeyRecursive(rootKey, entries, ref scannedCount, token);
                }, token);

                if (token.IsCancellationRequested)
                {
                    _statusLabel.Text = $"Scan cancelled - found {entries.Count:N0} keys";
                }
                else
                {
                    _allEntries = entries.OrderByDescending(e => e.LastModified).ToList();
                    _statusLabel.Text = $"Scanned {scannedCount:N0} keys - {_allEntries.Count:N0} have timestamps";
                }

                ApplyFilter();
                _exportButton.Enabled = _filteredEntries.Count > 0;
            }
            catch (OperationCanceledException)
            {
                _statusLabel.Text = $"Scan cancelled - found {_allEntries.Count:N0} keys";
            }
            catch (Exception ex)
            {
                _statusLabel.Text = $"Error: {ex.Message}";
                _statusLabel.ForeColor = ModernTheme.Error;
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
                            _statusLabel.Text = $"Scanning... {count:N0} keys processed";
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
                    DisplayPath = displayPath,
                    ValuesCount = key.Values?.Count ?? 0,
                    SubkeysCount = key.SubKeys?.Count ?? 0,
                    RegistryKey = key
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
                _pageSize = int.MaxValue; // Show all at once
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

            _filteredEntries = filtered.ToList();
            
            // Reset pagination - show first page
            _currentDisplayCount = Math.Min(_pageSize, _filteredEntries.Count);
            
            UpdateGridDisplay();
            _exportButton.Enabled = _filteredEntries.Count > 0;
        }

        private void UpdateGridDisplay()
        {
            _timelineGrid.Rows.Clear();
            
            // Display only up to _currentDisplayCount entries
            var entriesToShow = _filteredEntries.Take(_currentDisplayCount).ToList();
            
            foreach (var entry in entriesToShow)
            {
                var rowIndex = _timelineGrid.Rows.Add(
                    entry.LastModified,
                    entry.DisplayPath,
                    entry.ValuesCount,
                    entry.SubkeysCount
                );
                // Store the entry reference in the row's Tag for reliable retrieval after sorting
                _timelineGrid.Rows[rowIndex].Tag = entry;
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
            
            if (_filteredEntries.Count == 0)
            {
                _statusLabel.Text = $"No keys found ({filterText})";
            }
            else if (_currentDisplayCount >= _filteredEntries.Count)
            {
                _statusLabel.Text = $"Showing all {_filteredEntries.Count:N0} of {_allEntries.Count:N0} keys ({filterText})";
            }
            else
            {
                _statusLabel.Text = $"Showing {_currentDisplayCount:N0} of {_filteredEntries.Count:N0} keys ({filterText})";
            }
            _statusLabel.ForeColor = ModernTheme.TextSecondary;
        }

        private void LoadMore_Click(object? sender, EventArgs e)
        {
            // Add another page of results
            int previousCount = _currentDisplayCount;
            _currentDisplayCount = Math.Min(_currentDisplayCount + _pageSize, _filteredEntries.Count);
            
            // Append new rows to the grid (more efficient than rebuilding)
            var newEntries = _filteredEntries.Skip(previousCount).Take(_currentDisplayCount - previousCount);
            foreach (var entry in newEntries)
            {
                var rowIndex = _timelineGrid.Rows.Add(
                    entry.LastModified,
                    entry.DisplayPath,
                    entry.ValuesCount,
                    entry.SubkeysCount
                );
                // Store the entry reference in the row's Tag for reliable retrieval after sorting
                _timelineGrid.Rows[rowIndex].Tag = entry;
            }

            // Update Load More button visibility and text
            int remaining = _filteredEntries.Count - _currentDisplayCount;
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
                int scrollToRow = Math.Min(previousCount + 5, _timelineGrid.Rows.Count - 1);
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
                        _statusLabel.Text = "Path copied to clipboard";
                    }
                    catch (System.Runtime.InteropServices.ExternalException)
                    {
                        _statusLabel.Text = "Failed to copy - clipboard may be in use";
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
                    writer.WriteLine("Last Modified,Key Path,Values Count,Subkeys Count");
                    
                    // Write data
                    foreach (var entry in _filteredEntries)
                    {
                        var path = entry.DisplayPath.Replace("\"", "\"\""); // Escape quotes
                        writer.WriteLine($"\"{entry.LastModified:yyyy-MM-dd HH:mm:ss}\",\"{path}\",{entry.ValuesCount},{entry.SubkeysCount}");
                    }

                    _statusLabel.Text = $"Exported {_filteredEntries.Count:N0} entries to {Path.GetFileName(dialog.FileName)}";
                    _statusLabel.ForeColor = ModernTheme.Success;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Export failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    _statusLabel.Text = "Export failed";
                    _statusLabel.ForeColor = ModernTheme.Error;
                }
            }
        }

        private string ConvertRootPath(string path)
        {
            if (string.IsNullOrEmpty(path)) return path;

            var hiveType = _parser.CurrentHiveType;
            
            if (path.StartsWith("ROOT\\", StringComparison.OrdinalIgnoreCase))
            {
                return hiveType.ToString() + path.Substring(4);
            }
            else if (path.Equals("ROOT", StringComparison.OrdinalIgnoreCase))
            {
                return hiveType.ToString();
            }
            
            return path;
        }

        /// <summary>
        /// Refresh theme colors
        /// </summary>
        public void RefreshTheme()
        {
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
            _timelineGrid.GridColor = ModernTheme.Border;

            // Refresh Load More panel and button
            _loadMorePanel.BackColor = ModernTheme.Background;
            _loadMoreButton.BackColor = ModernTheme.Surface;
            _loadMoreButton.ForeColor = ModernTheme.TextPrimary;
            _loadMoreButton.FlatAppearance.BorderColor = ModernTheme.Border;
            _loadMoreButton.FlatAppearance.MouseOverBackColor = ModernTheme.Selection;

            // Refresh other controls
            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is Panel panel && ctrl != _loadMorePanel)
                {
                    panel.BackColor = ModernTheme.Surface;
                    foreach (Control child in panel.Controls)
                    {
                        if (child is Label label)
                            label.ForeColor = ModernTheme.TextSecondary;
                        else if (child is ComboBox combo)
                        {
                            combo.BackColor = ModernTheme.Surface;
                            combo.ForeColor = ModernTheme.TextPrimary;
                        }
                    }
                }
            }

            _statusLabel.ForeColor = ModernTheme.TextSecondary;
            this.Refresh();
        }

        private class TimelineEntry
        {
            public DateTime LastModified { get; set; }
            public string KeyPath { get; set; } = "";
            public string DisplayPath { get; set; } = "";
            public int ValuesCount { get; set; }
            public int SubkeysCount { get; set; }
            public RegistryKey? RegistryKey { get; set; }
        }
    }
}

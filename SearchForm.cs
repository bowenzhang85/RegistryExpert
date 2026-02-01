using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Registry.Abstractions;

namespace RegistryExpert
{
    public class SearchForm : Form
    {
        private readonly OfflineRegistryParser _parser;
        private readonly MainForm _mainForm;
        
        private TextBox _searchBox = null!;
        private Button _searchButton = null!;
        private Button _cancelButton = null!;
        private DataGridView _resultsGrid = null!;
        private Label _statusLabel = null!;
        private Panel _previewPanel = null!;
        private Label _previewPathLabel = null!;
        private RichTextBox _previewValueBox = null!;
        private List<SearchResult> _searchResults = new();
        private string _currentSearchTerm = "";
        private CancellationTokenSource? _searchCts;

        // Public properties to preserve search state across theme changes
        public string SearchTerm => _searchBox?.Text ?? "";

        public SearchForm(OfflineRegistryParser parser, MainForm mainForm, string? initialSearchTerm = null)
        {
            _parser = parser;
            _mainForm = mainForm;
            InitializeComponent();
            
            // If initial search term provided, set it and auto-search
            if (!string.IsNullOrEmpty(initialSearchTerm))
            {
                _searchBox.Text = initialSearchTerm;
                // Delay the search to allow form to fully load
                this.Shown += async (s, e) => await SearchAsync();
            }
        }

        /// <summary>
        /// Handle DPI changes when moving between monitors with different DPI settings.
        /// </summary>
        protected override void OnDpiChanged(DpiChangedEventArgs e)
        {
            // Reset the cached DPI scale factor so it gets recalculated
            DpiHelper.ResetScaleFactor();
            
            base.OnDpiChanged(e);
            
            // Update controls that need manual DPI adjustment
            _resultsGrid.RowTemplate.Height = DpiHelper.Scale(28);
            _resultsGrid.ColumnHeadersHeight = DpiHelper.Scale(32);
        }

        private void InitializeComponent()
        {
            this.Text = "Search Registry";
            this.Size = new Size(950, 650);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MinimumSize = new Size(700, 500);
            this.AutoScaleMode = AutoScaleMode.Dpi;
            ModernTheme.ApplyTo(this);

            // Top search panel - use FlowLayoutPanel for DPI scaling
            var topPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = DpiHelper.Scale(50),
                BackColor = ModernTheme.Surface,
                Padding = new Padding(15, 10, 15, 10),
                WrapContents = false
            };

            var searchLabel = new Label
            {
                Text = "Search for:",
                ForeColor = ModernTheme.TextSecondary,
                Font = ModernTheme.RegularFont,
                AutoSize = true,
                Margin = new Padding(0, 6, 10, 0)
            };

            _searchBox = new TextBox
            {
                Size = DpiHelper.ScaleSize(550, 28),
                Font = ModernTheme.RegularFont,
                Margin = new Padding(0, 0, 10, 0)
            };
            ModernTheme.ApplyTo(_searchBox);
            _searchBox.KeyDown += async (s, e) => { if (e.KeyCode == Keys.Enter) await SearchAsync(); };

            _searchButton = ModernTheme.CreateButton("Search", async (s, e) => await SearchAsync());
            _searchButton.AutoSize = true;
            _searchButton.Padding = new Padding(15, 5, 15, 5);
            _searchButton.Margin = new Padding(0, 0, 10, 0);

            _cancelButton = ModernTheme.CreateButton("Cancel", (s, e) => CancelSearch());
            _cancelButton.AutoSize = true;
            _cancelButton.Padding = new Padding(15, 5, 15, 5);
            _cancelButton.BackColor = ModernTheme.Error;
            _cancelButton.Visible = false;

            topPanel.Controls.AddRange(new Control[] { searchLabel, _searchBox, _searchButton, _cancelButton });

            // Status bar
            var statusPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = DpiHelper.Scale(32),
                BackColor = ModernTheme.Surface,
                Padding = new Padding(12, 0, 12, 0)
            };

            _statusLabel = new Label
            {
                Text = "Enter a search term and press Search",
                ForeColor = ModernTheme.TextSecondary,
                Font = ModernTheme.RegularFont,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };
            statusPanel.Controls.Add(_statusLabel);

            // Preview panel at bottom
            _previewPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = DpiHelper.Scale(120),
                BackColor = ModernTheme.TreeViewBack,
                Padding = new Padding(10)
            };

            var previewHeader = new Label
            {
                Text = "Registry Value Preview",
                ForeColor = ModernTheme.TextSecondary,
                Font = ModernTheme.BoldFont,
                Dock = DockStyle.Top,
                Height = DpiHelper.Scale(20)
            };

            _previewPathLabel = new Label
            {
                Text = "Select a result to preview",
                ForeColor = ModernTheme.Accent,
                Font = ModernTheme.MonoFont,
                Dock = DockStyle.Top,
                Height = DpiHelper.Scale(22),
                AutoEllipsis = true
            };

            _previewValueBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BackColor = ModernTheme.Surface,
                ForeColor = ModernTheme.TextPrimary,
                Font = ModernTheme.MonoFont,
                BorderStyle = BorderStyle.FixedSingle,
                DetectUrls = false
            };

            _previewPanel.Controls.Add(_previewValueBox);
            _previewPanel.Controls.Add(_previewPathLabel);
            _previewPanel.Controls.Add(previewHeader);

            // Results header
            var resultsHeader = new Panel
            {
                Dock = DockStyle.Top,
                Height = DpiHelper.Scale(36),
                BackColor = ModernTheme.Surface,
                Padding = DpiHelper.ScalePadding(12, 0, 12, 0)
            };

            var resultsLabel = new Label
            {
                Text = "Results",
                ForeColor = ModernTheme.TextPrimary,
                Font = ModernTheme.HeaderFont,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };
            
            var border = new Panel
            {
                Height = 1,
                Dock = DockStyle.Bottom,
                BackColor = ModernTheme.Border
            };
            
            resultsHeader.Controls.Add(resultsLabel);
            resultsHeader.Controls.Add(border);

            // Results DataGridView for better column handling
            _resultsGrid = new DataGridView { Dock = DockStyle.Fill };
            ModernTheme.ApplyTo(_resultsGrid);
            _resultsGrid.RowTemplate.Height = DpiHelper.Scale(28);  // Scale for DPI
            _resultsGrid.ColumnHeadersHeight = DpiHelper.Scale(32); // Scale for DPI

            // Add columns with proper sizing
            _resultsGrid.Columns.Add("keyPath", "Key Path");
            _resultsGrid.Columns.Add("matchType", "Type");
            _resultsGrid.Columns.Add("details", "Name");
            _resultsGrid.Columns["keyPath"].FillWeight = 50;
            _resultsGrid.Columns["matchType"].FillWeight = 15;
            _resultsGrid.Columns["details"].FillWeight = 35;

            _resultsGrid.SelectionChanged += ResultsGrid_SelectionChanged;
            _resultsGrid.CellDoubleClick += ResultsGrid_DoubleClick;

            // Add context menu for right-click options
            var contextMenu = new ContextMenuStrip();
            contextMenu.Renderer = new ToolStripProfessionalRenderer(new ModernColorTable());
            
            var copyPathItem = new ToolStripMenuItem("Copy Path", null, (s, ev) => CopySelectedPath());
            copyPathItem.ShortcutKeys = Keys.Control | Keys.C;
            var copyValueItem = new ToolStripMenuItem("Copy Value", null, (s, ev) => CopySelectedValue());
            var navigateItem = new ToolStripMenuItem("Navigate to Key", null, (s, ev) => NavigateToSelectedKey(closeAfterNavigate: false));
            
            contextMenu.Items.Add(copyPathItem);
            contextMenu.Items.Add(copyValueItem);
            contextMenu.Items.Add(new ToolStripSeparator());
            contextMenu.Items.Add(navigateItem);
            
            // Update colors when menu opens (supports theme switching)
            contextMenu.Opening += (s, ev) =>
            {
                contextMenu.BackColor = ModernTheme.Surface;
                contextMenu.ForeColor = ModernTheme.TextPrimary;
                foreach (ToolStripItem item in contextMenu.Items)
                {
                    item.ForeColor = ModernTheme.TextPrimary;
                    item.BackColor = ModernTheme.Surface;
                }
            };
            
            _resultsGrid.ContextMenuStrip = contextMenu;
            
            // Handle right-click to select row
            _resultsGrid.CellMouseDown += (s, ev) =>
            {
                if (ev.Button == MouseButtons.Right && ev.RowIndex >= 0)
                {
                    _resultsGrid.ClearSelection();
                    _resultsGrid.Rows[ev.RowIndex].Selected = true;
                }
            };

            // Add controls
            this.Controls.Add(_resultsGrid);
            this.Controls.Add(resultsHeader);
            this.Controls.Add(_previewPanel);
            this.Controls.Add(topPanel);
            this.Controls.Add(statusPanel);

            // Resize handler for search box - make it expand to fill available space
            // Note: Setting Left on buttons is not needed in FlowLayoutPanel (it manages positioning automatically)
            this.Resize += (s, e) =>
            {
                // Calculate available width after accounting for label, buttons, margins, and padding
                // Using DpiHelper.Scale for DPI-aware minimum width
                _searchBox.Width = Math.Max(DpiHelper.Scale(200), topPanel.Width - DpiHelper.Scale(380));
            };
        }

        private void CancelSearch()
        {
            _searchCts?.Cancel();
        }

        private async Task SearchAsync()
        {
            var searchTerm = _searchBox.Text.Trim();
            if (string.IsNullOrEmpty(searchTerm))
            {
                MessageBox.Show("Please enter a search term.", "Search", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            _currentSearchTerm = searchTerm;

            // Cancel and dispose any existing search (proper disposal before reassignment)
            var oldCts = _searchCts;
            _searchCts = new CancellationTokenSource();
            var token = _searchCts.Token;
            oldCts?.Cancel();
            oldCts?.Dispose();

            _resultsGrid.Rows.Clear();
            _searchResults.Clear();
            _searchButton.Enabled = false;
            _searchButton.Visible = false;
            _cancelButton.Visible = true;
            _statusLabel.Text = "Searching...";
            _statusLabel.ForeColor = ModernTheme.Warning;
            _previewPathLabel.Text = "Searching...";
            _previewValueBox.Text = "";

            try
            {
                // Run search on background thread (always case-insensitive)
                var results = await Task.Run(() => 
                    _parser.SearchKeys(searchTerm, caseSensitive: false), token);
                
                if (token.IsCancellationRequested) return;
                
                int count = 0;
                
                foreach (var key in results.Take(1000))
                {
                    if (token.IsCancellationRequested) break;
                    
                    var searchResult = new SearchResult { KeyPath = key.KeyPath, Key = key };
                    
                    if (key.KeyName.Contains(searchTerm, StringComparison.OrdinalIgnoreCase))
                    {
                        searchResult.MatchType = "Key";
                        searchResult.Details = key.KeyName;
                        searchResult.FullValue = $"Key: {GetDisplayPath(key.KeyPath)}";
                    }
                    else
                    {
                        var matchingValue = key.Values.FirstOrDefault(v =>
                            v.ValueName.Contains(searchTerm, StringComparison.OrdinalIgnoreCase) ||
                            (v.ValueData?.ToString() ?? "").Contains(searchTerm, StringComparison.OrdinalIgnoreCase));
                        
                        if (matchingValue != null)
                        {
                            searchResult.MatchType = matchingValue.ValueType;
                            searchResult.ValueName = matchingValue.ValueName;
                            searchResult.ValueData = CleanValueData(matchingValue.ValueData?.ToString() ?? "");
                            searchResult.ValueType = matchingValue.ValueType;
                            // Show value name in Name column (use "(Default)" for empty name)
                            searchResult.Details = string.IsNullOrEmpty(matchingValue.ValueName) ? "(Default)" : matchingValue.ValueName;
                            searchResult.FullValue = $"Name: {(string.IsNullOrEmpty(matchingValue.ValueName) ? "(Default)" : matchingValue.ValueName)}\nType: {matchingValue.ValueType}\nData: {searchResult.ValueData}";
                        }
                        else
                        {
                            searchResult.MatchType = "Data";
                            searchResult.Details = "";
                            searchResult.FullValue = $"Key: {GetDisplayPath(key.KeyPath)}";
                        }
                    }
                    
                    _searchResults.Add(searchResult);
                    _resultsGrid.Rows.Add(GetDisplayPath(key.KeyPath), searchResult.MatchType, searchResult.Details);
                    count++;
                    
                    // Update status periodically to show progress
                    if (count % 100 == 0)
                    {
                        _statusLabel.Text = $"Found {count} results so far...";
                    }
                }

                if (token.IsCancellationRequested)
                {
                    _statusLabel.Text = $"Search cancelled - {_searchResults.Count} results found";
                    _statusLabel.ForeColor = ModernTheme.Warning;
                }
                else
                {
                    _statusLabel.Text = $"Found {results.Count} results" + (results.Count > 1000 ? " (showing first 1000)" : "");
                    _statusLabel.ForeColor = results.Count > 0 ? ModernTheme.Success : ModernTheme.TextSecondary;
                }
                _previewPathLabel.Text = _searchResults.Count > 0 ? "Select a result to preview" : "No results found";
            }
            catch (OperationCanceledException)
            {
                _statusLabel.Text = $"Search cancelled - {_searchResults.Count} results found";
                _statusLabel.ForeColor = ModernTheme.Warning;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Search error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _statusLabel.Text = "Search failed";
                _statusLabel.ForeColor = ModernTheme.Error;
            }
            finally
            {
                _searchButton.Enabled = true;
                _searchButton.Visible = true;
                _cancelButton.Visible = false;
            }
        }

        private void ResultsGrid_SelectionChanged(object? sender, EventArgs e)
        {
            if (_resultsGrid.SelectedRows.Count > 0 && _resultsGrid.SelectedRows[0].Index < _searchResults.Count)
            {
                var result = _searchResults[_resultsGrid.SelectedRows[0].Index];
                _previewPathLabel.Text = GetDisplayPath(result.KeyPath);
                SetPreviewWithHighlight(result.FullValue, _currentSearchTerm);
            }
        }

        private void SetPreviewWithHighlight(string text, string searchTerm)
        {
            _previewValueBox.Text = text;
            
            if (string.IsNullOrEmpty(searchTerm) || string.IsNullOrEmpty(text))
                return;
            
            // Highlight all occurrences of the search term (case-insensitive)
            int index = 0;
            while ((index = text.IndexOf(searchTerm, index, StringComparison.OrdinalIgnoreCase)) >= 0)
            {
                _previewValueBox.Select(index, searchTerm.Length);
                _previewValueBox.SelectionBackColor = Color.FromArgb(255, 255, 0); // Yellow highlight
                _previewValueBox.SelectionColor = Color.Black;
                index += searchTerm.Length;
            }
            
            // Reset selection to start
            _previewValueBox.Select(0, 0);
        }

        // Pre-allocated array to avoid per-call allocations
        private static readonly string[] _prefixesToRemove = 
        { 
            "RegSzData: ", "RegSzData:",
            "RegExpandSzData: ", "RegExpandSzData:",
            "RegDwordData: ", "RegDwordData:",
            "RegQwordData: ", "RegQwordData:",
            "RegBinaryData: ", "RegBinaryData:",
            "RegMultiSzData: ", "RegMultiSzData:",
            "Name: ", "Name:",
            "Data: ", "Data:"
        };
        
        private string CleanValueData(string valueData)
        {
            if (string.IsNullOrEmpty(valueData))
                return valueData;
            
            var result = valueData;
            foreach (var prefix in _prefixesToRemove)
            {
                if (result.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    result = result.Substring(prefix.Length);
                    break;
                }
            }
            
            return result.Trim();
        }

        private void ResultsGrid_DoubleClick(object? sender, DataGridViewCellEventArgs e)
        {
            NavigateToSelectedKey(closeAfterNavigate: false);
        }

        private void CopySelectedPath()
        {
            if (_resultsGrid.SelectedRows.Count > 0 && _resultsGrid.SelectedRows[0].Index < _searchResults.Count)
            {
                var result = _searchResults[_resultsGrid.SelectedRows[0].Index];
                if (!string.IsNullOrEmpty(result.KeyPath))
                {
                    try
                    {
                        Clipboard.SetText(GetDisplayPath(result.KeyPath));
                    }
                    catch (System.Runtime.InteropServices.ExternalException)
                    {
                        // Clipboard may be locked by another application - silently ignore
                    }
                }
            }
        }

        private void CopySelectedValue()
        {
            if (_resultsGrid.SelectedRows.Count > 0 && _resultsGrid.SelectedRows[0].Index < _searchResults.Count)
            {
                var result = _searchResults[_resultsGrid.SelectedRows[0].Index];
                try
                {
                    if (!string.IsNullOrEmpty(result.ValueData))
                    {
                        Clipboard.SetText(result.ValueData);
                    }
                    else if (!string.IsNullOrEmpty(result.FullValue))
                    {
                        Clipboard.SetText(result.FullValue);
                    }
                }
                catch (System.Runtime.InteropServices.ExternalException)
                {
                    // Clipboard may be locked by another application - silently ignore
                }
            }
        }

        private void NavigateToSelectedKey(bool closeAfterNavigate = false)
        {
            if (_resultsGrid.SelectedRows.Count > 0 && _resultsGrid.SelectedRows[0].Index < _searchResults.Count)
            {
                var result = _searchResults[_resultsGrid.SelectedRows[0].Index];
                
                // If the match is a value (not a key), pass the value name to auto-select it
                string? valueNameToSelect = null;
                if (result.MatchType != "Key" && !string.IsNullOrEmpty(result.ValueName))
                {
                    valueNameToSelect = string.IsNullOrEmpty(result.ValueName) ? "(Default)" : result.ValueName;
                }
                else if (result.MatchType != "Key" && result.Details != null && result.Details != "")
                {
                    // Details contains the value name for value matches
                    valueNameToSelect = result.Details;
                }
                
                _mainForm.NavigateToKey(result.KeyPath, valueNameToSelect);
                _mainForm.BringToFront();
                _mainForm.Activate();
                
                if (closeAfterNavigate)
                {
                    this.Close();
                }
                else
                {
                    // Bring search window back to front after a brief delay
                    this.BringToFront();
                }
            }
        }

        /// <summary>
        /// Refresh all control colors for theme change without reloading data
        /// </summary>
        public void RefreshTheme()
        {
            // Apply theme to form
            this.BackColor = ModernTheme.Background;
            this.ForeColor = ModernTheme.TextPrimary;
            
            // Recursively update all controls
            RefreshControlTheme(this);
            
            // Update grid specific styles
            _resultsGrid.BackgroundColor = ModernTheme.Surface;
            _resultsGrid.ForeColor = ModernTheme.TextPrimary;
            _resultsGrid.GridColor = ModernTheme.Border;
            _resultsGrid.DefaultCellStyle.BackColor = ModernTheme.Surface;
            _resultsGrid.DefaultCellStyle.ForeColor = ModernTheme.TextPrimary;
            _resultsGrid.DefaultCellStyle.SelectionBackColor = ModernTheme.Selection;
            _resultsGrid.DefaultCellStyle.SelectionForeColor = ModernTheme.TextPrimary;
            _resultsGrid.AlternatingRowsDefaultCellStyle.BackColor = ModernTheme.ListViewAltRow;
            _resultsGrid.ColumnHeadersDefaultCellStyle.BackColor = ModernTheme.TreeViewBack;
            _resultsGrid.ColumnHeadersDefaultCellStyle.ForeColor = ModernTheme.TextSecondary;
            _resultsGrid.ColumnHeadersDefaultCellStyle.SelectionBackColor = ModernTheme.TreeViewBack;
            _resultsGrid.ColumnHeadersDefaultCellStyle.SelectionForeColor = ModernTheme.TextSecondary;
            
            // Update preview box
            _previewValueBox.BackColor = ModernTheme.Surface;
            _previewValueBox.ForeColor = ModernTheme.TextPrimary;
            
            // Re-highlight search term in preview if there's content
            if (!string.IsNullOrEmpty(_previewValueBox.Text) && !string.IsNullOrEmpty(_currentSearchTerm))
            {
                var text = _previewValueBox.Text;
                SetPreviewWithHighlight(text, _currentSearchTerm);
            }
            
            this.Refresh();
        }

        private void RefreshControlTheme(Control parent)
        {
            foreach (Control ctrl in parent.Controls)
            {
                if (ctrl is Panel panel)
                {
                    panel.BackColor = ModernTheme.Surface;
                }
                else if (ctrl is Label label)
                {
                    if (label == _previewPathLabel)
                        label.ForeColor = ModernTheme.Accent;
                    else if (label == _statusLabel)
                        label.ForeColor = ModernTheme.TextSecondary;
                    else
                        label.ForeColor = ModernTheme.TextPrimary;
                }
                else if (ctrl is TextBox textBox)
                {
                    textBox.BackColor = ModernTheme.Surface;
                    textBox.ForeColor = ModernTheme.TextPrimary;
                }
                else if (ctrl is RichTextBox richTextBox)
                {
                    richTextBox.BackColor = ModernTheme.Surface;
                    richTextBox.ForeColor = ModernTheme.TextPrimary;
                }
                else if (ctrl is CheckBox checkBox)
                {
                    checkBox.ForeColor = ModernTheme.TextSecondary;
                }
                else if (ctrl is Button button)
                {
                    button.BackColor = ModernTheme.Accent;
                    button.ForeColor = Color.White;
                    button.FlatAppearance.BorderColor = ModernTheme.Accent;
                }
                
                // Recurse into child controls
                if (ctrl.Controls.Count > 0)
                {
                    RefreshControlTheme(ctrl);
                }
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _searchCts?.Cancel();
                _searchCts?.Dispose();
            }
            base.Dispose(disposing);
        }

        /// <summary>
        /// Convert a KeyPath from ROOT\... to HIVENAME\... for display
        /// </summary>
        private string GetDisplayPath(string keyPath)
        {
            if (string.IsNullOrEmpty(keyPath)) return keyPath;
            
            var hiveName = _parser.CurrentHiveType.ToString();
            
            // Replace ROOT\ with HIVENAME\ or just ROOT with HIVENAME if it's the root itself
            if (keyPath.StartsWith("ROOT\\", StringComparison.OrdinalIgnoreCase))
            {
                return hiveName + keyPath.Substring(4); // "ROOT" is 4 chars
            }
            else if (keyPath.Equals("ROOT", StringComparison.OrdinalIgnoreCase))
            {
                return hiveName;
            }
            
            return keyPath;
        }

        private class SearchResult
        {
            public string KeyPath { get; set; } = "";
            public string MatchType { get; set; } = "";
            public string Details { get; set; } = "";
            public string ValueName { get; set; } = "";
            public string ValueData { get; set; } = "";
            public string ValueType { get; set; } = "";
            public string FullValue { get; set; } = "";
            public RegistryKey? Key { get; set; }
        }
    }
}

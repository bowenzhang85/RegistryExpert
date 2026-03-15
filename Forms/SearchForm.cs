using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using RegistryParser.Abstractions;

namespace RegistryExpert
{
    public class SearchForm : Form
    {
        private readonly IReadOnlyList<(OfflineRegistryParser Parser, string HiveTypeName)> _parsers;
        private readonly MainForm _mainForm;
        
        private TextBox _searchBox = null!;
        private Button _searchButton = null!;
        private Button _cancelButton = null!;
        private CheckBox _matchWholeWordCheckBox = null!;
        private DataGridView _resultsGrid = null!;
        private Label _statusLabel = null!;
        private SplitContainer _splitContainer = null!;
        private Label _previewPathLabel = null!;
        private RichTextBox _previewValueBox = null!;
        private List<SearchResult> _searchResults = new();
        private List<(SearchMatch Match, OfflineRegistryParser Parser)> _allMatches = new();
        private string _currentSearchTerm = "";
        private CancellationTokenSource? _searchCts;
        private Panel _loadMorePanel = null!;
        private Button _loadMoreButton = null!;
        private ImageList? _searchIconList;
        private List<Panel> _separators = new();
        private const int PageSize = 1000;
        private int _displayedMatchCount;

        // Public properties to preserve search state across theme changes
        public string SearchTerm => _searchBox?.Text ?? "";

        public SearchForm(IReadOnlyList<(OfflineRegistryParser Parser, string HiveTypeName)> parsers, MainForm mainForm, string? initialSearchTerm = null)
        {
            _parsers = parsers;
            _mainForm = mainForm;
            InitializeComponent();
            this.Icon = mainForm.Icon;
            
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
        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            ModernTheme.ApplyWindowStyle(this);
        }

        protected override void OnDpiChanged(DpiChangedEventArgs e)
        {
            // Reset the cached DPI scale factor so it gets recalculated
            DpiHelper.ResetScaleFactor();
            
            base.OnDpiChanged(e);
            
            // Update controls that need manual DPI adjustment
            ModernTheme.ApplyTo(_resultsGrid);

            // Rescale pill button sizes and margins for new DPI
            _searchButton.Size = DpiHelper.ScaleSize(90, 30);
            _searchButton.Margin = DpiHelper.ScalePadding(0, 0, 6, 0);
            _cancelButton.Size = DpiHelper.ScaleSize(90, 30);
            _loadMoreButton.Size = DpiHelper.ScaleSize(130, 30);
            _loadMorePanel.Height = DpiHelper.Scale(40);
            
            // Rescale SplitContainer min sizes for new DPI
            _splitContainer.Panel1MinSize = DpiHelper.Scale(100);
            _splitContainer.Panel2MinSize = DpiHelper.Scale(80);
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
                Margin = new Padding(0, 0, 10, 0),
                AccessibleName = "Search for"
            };
            ModernTheme.ApplyTo(_searchBox);
            _searchBox.KeyDown += async (s, e) => { if (e.KeyCode == Keys.Enter) await SearchAsync(); };

            _searchButton = CreatePillButton("Search", async (s, e) => await SearchAsync());
            _searchButton.Margin = DpiHelper.ScalePadding(0, 0, 0, 0);

            _cancelButton = CreatePillButton("Cancel", (s, e) => CancelSearch(), ModernTheme.Error);
            _cancelButton.Margin = DpiHelper.ScalePadding(0, 0, 0, 0);
            _cancelButton.Visible = false;

            // Vertical separators flanking the search/cancel button (matching TimelineForm toolbar)
            var sepLeft = CreateFilterSeparator();
            var sepRight = CreateFilterSeparator();

            _matchWholeWordCheckBox = new CheckBox
            {
                Text = "Match whole word",
                ForeColor = ModernTheme.TextSecondary,
                Font = ModernTheme.RegularFont,
                AutoSize = true,
                Margin = DpiHelper.ScalePadding(8, 6, 0, 0),
                Checked = false
            };

            topPanel.Controls.AddRange(new Control[] { searchLabel, _searchBox, sepLeft, _searchButton, _cancelButton, sepRight, _matchWholeWordCheckBox });

            // Status bar
            var statusPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = DpiHelper.Scale(32),
                BackColor = ModernTheme.Surface,
                Padding = new Padding(12, 0, 12, 0)
            };

            _loadMoreButton = CreatePillButton("Load next 1000", async (s, e) => await LoadMoreResultsAsync());
            _loadMoreButton.Size = DpiHelper.ScaleSize(130, 30);
            _loadMoreButton.Anchor = AnchorStyles.None;

            _loadMorePanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = DpiHelper.Scale(40),
                BackColor = ModernTheme.Surface,
                Visible = false
            };
            _loadMorePanel.Controls.Add(_loadMoreButton);

            // Center the button horizontally when panel resizes
            _loadMorePanel.Resize += (s, e) =>
            {
                _loadMoreButton.Left = (_loadMorePanel.Width - _loadMoreButton.Width) / 2;
                _loadMoreButton.Top = (_loadMorePanel.Height - _loadMoreButton.Height) / 2;
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

            // Preview panel contents
            var previewPanel = new Panel
            {
                Dock = DockStyle.Fill,
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
                DetectUrls = false,
                AccessibleName = "Registry Value Preview"
            };

            previewPanel.Controls.Add(_previewValueBox);
            previewPanel.Controls.Add(_previewPathLabel);
            previewPanel.Controls.Add(previewHeader);

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
            _resultsGrid = new DataGridView { Dock = DockStyle.Fill, AccessibleName = "Search Results" };
            ModernTheme.ApplyTo(_resultsGrid);

            // Create icon ImageList for search result type icons
            _searchIconList = ModernTheme.CreateSearchImageList();

            // Add columns with proper sizing
            _resultsGrid.Columns.Add("keyPath", "Key Path");
            _resultsGrid.Columns.Add("details", "Name");
            _resultsGrid.Columns.Add("matchType", "Type");
            _resultsGrid.Columns.Add("data", "Data");
            _resultsGrid.Columns["keyPath"].FillWeight = 40;
            _resultsGrid.Columns["details"].FillWeight = 15;
            _resultsGrid.Columns["matchType"].FillWeight = 10;
            _resultsGrid.Columns["data"].FillWeight = 35;

            // Reserve left padding in Key Path column for the icon
            var iconPadding = _searchIconList.ImageSize.Width + 4;
            _resultsGrid.Columns["keyPath"].DefaultCellStyle.Padding = new Padding(iconPadding, 0, 0, 0);

            _resultsGrid.SelectionChanged += ResultsGrid_SelectionChanged;
            _resultsGrid.CellDoubleClick += ResultsGrid_DoubleClick;
            _resultsGrid.CellPainting += ResultsGrid_CellPainting;

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

            // SplitContainer for resizable results/preview layout
            // NOTE: Panel1MinSize and Panel2MinSize must NOT be set here — the container
            // is still at its default size (150x100) and the combined min sizes would
            // exceed that, causing InvalidOperationException. They are deferred to
            // VisibleChanged below where the container has its final layout dimensions.
            _splitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                BackColor = ModernTheme.Border,
                SplitterWidth = 3
            };
            _splitContainer.Panel1.BackColor = ModernTheme.Background;
            _splitContainer.Panel2.BackColor = ModernTheme.Background;

            // Panel1: results header + load more + grid
            _splitContainer.Panel1.Controls.Add(_resultsGrid);
            _splitContainer.Panel1.Controls.Add(_loadMorePanel);
            _splitContainer.Panel1.Controls.Add(resultsHeader);

            // Panel2: preview
            _splitContainer.Panel2.Controls.Add(previewPanel);

            // Set min sizes and initial splitter distance after the container is properly sized
            _splitContainer.VisibleChanged += (s, e) =>
            {
                if (_splitContainer.Visible && _splitContainer.Height > 0)
                {
                    try
                    {
                        _splitContainer.Panel1MinSize = DpiHelper.Scale(100);
                        _splitContainer.Panel2MinSize = DpiHelper.Scale(80);
                        _splitContainer.SplitterDistance = _splitContainer.Height * 3 / 4;
                    }
                    catch { }
                }
            };

            // Add controls
            this.Controls.Add(_splitContainer);
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
            _allMatches.Clear();
            _displayedMatchCount = 0;
            _loadMorePanel.Visible = false;
            _searchButton.Enabled = false;
            _searchButton.Visible = false;
            _cancelButton.Visible = true;
            _statusLabel.Text = "Searching...";
            _statusLabel.ForeColor = ModernTheme.Warning;
            _previewPathLabel.Text = "Searching...";
            _previewValueBox.Text = "";

            try
            {
                // Run search on background thread across all loaded parsers
                var wholeWord = _matchWholeWordCheckBox.Checked;
                foreach (var (parser, hiveTypeName) in _parsers)
                {
                    if (token.IsCancellationRequested) break;
                    var results = await Task.Run(() =>
                        parser.SearchAll(searchTerm, caseSensitive: false, wholeWord: wholeWord, cancellationToken: token), token).ConfigureAwait(true);

                    if (token.IsCancellationRequested) break;

                    foreach (var match in results)
                    {
                        _allMatches.Add((match, parser));
                    }
                }
                
                if (token.IsCancellationRequested) return;

                // Store all matches for pagination
                _displayedMatchCount = 0;

                // Display first page
                await DisplayNextPageAsync(token).ConfigureAwait(true);

                if (token.IsCancellationRequested)
                {
                    _statusLabel.Text = $"Search cancelled - {_searchResults.Count} results found";
                    _statusLabel.ForeColor = ModernTheme.Warning;
                }
                else
                {
                    UpdateStatusAndLoadMoreButton();
                }
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

        /// <summary>
        /// Display the next page of results from _allMatches into the grid.
        /// Each SearchMatch is already a value-level result from the backend.
        /// </summary>
        private Task DisplayNextPageAsync(CancellationToken token)
        {
            var searchTerm = _currentSearchTerm;

            // Use index-based access instead of Skip() for O(1) slicing
            int startIndex = _displayedMatchCount;
            int endIndex = Math.Min(startIndex + PageSize, _allMatches.Count);

            int count = 0;

            _resultsGrid.SuspendLayout();
            try
            {
                for (int i = startIndex; i < endIndex; i++)
                {
                    if (token.IsCancellationRequested) break;

                    var (match, matchParser) = _allMatches[i];
                    var searchResult = new SearchResult { KeyPath = match.Key.KeyPath, SourceParser = matchParser };

                    if (match.MatchKind == "Key")
                    {
                        // Key name match
                        searchResult.MatchType = "Key";
                        searchResult.Details = match.Key.KeyName;
                        searchResult.ValueData = "";
                        searchResult.FullValue = $"Key: {GetDisplayPath(match.Key.KeyPath, matchParser)}";
                    }
                    else if (match.MatchedValue != null)
                    {
                        // Value match (name or data)
                        var v = match.MatchedValue;
                        searchResult.MatchType = v.ValueType;
                        searchResult.ValueName = v.ValueName;
                        searchResult.ValueData = CleanValueData(v.ValueData?.ToString() ?? "");
                        searchResult.Details = string.IsNullOrEmpty(v.ValueName) ? "(Default)" : v.ValueName;
                        searchResult.FullValue = $"Name: {searchResult.Details}\nType: {v.ValueType}\nData: {searchResult.ValueData}";
                    }
                    else
                    {
                        // Fallback (should not normally happen)
                        searchResult.MatchType = "Key";
                        searchResult.Details = "";
                        searchResult.ValueData = "";
                        searchResult.FullValue = $"Key: {GetDisplayPath(match.Key.KeyPath, matchParser)}";
                    }

                    _searchResults.Add(searchResult);

                    // Truncate data for grid display (full data shown in preview panel)
                    var gridData = searchResult.ValueData.Length > 200
                        ? searchResult.ValueData.Substring(0, 200) + "..."
                        : searchResult.ValueData;

                    var rowIndex = _resultsGrid.Rows.Add(
                        GetDisplayPath(match.Key.KeyPath, matchParser),
                        searchResult.Details,
                        searchResult.MatchType,
                        gridData);
                    _resultsGrid.Rows[rowIndex].Tag = searchResult;
                    count++;

                    _displayedMatchCount++;
                }
            }
            finally
            {
                _resultsGrid.ResumeLayout();
            }

            if (count > 0)
            {
                _statusLabel.Text = $"Loading results... {_searchResults.Count} so far";
            }

            return Task.CompletedTask;
        }

        /// <summary>
        /// Handler for Load More button — appends next page of results.
        /// </summary>
        private async Task LoadMoreResultsAsync()
        {
            _loadMoreButton.Enabled = false;
            _statusLabel.Text = "Loading more results...";

            using var cts = new CancellationTokenSource();
            await DisplayNextPageAsync(cts.Token).ConfigureAwait(true);
            UpdateStatusAndLoadMoreButton();

            _loadMoreButton.Enabled = true;
        }

        /// <summary>
        /// Update the status label and Load More button visibility based on current pagination state.
        /// </summary>
        private void UpdateStatusAndLoadMoreButton()
        {
            var totalMatches = _allMatches.Count;
            var displayed = _searchResults.Count;
            var hasMore = _displayedMatchCount < totalMatches;

            if (displayed == 0)
            {
                _statusLabel.Text = "No results found";
                _statusLabel.ForeColor = ModernTheme.TextSecondary;
                _loadMorePanel.Visible = false;
            }
            else if (hasMore)
            {
                var remaining = totalMatches - _displayedMatchCount;
                var nextBatch = Math.Min(remaining, PageSize);
                _statusLabel.Text = $"Showing {displayed} of {totalMatches} results";
                _statusLabel.ForeColor = ModernTheme.Success;
                _loadMoreButton.Text = $"Load next {nextBatch}";
                _loadMorePanel.Visible = true;
            }
            else
            {
                _statusLabel.Text = totalMatches > PageSize
                    ? $"Showing all {displayed} results"
                    : $"Found {displayed} results";
                _statusLabel.ForeColor = ModernTheme.Success;
                _loadMorePanel.Visible = false;
            }

            _previewPathLabel.Text = displayed > 0 ? "Select a result to preview" : "No results found";
        }

        private void ResultsGrid_SelectionChanged(object? sender, EventArgs e)
        {
            if (_resultsGrid.SelectedRows.Count > 0 && _resultsGrid.SelectedRows[0].Tag is SearchResult result)
            {
                _previewPathLabel.Text = result.SourceParser != null ? GetDisplayPath(result.KeyPath, result.SourceParser) : result.KeyPath;
                SetPreviewWithHighlight(result.FullValue, _currentSearchTerm);
            }
        }

        private void ResultsGrid_CellPainting(object? sender, DataGridViewCellPaintingEventArgs e)
        {
            // Only draw icons in the "keyPath" column (index 0), skip header row
            if (e.ColumnIndex != 0 || e.RowIndex < 0 || _searchIconList == null)
                return;

            if (_resultsGrid.Rows[e.RowIndex].Tag is not SearchResult result)
                return;

            var matchType = result.MatchType;
            var imageKey = matchType == "Key" ? "folder" : ModernTheme.GetValueImageKey(matchType);
            var imgIndex = _searchIconList.Images.IndexOfKey(imageKey);
            if (imgIndex < 0) return;

            // Paint background and borders (let the grid handle these parts)
            e.PaintBackground(e.ClipBounds, true);

            var img = _searchIconList.Images[imgIndex];
            var iconX = e.CellBounds.X + 4;
            var iconY = e.CellBounds.Y + (e.CellBounds.Height - img.Height) / 2;
            e.Graphics.DrawImage(img, iconX, iconY, img.Width, img.Height);

            // Draw the text offset past the icon
            var textX = iconX + img.Width + 4;
            var textRect = new Rectangle(textX, e.CellBounds.Y, e.CellBounds.Right - textX, e.CellBounds.Height);
            var textColor = (e.State & DataGridViewElementStates.Selected) != 0
                ? e.CellStyle.SelectionForeColor
                : e.CellStyle.ForeColor;

            TextRenderer.DrawText(e.Graphics, e.FormattedValue?.ToString() ?? "", e.CellStyle.Font,
                textRect, textColor, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);

            e.Handled = true;
        }

        private void SetPreviewWithHighlight(string text, string searchTerm)
        {
            _previewValueBox.Text = text;
            
            if (string.IsNullOrEmpty(searchTerm) || string.IsNullOrEmpty(text))
                return;
            
            // Highlight all occurrences of the search term
            var comparison = StringComparison.OrdinalIgnoreCase;
            var wholeWord = _matchWholeWordCheckBox.Checked;
            int index = 0;
            while ((index = text.IndexOf(searchTerm, index, comparison)) >= 0)
            {
                // When whole-word matching, only highlight matches at word boundaries
                if (wholeWord)
                {
                    bool startBoundary = index == 0 || !char.IsLetterOrDigit(text[index - 1]);
                    int endPos = index + searchTerm.Length;
                    bool endBoundary = endPos >= text.Length || !char.IsLetterOrDigit(text[endPos]);
                    if (!startBoundary || !endBoundary)
                    {
                        index += searchTerm.Length;
                        continue;
                    }
                }
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
            if (_resultsGrid.SelectedRows.Count > 0 && _resultsGrid.SelectedRows[0].Tag is SearchResult result)
            {
                if (!string.IsNullOrEmpty(result.KeyPath))
                {
                    try
                    {
                        var displayPath = result.SourceParser != null ? GetDisplayPath(result.KeyPath, result.SourceParser) : result.KeyPath;
                        Clipboard.SetText(displayPath);
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
            if (_resultsGrid.SelectedRows.Count > 0 && _resultsGrid.SelectedRows[0].Tag is SearchResult result)
            {
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
            if (_resultsGrid.SelectedRows.Count > 0 && _resultsGrid.SelectedRows[0].Tag is SearchResult result)
            {
                
                // If the match is a value (not a key), pass the value name to auto-select it
                string? valueNameToSelect = null;
                if (result.MatchType != "Key" && !string.IsNullOrEmpty(result.ValueName))
                {
                    valueNameToSelect = result.ValueName;
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
            ModernTheme.ApplyWindowStyle(this);
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
            
            // Update split container colors
            _splitContainer.BackColor = ModernTheme.Border;
            _splitContainer.Panel1.BackColor = ModernTheme.Background;
            _splitContainer.Panel2.BackColor = ModernTheme.Background;

            // Refresh separators
            foreach (var sep in _separators)
                sep.BackColor = ModernTheme.Border;
            
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
                if (ctrl is SplitContainer split)
                {
                    split.BackColor = ModernTheme.Border;
                    split.Panel1.BackColor = ModernTheme.Background;
                    split.Panel2.BackColor = ModernTheme.Background;
                    RefreshControlTheme(split.Panel1);
                    RefreshControlTheme(split.Panel2);
                }
                else if (ctrl is Panel panel)
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
                    // Pill buttons are owner-drawn; just invalidate to repaint with current theme colors
                    button.Invalidate();
                }
                
                // Recurse into child controls
                if (ctrl.Controls.Count > 0)
                {
                    RefreshControlTheme(ctrl);
                }
            }
        }

        /// <summary>
        /// Create a pill-shaped button matching TimelineForm toolbar style.
        /// Transparent background with owner-drawn rounded-rect hover/press effects.
        /// </summary>
        private Button CreatePillButton(string text, EventHandler onClick, Color? textColor = null)
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

            var normalColor = textColor;

            btn.Paint += (s, e) =>
            {
                var g = e.Graphics;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                // Clear background to prevent ghosting
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

                // Draw centered text
                var displayText = btn.Text;
                var fgColor = isPressed ? Color.White : (normalColor ?? ModernTheme.TextPrimary);
                using var textBrush = new SolidBrush(fgColor);
                var textSize = g.MeasureString(displayText, ModernTheme.RegularFont);
                var textX = (btn.Width - textSize.Width) / 2;
                var textY = (btn.Height - textSize.Height) / 2;
                g.DrawString(displayText, ModernTheme.RegularFont, textBrush, textX, textY);
            };

            return btn;
        }

        /// <summary>
        /// Create a vertical separator for the toolbar, matching TimelineForm style.
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

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                var cts = _searchCts;
                _searchCts = null;
                try { cts?.Cancel(); } catch (ObjectDisposedException) { }
                cts?.Dispose();
                _searchIconList?.Dispose();
                _allMatches.Clear();
                _searchResults.Clear();
                _resultsGrid.Rows.Clear();
            }
            base.Dispose(disposing);
        }

        /// <summary>
        /// Convert a KeyPath from ROOT\... to HIVENAME\... for display
        /// </summary>
        private static string GetDisplayPath(string keyPath, OfflineRegistryParser parser) => parser.ConvertRootPath(keyPath);

        private class SearchResult
        {
            public string KeyPath { get; set; } = "";
            public string MatchType { get; set; } = "";
            public string Details { get; set; } = "";
            public string ValueName { get; set; } = "";
            public string ValueData { get; set; } = "";
            public string FullValue { get; set; } = "";
            public OfflineRegistryParser? SourceParser { get; set; }
        }
    }
}

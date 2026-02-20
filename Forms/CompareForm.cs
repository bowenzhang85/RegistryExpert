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
    /// <summary>
    /// Form for comparing two registry hives side-by-side
    /// User-friendly design with clear step-by-step flow
    /// </summary>
    public class CompareForm : Form
    {
        private OfflineRegistryParser? _leftParser;
        private OfflineRegistryParser? _rightParser;
        
        // Lookup dictionaries for fast path-based comparison (built sequentially, only read afterward)
        private Dictionary<string, RegistryKey> _leftKeysByPath = new(StringComparer.OrdinalIgnoreCase);
        private Dictionary<string, RegistryKey> _rightKeysByPath = new(StringComparer.OrdinalIgnoreCase);
        
        // Lookup dictionaries for O(1) node finding (instead of O(n) tree traversal)
        private Dictionary<string, TreeNode> _leftNodesByPath = new(StringComparer.OrdinalIgnoreCase);
        private Dictionary<string, TreeNode> _rightNodesByPath = new(StringComparer.OrdinalIgnoreCase);

        // Landing panels (shown before comparison)
        private Panel _leftLandingPanel = null!;
        private Panel _rightLandingPanel = null!;
        private Panel _centerButtonPanel = null!;
        private Button _leftLoadButton = null!;
        private Button _rightLoadButton = null!;
        private Label _leftStatusLabel = null!;
        private Label _rightStatusLabel = null!;
        private Label _leftFileNameLabel = null!;
        private Label _rightFileNameLabel = null!;
        private Button _compareButton = null!;

        // Comparison panels (shown after comparison)
        private Panel _leftComparePanel = null!;
        private Panel _rightComparePanel = null!;
        
        // Path display
        private TextBox _leftPathBox = null!;
        private TextBox _rightPathBox = null!;

        // Trees and grids
        private TreeView _leftTreeView = null!;
        private TreeView _rightTreeView = null!;
        private DataGridView _leftValuesGrid = null!;
        private DataGridView _rightValuesGrid = null!;

        // Main layout
        private SplitContainer _mainSplit = null!;
        private Panel _togglePanel = null!;
        private CheckBox _diffToggle = null!;

        // State
        private volatile bool _isSyncing = false;
        private bool _comparisonDone = false;
        private bool _isDisposed = false;
        private bool _showDifferencesOnly = true;
        private CancellationTokenSource? _cancellationTokenSource;
        private EventHandler? _themeChangedHandler;

        // Progress overlay
        private Panel _progressOverlay = null!;
        private Label _progressLabel = null!;
        private ProgressBar _progressBar = null!;
        private ImageList? _valueImageList; // Track for disposal (value type icons)

        public CompareForm()
        {
            InitializeComponent();
            ApplyTheme();
            
            // Subscribe to theme changes with a stored handler so we can unsubscribe
            _themeChangedHandler = (s, e) => ApplyTheme();
            ModernTheme.ThemeChanged += _themeChangedHandler;
            
            // Ensure proper cleanup when form is closing (before it closes)
            this.FormClosing += CompareForm_FormClosing;
        }

        private void CompareForm_FormClosing(object? sender, FormClosingEventArgs e)
        {
            // IMPORTANT: Hide form immediately so user sees instant close
            this.Visible = false;
            
            // Mark as disposed to prevent any further operations
            _isDisposed = true;
            
            // Cancel any ongoing operations first
            try
            {
                _cancellationTokenSource?.Cancel();
            }
            catch { }
            
            // Unsubscribe from theme changes first (must be on UI thread)
            if (_themeChangedHandler != null)
            {
                ModernTheme.ThemeChanged -= _themeChangedHandler;
                _themeChangedHandler = null;
            }
            
            // Clear UI controls on UI thread (safe)
            try
            {
                _leftTreeView?.Nodes.Clear();
                _rightTreeView?.Nodes.Clear();
                _leftValuesGrid?.Rows.Clear();
                _rightValuesGrid?.Rows.Clear();
            }
            catch { }
            
            // Clear dictionaries (thread-safe, no UI)
            _leftKeysByPath?.Clear();
            _rightKeysByPath?.Clear();
            _leftNodesByPath?.Clear();
            _rightNodesByPath?.Clear();
            
            // Capture references for background disposal
            var leftParser = _leftParser;
            var rightParser = _rightParser;
            var cts = _cancellationTokenSource;
            
            // Clear references immediately
            _leftParser = null;
            _rightParser = null;
            _cancellationTokenSource = null;
            
            // Dispose heavy resources in background (no UI access)
            Task.Run(() =>
            {
                try { cts?.Dispose(); } catch { }
                try { leftParser?.Dispose(); } catch { }
                try { rightParser?.Dispose(); } catch { }
            });
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
            _centerButtonPanel.Size = DpiHelper.ScaleSize(180, 60);
            _compareButton.Size = DpiHelper.ScaleSize(160, 50);
            _progressOverlay.Size = DpiHelper.ScaleSize(300, 150);
        }

        /// <summary>
        /// Pre-load a hive file as the left (base) hive
        /// </summary>
        public void SetLeftHive(string filePath)
        {
            LoadHiveFile(filePath, isLeft: true);
        }

        private void InitializeComponent()
        {
            // Enable DPI scaling (no AutoScaleDimensions - DpiHelper handles scaling)
            this.AutoScaleMode = AutoScaleMode.Dpi;

            this.Text = "Compare Registry Hives";
            this.Size = new Size(1400, 900);
            this.MinimumSize = new Size(1000, 700);
            this.StartPosition = FormStartPosition.CenterParent;
            this.BackColor = ModernTheme.Background;

            // Main split container
            _mainSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                BackColor = ModernTheme.Border,
                SplitterWidth = 4
            };

            // Create left side
            var leftContainer = new Panel { Dock = DockStyle.Fill, BackColor = ModernTheme.Background };
            _leftLandingPanel = CreateLandingPanel(isLeft: true, out _leftLoadButton, out _leftStatusLabel, out _leftFileNameLabel);
            _leftComparePanel = CreateComparePanel(isLeft: true, out _leftPathBox, out _leftTreeView, out _leftValuesGrid);
            _leftComparePanel.Visible = false;
            leftContainer.Controls.Add(_leftComparePanel);
            leftContainer.Controls.Add(_leftLandingPanel);

            // Create right side
            var rightContainer = new Panel { Dock = DockStyle.Fill, BackColor = ModernTheme.Background };
            _rightLandingPanel = CreateLandingPanel(isLeft: false, out _rightLoadButton, out _rightStatusLabel, out _rightFileNameLabel);
            _rightComparePanel = CreateComparePanel(isLeft: false, out _rightPathBox, out _rightTreeView, out _rightValuesGrid);
            _rightComparePanel.Visible = false;
            rightContainer.Controls.Add(_rightComparePanel);
            rightContainer.Controls.Add(_rightLandingPanel);

            _mainSplit.Panel1.Controls.Add(leftContainer);
            _mainSplit.Panel2.Controls.Add(rightContainer);

            // Create center Compare button panel (overlays the split)
            _centerButtonPanel = new Panel
            {
                Size = DpiHelper.ScaleSize(180, 60),
                BackColor = Color.Transparent,
                Visible = false
            };

            _compareButton = new Button
            {
                Text = "Compare Hives",
                Size = DpiHelper.ScaleSize(160, 50),
                Location = DpiHelper.ScalePoint(10, 5),
                FlatStyle = FlatStyle.Flat,
                BackColor = ModernTheme.Success,
                ForeColor = Color.White,
                Font = new Font(ModernTheme.RegularFont.FontFamily, 12F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            _compareButton.FlatAppearance.BorderSize = 0;
            _compareButton.FlatAppearance.MouseOverBackColor = Color.FromArgb(40, 180, 100);
            _compareButton.Click += CompareButton_Click;
            _centerButtonPanel.Controls.Add(_compareButton);

            // Create filter toggle panel (shown after comparison)
            _togglePanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = DpiHelper.Scale(36),
                BackColor = ModernTheme.Surface,
                Visible = false
            };

            _diffToggle = new CheckBox
            {
                Text = "Show Differences Only",
                Checked = true,
                FlatStyle = FlatStyle.Standard,
                ForeColor = ModernTheme.TextPrimary,
                BackColor = ModernTheme.Surface,
                AutoSize = true,
                Font = new Font("Segoe UI", 9F),
                Dock = DockStyle.Left
            };
            _diffToggle.CheckedChanged += DiffToggle_CheckedChanged;
            _togglePanel.Controls.Add(_diffToggle);

            this.Controls.Add(_centerButtonPanel);
            this.Controls.Add(_mainSplit);
            this.Controls.Add(_togglePanel);

            // Bring center button to front
            _centerButtonPanel.BringToFront();

            // Position center button and splitter on load and resize
            this.Load += (s, e) =>
            {
                CenterCompareButton();
                CenterSplitter();
            };

            this.Resize += (s, e) =>
            {
                CenterCompareButton();
                CenterSplitter();
                CenterProgressOverlay();
            };

            // Create progress overlay (initially hidden)
            CreateProgressOverlay();

            // Initial state
            UpdateUI();
        }

        private void CreateProgressOverlay()
        {
            _progressOverlay = new Panel
            {
                Visible = false,
                BackColor = Color.FromArgb(200, 30, 30, 30)  // Semi-transparent dark
            };

            // Center container
            var centerContainer = new Panel
            {
                Size = DpiHelper.ScaleSize(350, 130),
                BackColor = ModernTheme.Surface,
                Padding = new Padding(20, 15, 20, 15)
            };

            // Status label
            _progressLabel = new Label
            {
                Text = "Preparing comparison...",
                Font = new Font(ModernTheme.RegularFont.FontFamily, 12F),
                ForeColor = ModernTheme.TextPrimary,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Top,
                Height = DpiHelper.Scale(45),
                AutoSize = false
            };

            // Marquee progress bar
            _progressBar = new ProgressBar
            {
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 30,
                Height = DpiHelper.Scale(25),
                Dock = DockStyle.Bottom
            };

            centerContainer.Controls.Add(_progressBar);
            centerContainer.Controls.Add(_progressLabel);

            _progressOverlay.Controls.Add(centerContainer);

            // Center the container when overlay resizes
            _progressOverlay.Resize += (s, e) =>
            {
                centerContainer.Location = new Point(
                    (_progressOverlay.Width - centerContainer.Width) / 2,
                    (_progressOverlay.Height - centerContainer.Height) / 2
                );
            };

            this.Controls.Add(_progressOverlay);
            _progressOverlay.BringToFront();
        }

        private void CenterProgressOverlay()
        {
            if (_progressOverlay != null)
            {
                _progressOverlay.Size = this.ClientSize;
                _progressOverlay.Location = Point.Empty;
            }
        }

        private void CenterSplitter()
        {
            if (_mainSplit.Width > 0)
            {
                _mainSplit.SplitterDistance = _mainSplit.Width / 2;
            }
        }

        private void CenterCompareButton()
        {
            // Position compare button exactly in the center of the form
            _centerButtonPanel.Location = new Point(
                (this.ClientSize.Width - _centerButtonPanel.Width) / 2,
                (this.ClientSize.Height - _centerButtonPanel.Height) / 2
            );
        }

        private Panel CreateLandingPanel(bool isLeft, out Button loadButton, out Label statusLabel, out Label fileNameLabel)
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ModernTheme.Background
            };

            // Center container
            var centerPanel = new Panel
            {
                Size = DpiHelper.ScaleSize(350, 300),
                BackColor = ModernTheme.Background
            };

            // Step indicator with circle
            var stepPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = DpiHelper.Scale(60),
                BackColor = ModernTheme.Background
            };
            
            var stepNumber = isLeft ? "1" : "2";
            stepPanel.Paint += (s, e) =>
            {
                var g = e.Graphics;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                
                // Draw circle
                int circleSize = 44;
                int circleX = (stepPanel.Width - circleSize) / 2;
                int circleY = 5;
                
                using var circlePen = new Pen(ModernTheme.Accent, 3);
                g.DrawEllipse(circlePen, circleX, circleY, circleSize, circleSize);
                
                // Draw number inside circle
                using var font = new Font(ModernTheme.RegularFont.FontFamily, 18F, FontStyle.Bold);
                using var brush = new SolidBrush(ModernTheme.Accent);
                var textSize = g.MeasureString(stepNumber, font);
                float textX = circleX + (circleSize - textSize.Width) / 2;
                float textY = circleY + (circleSize - textSize.Height) / 2;
                g.DrawString(stepNumber, font, brush, textX, textY);
            };

            // Title
            var titleLabel = new Label
            {
                Text = isLeft ? "Load First Hive" : "Load Second Hive",
                Font = new Font(ModernTheme.RegularFont.FontFamily, 16F, FontStyle.Bold),
                ForeColor = ModernTheme.TextPrimary,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Top,
                Height = DpiHelper.Scale(40)
            };

            // Description
            var descLabel = new Label
            {
                Text = isLeft 
                    ? "Select the base registry hive file to compare from" 
                    : "Select the second registry hive file to compare against",
                Font = ModernTheme.RegularFont,
                ForeColor = ModernTheme.TextSecondary,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Top,
                Height = DpiHelper.Scale(50)
            };

            // Large load button
            loadButton = new Button
            {
                Text = "Load Hive File",
                Size = DpiHelper.ScaleSize(200, 50),
                FlatStyle = FlatStyle.Flat,
                BackColor = ModernTheme.Accent,
                ForeColor = Color.White,
                Font = new Font(ModernTheme.RegularFont.FontFamily, 12F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            loadButton.FlatAppearance.BorderSize = 0;
            loadButton.FlatAppearance.MouseOverBackColor = ModernTheme.AccentHover;
            loadButton.Click += (s, e) => LoadHive(isLeft);

            // Button container for centering
            var buttonPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = DpiHelper.Scale(70)
            };

            // File name label (shown after loading)
            fileNameLabel = new Label
            {
                Text = "",
                Font = new Font(ModernTheme.RegularFont.FontFamily, 10F),
                ForeColor = ModernTheme.Success,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Top,
                Height = DpiHelper.Scale(30),
                Visible = false
            };

            // Status label
            statusLabel = new Label
            {
                Text = "",
                Font = ModernTheme.RegularFont,
                ForeColor = ModernTheme.TextSecondary,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Top,
                Height = DpiHelper.Scale(30)
            };

            // Position button in center using fixed calculation
            // Note: Do NOT use a Resize handler here — during PerMonitorV2 DPI scaling,
            // the parent panel is scaled before child controls, causing the Resize handler
            // to compute positions with stale child dimensions, which are then double-scaled.
            loadButton.Location = new Point(DpiHelper.Scale(75), DpiHelper.Scale(10));
            buttonPanel.Controls.Add(loadButton);

            // Add controls
            centerPanel.Controls.Add(statusLabel);
            centerPanel.Controls.Add(fileNameLabel);
            centerPanel.Controls.Add(buttonPanel);
            centerPanel.Controls.Add(descLabel);
            centerPanel.Controls.Add(titleLabel);
            centerPanel.Controls.Add(stepPanel);

            // Center the panel
            panel.Resize += (s, e) =>
            {
                centerPanel.Location = new Point(
                    (panel.Width - centerPanel.Width) / 2,
                    (panel.Height - centerPanel.Height) / 2
                );
            };

            panel.Controls.Add(centerPanel);
            return panel;
        }

        private Panel CreateComparePanel(bool isLeft, out TextBox pathBox, out TreeView treeView, out DataGridView valuesGrid)
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ModernTheme.Background
            };

            // Header with file info and back button
            var headerPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = DpiHelper.Scale(36),
                BackColor = ModernTheme.Surface
            };

            var backButton = new Button
            {
                Text = "\uE72B",  // Back arrow
                Font = new Font("Segoe MDL2 Assets", 10F),
                FlatStyle = FlatStyle.Flat,
                BackColor = ModernTheme.Surface,
                ForeColor = ModernTheme.TextPrimary,
                Size = new Size(32, 32),
                Location = new Point(2, 2),
                Cursor = Cursors.Hand
            };
            backButton.FlatAppearance.BorderSize = 0;
            backButton.FlatAppearance.MouseOverBackColor = ModernTheme.SurfaceHover;
            backButton.Click += (s, e) => ResetComparison();

            var fileLabel = new Label
            {
                Text = isLeft ? "Left Hive" : "Right Hive",
                Font = ModernTheme.BoldFont,
                ForeColor = ModernTheme.TextPrimary,
                Location = new Point(40, 8),
                AutoSize = true
            };

            headerPanel.Controls.Add(backButton);
            headerPanel.Controls.Add(fileLabel);

            // Path display
            pathBox = new TextBox
            {
                Dock = DockStyle.Top,
                Height = DpiHelper.Scale(26),
                ReadOnly = true,
                BackColor = ModernTheme.Surface,
                ForeColor = ModernTheme.TextSecondary,
                BorderStyle = BorderStyle.None,
                Font = ModernTheme.SmallFont,
                Text = ""
            };

            // Split container for tree and values
            var splitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                BackColor = ModernTheme.Border,
                SplitterWidth = 4,
                SplitterDistance = 400
            };

            // Tree view
            treeView = new TreeView
            {
                Dock = DockStyle.Fill,
                DrawMode = TreeViewDrawMode.OwnerDrawText,
                HideSelection = false,
                ShowLines = true,
                ShowPlusMinus = true,
                ShowRootLines = true,
                Font = ModernTheme.DataFont
            };
            ModernTheme.ApplyTo(treeView);
            treeView.DrawNode += TreeView_DrawNode;

            if (isLeft)
            {
                treeView.AfterSelect += LeftTreeView_AfterSelect;
                treeView.AfterExpand += LeftTreeView_AfterExpand;
                treeView.AfterCollapse += LeftTreeView_AfterCollapse;
            }
            else
            {
                treeView.AfterSelect += RightTreeView_AfterSelect;
                treeView.AfterExpand += RightTreeView_AfterExpand;
                treeView.AfterCollapse += RightTreeView_AfterCollapse;
            }

            // Values grid
            valuesGrid = CreateValuesGrid();

            splitContainer.Panel1.Controls.Add(treeView);
            splitContainer.Panel2.Controls.Add(valuesGrid);

            panel.Controls.Add(splitContainer);
            panel.Controls.Add(pathBox);
            panel.Controls.Add(headerPanel);

            return panel;
        }

        private DataGridView CreateValuesGrid()
        {
            var grid = new DataGridView { Dock = DockStyle.Fill };
            ModernTheme.ApplyTo(grid);
            
            // CompareForm-specific overrides
            grid.MultiSelect = false;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            grid.BackgroundColor = ModernTheme.Background;
            grid.DefaultCellStyle.BackColor = ModernTheme.Background;

            grid.Columns.Add("Name", "Name");
            grid.Columns.Add("Type", "Type");
            grid.Columns.Add("Value", "Value");

            grid.Columns["Name"].Width = 150;
            grid.Columns["Type"].Width = 100;
            grid.Columns["Value"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            // Draw value type icons in the Name column
            if (_valueImageList == null)
                _valueImageList = CreateValueImageList();

            grid.CellPainting += (s, ev) =>
            {
                if (ev.RowIndex < 0 || ev.ColumnIndex != 0 || _valueImageList == null || ev.Graphics == null)
                    return;

                var typeValue = grid.Rows[ev.RowIndex].Cells[1].Value?.ToString() ?? "";
                var imageKey = GetValueImageKey(typeValue);
                var imgIndex = _valueImageList.Images.IndexOfKey(imageKey);
                if (imgIndex < 0) return;

                ev.PaintBackground(ev.ClipBounds, true);

                var img = _valueImageList.Images[imgIndex];
                var iconY = ev.CellBounds.Y + (ev.CellBounds.Height - img.Height) / 2;
                var iconX = ev.CellBounds.X + 4;
                ev.Graphics.DrawImage(img, iconX, iconY, img.Width, img.Height);

                var textX = iconX + img.Width + 4;
                var textRect = new Rectangle(textX, ev.CellBounds.Y, ev.CellBounds.Width - (textX - ev.CellBounds.X), ev.CellBounds.Height);
                var textColor = ev.CellStyle?.ForeColor ?? ModernTheme.TextPrimary;
                if (ev.State.HasFlag(DataGridViewElementStates.Selected))
                    textColor = ev.CellStyle?.SelectionForeColor ?? ModernTheme.TextPrimary;

                TextRenderer.DrawText(ev.Graphics, ev.Value?.ToString() ?? "", ev.CellStyle?.Font ?? ModernTheme.DataFont,
                    textRect, textColor, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);

                ev.Handled = true;
            };

            return grid;
        }

        private ImageList CreateValueImageList() => ModernTheme.CreateValueImageList();

        private static string GetValueImageKey(string? valueType) => ModernTheme.GetValueImageKey(valueType);

        private void LoadHive(bool isLeft)
        {
            using var dialog = new OpenFileDialog
            {
                Title = isLeft ? "Open First Registry Hive" : "Open Second Registry Hive",
                Filter = "All Files|*.*|Registry Hives|NTUSER.DAT;SAM;SECURITY;SOFTWARE;SYSTEM;USRCLASS.DAT;DEFAULT;Amcache.hve;BCD",
                FilterIndex = 1
            };

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                LoadHiveFile(dialog.FileName, isLeft);
            }
        }

        private void LoadHiveFile(string filePath, bool isLeft)
        {
            try
            {
                var parser = new OfflineRegistryParser();
                parser.LoadHive(filePath);

                if (isLeft)
                {
                    _leftParser?.Dispose();
                    _leftParser = parser;
                    _leftFileNameLabel.Text = $"{parser.CurrentHiveType}: {System.IO.Path.GetFileName(filePath)}";
                    _leftFileNameLabel.Visible = true;
                    _leftStatusLabel.Text = "Loaded successfully";
                    _leftStatusLabel.ForeColor = ModernTheme.Success;
                    _leftLoadButton.Text = "Change File";
                    _leftLoadButton.BackColor = ModernTheme.Surface;
                    _leftLoadButton.ForeColor = ModernTheme.TextPrimary;
                    _leftLoadButton.FlatAppearance.BorderColor = ModernTheme.Border;
                    _leftLoadButton.FlatAppearance.BorderSize = 1;
                }
                else
                {
                    // Validate same hive type
                    if (_leftParser != null && _leftParser.CurrentHiveType != parser.CurrentHiveType)
                    {
                        parser.Dispose();
                        MessageBox.Show(
                            $"Hive type mismatch!\n\nFirst hive: {_leftParser.CurrentHiveType}\nSecond hive: {parser.CurrentHiveType}\n\nBoth hives must be the same type for comparison.",
                            "Type Mismatch",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        return;
                    }

                    _rightParser?.Dispose();
                    _rightParser = parser;
                    _rightFileNameLabel.Text = $"{parser.CurrentHiveType}: {System.IO.Path.GetFileName(filePath)}";
                    _rightFileNameLabel.Visible = true;
                    _rightStatusLabel.Text = "Loaded successfully";
                    _rightStatusLabel.ForeColor = ModernTheme.Success;
                    _rightLoadButton.Text = "Change File";
                    _rightLoadButton.BackColor = ModernTheme.Surface;
                    _rightLoadButton.ForeColor = ModernTheme.TextPrimary;
                    _rightLoadButton.FlatAppearance.BorderColor = ModernTheme.Border;
                    _rightLoadButton.FlatAppearance.BorderSize = 1;
                }

                UpdateUI();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading hive:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateUI()
        {
            // Update right panel prompts based on left panel state
            if (_leftParser == null)
            {
                _rightLoadButton.Enabled = false;
                _rightStatusLabel.Text = "Load the first hive first";
                _rightStatusLabel.ForeColor = ModernTheme.TextSecondary;
            }
            else
            {
                _rightLoadButton.Enabled = true;
                if (_rightParser == null)
                {
                    _rightStatusLabel.Text = $"Select a {_leftParser.CurrentHiveType} hive to compare";
                    _rightStatusLabel.ForeColor = ModernTheme.TextSecondary;
                }
            }

            // Show compare button when both are loaded
            _centerButtonPanel.Visible = _leftParser != null && _rightParser != null && !_comparisonDone;
        }

        private void DiffToggle_CheckedChanged(object? sender, EventArgs e)
        {
            _showDifferencesOnly = _diffToggle.Checked;
            RebuildFilteredTrees();
        }

        private async void CompareButton_Click(object? sender, EventArgs e)
        {
            if (_leftParser == null || _rightParser == null)
                return;

            // Properly dispose old CancellationTokenSource before creating new one
            _cancellationTokenSource?.Cancel();
            _cancellationTokenSource?.Dispose();
            _cancellationTokenSource = new CancellationTokenSource();
            var token = _cancellationTokenSource.Token;

            _compareButton.Enabled = false;
            _compareButton.Text = "Comparing...";

            // Show progress overlay
            CenterProgressOverlay();
            _progressLabel.Text = "Building key index...";
            _progressOverlay.Visible = true;
            _progressOverlay.BringToFront();
            _progressOverlay.Refresh();  // Force immediate repaint
            
            // Small delay to ensure overlay is rendered before heavy work
            await Task.Delay(50, token).ConfigureAwait(true);

            try
            {
                // Capture parser references to avoid null issues if form closes during operation
                var leftParser = _leftParser;
                var rightParser = _rightParser;
                
                if (leftParser == null || rightParser == null || token.IsCancellationRequested)
                    return;

                // Build lookup dictionaries
                await Task.Run(() =>
                {
                    if (token.IsCancellationRequested) return;
                    _leftKeysByPath = BuildKeyIndex(leftParser.GetRootKey(), "");
                    if (token.IsCancellationRequested) return;
                    _rightKeysByPath = BuildKeyIndex(rightParser.GetRootKey(), "");
                }, token).ConfigureAwait(true);

                if (token.IsCancellationRequested) return;

                // Update progress
                _progressLabel.Text = "Analyzing differences...";
                _progressLabel.Refresh();  // Force immediate repaint
                await Task.Delay(10, token).ConfigureAwait(true);

                if (token.IsCancellationRequested) return;

                // Populate trees (run on background thread to keep UI responsive)
                TreeNode? leftRootNode = null;
                TreeNode? rightRootNode = null;
                
                await Task.Run(() =>
                {
                    if (token.IsCancellationRequested) return;
                    var leftRoot = leftParser.GetRootKey();
                    var rightRoot = rightParser.GetRootKey();
                    
                    if (leftRoot != null && !token.IsCancellationRequested)
                        leftRootNode = CreateTreeNode(leftRoot, "", isLeftTree: true);
                    if (rightRoot != null && !token.IsCancellationRequested)
                        rightRootNode = CreateTreeNode(rightRoot, "", isLeftTree: false);
                }, token).ConfigureAwait(true);

                if (token.IsCancellationRequested) return;

                // Update progress
                _progressLabel.Text = "Building tree view...";
                _progressLabel.Refresh();  // Force immediate repaint

                // Add nodes to tree views on UI thread
                _leftTreeView.BeginUpdate();
                _rightTreeView.BeginUpdate();
                
                _leftTreeView.Nodes.Clear();
                _rightTreeView.Nodes.Clear();
                
                if (leftRootNode != null)
                {
                    _leftTreeView.Nodes.Add(leftRootNode);
                    leftRootNode.Expand();
                }
                if (rightRootNode != null)
                {
                    _rightTreeView.Nodes.Add(rightRootNode);
                    rightRootNode.Expand();
                }
                
                _leftTreeView.EndUpdate();
                _rightTreeView.EndUpdate();

                // Apply initial filter (toggle defaults to ON)
                if (_showDifferencesOnly)
                {
                    if (leftRootNode != null) FilterNodeChildren(leftRootNode, _leftNodesByPath);
                    if (rightRootNode != null) FilterNodeChildren(rightRootNode, _rightNodesByPath);
                }

                // Switch to comparison view
                _comparisonDone = true;
                _centerButtonPanel.Visible = false;
                _leftLandingPanel.Visible = false;
                _rightLandingPanel.Visible = false;
                _leftComparePanel.Visible = true;
                _rightComparePanel.Visible = true;
                _togglePanel.Visible = true;
            }
            catch (OperationCanceledException)
            {
                // Cancelled - don't show error
            }
            catch (Exception ex)
            {
                if (!token.IsCancellationRequested)
                    MessageBox.Show($"Error during comparison:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _progressOverlay.Visible = false;
                _compareButton.Enabled = true;
                _compareButton.Text = "Compare Hives";
            }
        }

        private void ResetComparison()
        {
            _comparisonDone = false;
            _leftComparePanel.Visible = false;
            _rightComparePanel.Visible = false;
            _leftLandingPanel.Visible = true;
            _rightLandingPanel.Visible = true;
            _togglePanel.Visible = false;
            _showDifferencesOnly = true;
            _diffToggle.Checked = true;
            
            _leftTreeView.Nodes.Clear();
            _rightTreeView.Nodes.Clear();
            _leftValuesGrid.Rows.Clear();
            _rightValuesGrid.Rows.Clear();
            _leftPathBox.Text = "";
            _rightPathBox.Text = "";
            
            // Clear node lookup dictionaries
            _leftNodesByPath.Clear();
            _rightNodesByPath.Clear();

            UpdateUI();
        }

        private Dictionary<string, RegistryKey> BuildKeyIndex(RegistryKey? key, string parentPath)
        {
            var result = new Dictionary<string, RegistryKey>(StringComparer.OrdinalIgnoreCase);
            if (key == null) return result;

            // Use non-recursive approach with stack to avoid dictionary merging overhead
            BuildKeyIndexRecursive(key, parentPath, result);
            return result;
        }
        
        private void BuildKeyIndexRecursive(RegistryKey key, string parentPath, Dictionary<string, RegistryKey> result)
        {
            var path = string.IsNullOrEmpty(parentPath) ? key.KeyName : $"{parentPath}\\{key.KeyName}";
            result[path] = key;

            if (key.SubKeys != null)
            {
                foreach (var subKey in key.SubKeys)
                {
                    BuildKeyIndexRecursive(subKey, path, result);
                }
            }
        }

        private TreeNode? CreateTreeNode(RegistryKey key, string parentPath, bool isLeftTree)
        {
            var path = string.IsNullOrEmpty(parentPath) ? key.KeyName : $"{parentPath}\\{key.KeyName}";
            
            var otherIndex = isLeftTree ? _rightKeysByPath : _leftKeysByPath;
            
            bool existsInOther = otherIndex.ContainsKey(path);
            bool hasValueDifferences = false;
            
            if (existsInOther)
            {
                hasValueDifferences = HasValueDifferences(key, otherIndex[path]);
            }

            var nodeTag = new NodeTag 
            { 
                Path = path, 
                Key = key, 
                HasDifference = false,
                IsUniqueToThisHive = false,
                HasValueDifference = false
            };
            var node = new TreeNode(key.KeyName)
            {
                Tag = nodeTag
            };

            // Add child nodes first and track their difference states
            bool childHasDifference = false;
            bool allChildrenAreUnique = true;  // Track if ALL children are 100% unique
            
            if (key.SubKeys != null)
            {
                foreach (var subKey in key.SubKeys.OrderBy(k => k.KeyName, StringComparer.OrdinalIgnoreCase))
                {
                    var childNode = CreateTreeNode(subKey, path, isLeftTree);
                    if (childNode != null)
                    {
                        node.Nodes.Add(childNode);
                        var childTag = childNode.Tag as NodeTag;
                        if (childTag?.HasDifference == true)
                        {
                            childHasDifference = true;
                            // If any child is NOT purely unique (has value diff or mixed), then not all children are unique
                            if (!childTag.IsUniqueToThisHive || childTag.HasValueDifference)
                            {
                                allChildrenAreUnique = false;
                            }
                        }
                    }
                }
            }

            // Determine the color for this node:
            // GREEN = This key doesn't exist in other hive AND all children are also 100% unique (or no children)
            // RED = Any other difference (value diff, key exists but children differ, mixed situation)
            
            bool thisKeyIsUnique = !existsInOther;
            bool hasDifference = thisKeyIsUnique || hasValueDifferences || childHasDifference;
            
            if (hasDifference)
            {
                nodeTag.HasDifference = true;
                
                if (thisKeyIsUnique && !childHasDifference)
                {
                    // Key is unique and has no children with differences -> GREEN
                    nodeTag.IsUniqueToThisHive = true;
                }
                else if (thisKeyIsUnique && childHasDifference && allChildrenAreUnique)
                {
                    // Key is unique AND all children are also purely unique -> GREEN
                    nodeTag.IsUniqueToThisHive = true;
                }
                else
                {
                    // All other cases -> RED:
                    // - Key exists in both but has value differences
                    // - Key is unique but some children have value differences (mixed)
                    // - Key exists in both and children have differences
                    nodeTag.HasValueDifference = true;
                }
            }

            // Register node in lookup dictionary for O(1) finding
            var nodeLookup = isLeftTree ? _leftNodesByPath : _rightNodesByPath;
            nodeLookup[path] = node;

            return node;
        }

        private bool HasValueDifferences(RegistryKey left, RegistryKey right)
        {
            var leftValues = left.Values ?? new List<KeyValue>();
            var rightValues = right.Values ?? new List<KeyValue>();

            if (leftValues.Count != rightValues.Count)
                return true;

            var leftDict = new Dictionary<string, KeyValue>(StringComparer.OrdinalIgnoreCase);
            foreach (var v in leftValues)
            {
                var name = v.ValueName ?? "(Default)";
                if (!leftDict.ContainsKey(name))
                    leftDict[name] = v;
            }

            foreach (var rv in rightValues)
            {
                var name = rv.ValueName ?? "(Default)";
                if (!leftDict.TryGetValue(name, out var lv))
                    return true;

                if (lv.ValueType != rv.ValueType)
                    return true;

                var leftData = lv.ValueData?.ToString() ?? "";
                var rightData = rv.ValueData?.ToString() ?? "";
                if (leftData != rightData)
                    return true;
            }

            return false;
        }

        private void TreeView_DrawNode(object? sender, DrawTreeNodeEventArgs e)
        {
            if (e.Node == null) return;

            var bounds = e.Bounds;
            var tag = e.Node.Tag as NodeTag;
            
            // Extend bounds to the full width of the TreeView to prevent text clipping
            if (e.Node.TreeView != null)
            {
                bounds = new Rectangle(bounds.X, bounds.Y, e.Node.TreeView.ClientSize.Width - bounds.X, bounds.Height);
            }
            
            // Determine color based on difference type:
            // GREEN = unique to this hive (doesn't exist in other)
            // RED = exists in both but has different values
            // Normal = identical in both hives
            Color foreColor;
            if (tag?.HasDifference == true)
            {
                if (tag.IsUniqueToThisHive)
                    foreColor = ModernTheme.DiffAdded;  // GREEN - unique to this hive
                else if (tag.HasValueDifference)
                    foreColor = ModernTheme.DiffRemoved;  // RED - value differences
                else
                    foreColor = ModernTheme.DiffRemoved;  // Default to RED for any other difference
            }
            else
            {
                foreColor = ModernTheme.TextPrimary;  // Normal - no difference
            }

            if ((e.State & TreeNodeStates.Selected) != 0)
            {
                using var selBrush = new SolidBrush(ModernTheme.Selection);
                e.Graphics.FillRectangle(selBrush, bounds);
            }

            TextRenderer.DrawText(e.Graphics, e.Node.Text, ModernTheme.DataFont, bounds, foreColor,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
        }

        private void LeftTreeView_AfterSelect(object? sender, TreeViewEventArgs e)
        {
            if (_isSyncing || e.Node == null) return;

            _isSyncing = true;
            try
            {
                var tag = e.Node.Tag as NodeTag;
                if (tag != null)
                {
                    _leftPathBox.Text = tag.Path;
                    ShowValues(_leftValuesGrid, tag.Key, tag.Path, isLeftGrid: true);
                    
                    var rightNode = FindNodeByPath(_rightTreeView, tag.Path);
                    if (rightNode != null)
                    {
                        _rightTreeView.SelectedNode = rightNode;
                        rightNode.EnsureVisible();
                        
                        var rightTag = rightNode.Tag as NodeTag;
                        if (rightTag != null)
                        {
                            _rightPathBox.Text = rightTag.Path;
                            ShowValues(_rightValuesGrid, rightTag.Key, rightTag.Path, isLeftGrid: false);
                        }
                    }
                    else
                    {
                        _rightPathBox.Text = tag.Path + " (not found)";
                        _rightValuesGrid.Rows.Clear();
                    }
                }
            }
            finally
            {
                _isSyncing = false;
            }
        }

        private void RightTreeView_AfterSelect(object? sender, TreeViewEventArgs e)
        {
            if (_isSyncing || e.Node == null) return;

            _isSyncing = true;
            try
            {
                var tag = e.Node.Tag as NodeTag;
                if (tag != null)
                {
                    _rightPathBox.Text = tag.Path;
                    ShowValues(_rightValuesGrid, tag.Key, tag.Path, isLeftGrid: false);
                    
                    var leftNode = FindNodeByPath(_leftTreeView, tag.Path);
                    if (leftNode != null)
                    {
                        _leftTreeView.SelectedNode = leftNode;
                        leftNode.EnsureVisible();
                        
                        var leftTag = leftNode.Tag as NodeTag;
                        if (leftTag != null)
                        {
                            _leftPathBox.Text = leftTag.Path;
                            ShowValues(_leftValuesGrid, leftTag.Key, leftTag.Path, isLeftGrid: true);
                        }
                    }
                    else
                    {
                        _leftPathBox.Text = tag.Path + " (not found)";
                        _leftValuesGrid.Rows.Clear();
                    }
                }
            }
            finally
            {
                _isSyncing = false;
            }
        }

        private void ShowValues(DataGridView grid, RegistryKey key, string path, bool isLeftGrid)
        {
            grid.Rows.Clear();
            
            if (key.Values == null || key.Values.Count == 0)
                return;

            // Use the path directly instead of slow GetPathForKey lookup
            Dictionary<string, KeyValue>? otherValues = null;
            
            var otherIndex = isLeftGrid ? _rightKeysByPath : _leftKeysByPath;
            if (otherIndex.TryGetValue(path, out var otherKey) && otherKey.Values != null)
            {
                otherValues = new Dictionary<string, KeyValue>(StringComparer.OrdinalIgnoreCase);
                foreach (var v in otherKey.Values)
                {
                    var name = v.ValueName ?? "(Default)";
                    if (!otherValues.ContainsKey(name))
                        otherValues[name] = v;
                }
            }

            // Show values - color-coded by diff status (filter applied if differences-only mode is active)
            foreach (var value in key.Values.OrderBy(v => v.ValueName ?? "(Default)", StringComparer.OrdinalIgnoreCase))
            {
                var name = value.ValueName ?? "(Default)";
                var type = value.ValueType ?? "";
                var data = FormatValue(value);

                bool isUnique = false;      // GREEN - only exists in this hive
                bool isValueDiff = false;   // RED - exists in both but different
                
                if (otherValues == null || !otherValues.ContainsKey(name))
                {
                    // Value only exists in this hive - GREEN
                    isUnique = true;
                }
                else
                {
                    var otherValue = otherValues[name];
                    var otherData = FormatValue(otherValue);
                    if (type != otherValue.ValueType || data != otherData)
                    {
                        // Value exists in both but is different - RED
                        isValueDiff = true;
                    }
                }
                
                // In differences-only mode, skip identical values
                if (_showDifferencesOnly && !isUnique && !isValueDiff)
                    continue;

                var rowIndex = grid.Rows.Add(name, type, data);
                if (isUnique)
                    grid.Rows[rowIndex].DefaultCellStyle.ForeColor = ModernTheme.DiffAdded;    // GREEN - unique to this hive
                else if (isValueDiff)
                    grid.Rows[rowIndex].DefaultCellStyle.ForeColor = ModernTheme.DiffRemoved;  // RED - value differs
                // else: identical value — keep default TextPrimary color
            }
        }

        private string FormatValue(KeyValue value)
        {
            if (value.ValueData == null)
                return "(null)";

            try
            {
                var data = value.ValueData.ToString() ?? "";
                
                if (value.ValueType?.ToUpperInvariant() == "REGBINARY" && value.ValueDataRaw != null && value.ValueDataRaw.Length > 0)
                {
                    var hex = BitConverter.ToString(value.ValueDataRaw.Take(32).ToArray()).Replace("-", " ");
                    if (value.ValueDataRaw.Length > 32)
                        hex += $"... ({value.ValueDataRaw.Length} bytes)";
                    return hex;
                }

                if (data.Length > 200)
                    return data.Substring(0, 200) + "...";

                return data;
            }
            catch
            {
                return value.ValueData?.ToString() ?? "(error)";
            }
        }

        private TreeNode? FindNodeByPath(TreeView tree, string path)
        {
            // O(1) lookup using dictionary
            var nodeLookup = tree == _leftTreeView ? _leftNodesByPath : _rightNodesByPath;
            return nodeLookup.TryGetValue(path, out var node) ? node : null;
        }

        private void LeftTreeView_AfterExpand(object? sender, TreeViewEventArgs e)
        {
            SyncExpand(e.Node, _rightTreeView);
        }

        private void LeftTreeView_AfterCollapse(object? sender, TreeViewEventArgs e)
        {
            SyncCollapse(e.Node, _rightTreeView);
        }

        private void RightTreeView_AfterExpand(object? sender, TreeViewEventArgs e)
        {
            SyncExpand(e.Node, _leftTreeView);
        }

        private void RightTreeView_AfterCollapse(object? sender, TreeViewEventArgs e)
        {
            SyncCollapse(e.Node, _leftTreeView);
        }

        private void SyncExpand(TreeNode? sourceNode, TreeView targetTree)
        {
            if (_isSyncing || sourceNode == null) return;

            _isSyncing = true;
            try
            {
                var tag = sourceNode.Tag as NodeTag;
                if (tag != null)
                {
                    var targetNode = FindNodeByPath(targetTree, tag.Path);
                    targetNode?.Expand();
                }
            }
            finally
            {
                _isSyncing = false;
            }
        }

        private void SyncCollapse(TreeNode? sourceNode, TreeView targetTree)
        {
            if (_isSyncing || sourceNode == null) return;

            _isSyncing = true;
            try
            {
                var tag = sourceNode.Tag as NodeTag;
                if (tag != null)
                {
                    var targetNode = FindNodeByPath(targetTree, tag.Path);
                    targetNode?.Collapse();
                }
            }
            finally
            {
                _isSyncing = false;
            }
        }

        private void ApplyTheme()
        {
            this.BackColor = ModernTheme.Background;
            
            _leftPathBox.BackColor = ModernTheme.Surface;
            _leftPathBox.ForeColor = ModernTheme.TextSecondary;
            _rightPathBox.BackColor = ModernTheme.Surface;
            _rightPathBox.ForeColor = ModernTheme.TextSecondary;

            _leftTreeView.BackColor = ModernTheme.Background;
            _rightTreeView.BackColor = ModernTheme.Background;

            _togglePanel.BackColor = ModernTheme.Surface;
            _diffToggle.ForeColor = ModernTheme.TextPrimary;
            _diffToggle.BackColor = ModernTheme.Surface;

            ApplyThemeToGrid(_leftValuesGrid);
            ApplyThemeToGrid(_rightValuesGrid);
        }

        private void ApplyThemeToGrid(DataGridView grid)
        {
            grid.BackgroundColor = ModernTheme.Background;
            grid.GridColor = ModernTheme.Border;
            grid.DefaultCellStyle.BackColor = ModernTheme.Background;
            grid.DefaultCellStyle.ForeColor = ModernTheme.TextPrimary;
            grid.DefaultCellStyle.SelectionBackColor = ModernTheme.Selection;
            grid.DefaultCellStyle.SelectionForeColor = ModernTheme.TextPrimary;
            grid.ColumnHeadersDefaultCellStyle.BackColor = ModernTheme.Surface;
            grid.ColumnHeadersDefaultCellStyle.ForeColor = ModernTheme.TextPrimary;
            grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = ModernTheme.Surface;
        }

        private void RebuildFilteredTrees()
        {
            if (!_comparisonDone) return;

            // 1. Capture expanded paths before rebuild
            var leftExpanded = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var rightExpanded = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            CollectExpandedPaths(_leftTreeView.Nodes, leftExpanded);
            CollectExpandedPaths(_rightTreeView.Nodes, rightExpanded);

            // 2. Capture selected paths before rebuild
            string? leftSelectedPath = (_leftTreeView.SelectedNode?.Tag as NodeTag)?.Path;
            string? rightSelectedPath = (_rightTreeView.SelectedNode?.Tag as NodeTag)?.Path;

            // 3. Clear lookup dictionaries (CRITICAL — stale entries break FindNodeByPath)
            _leftNodesByPath.Clear();
            _rightNodesByPath.Clear();

            // 4. Rebuild trees from scratch
            _leftTreeView.BeginUpdate();
            _rightTreeView.BeginUpdate();
            _leftTreeView.Nodes.Clear();
            _rightTreeView.Nodes.Clear();

            var leftRoot = _leftParser?.GetRootKey();
            var rightRoot = _rightParser?.GetRootKey();

            TreeNode? leftRootNode = leftRoot != null ? CreateTreeNode(leftRoot, "", isLeftTree: true) : null;
            TreeNode? rightRootNode = rightRoot != null ? CreateTreeNode(rightRoot, "", isLeftTree: false) : null;

            // 5. Apply filter if differences-only mode
            if (_showDifferencesOnly)
            {
                if (leftRootNode != null) FilterNodeChildren(leftRootNode, _leftNodesByPath);
                if (rightRootNode != null) FilterNodeChildren(rightRootNode, _rightNodesByPath);
            }

            if (leftRootNode != null)
            {
                _leftTreeView.Nodes.Add(leftRootNode);
                leftRootNode.Expand();
            }
            if (rightRootNode != null)
            {
                _rightTreeView.Nodes.Add(rightRootNode);
                rightRootNode.Expand();
            }

            _leftTreeView.EndUpdate();
            _rightTreeView.EndUpdate();

            // 6. Restore expanded state
            RestoreExpandedPaths(_leftTreeView.Nodes, leftExpanded);
            RestoreExpandedPaths(_rightTreeView.Nodes, rightExpanded);

            // 7. Restore selection (fall back to root if path no longer exists after filtering)
            string? restorePath = leftSelectedPath ?? rightSelectedPath;
            TreeNode? leftNode = restorePath != null ? FindNodeByPath(_leftTreeView, restorePath) : null;
            if (leftNode == null && _leftTreeView.Nodes.Count > 0) leftNode = _leftTreeView.Nodes[0];
            TreeNode? rightNode = restorePath != null ? FindNodeByPath(_rightTreeView, restorePath) : null;
            if (rightNode == null && _rightTreeView.Nodes.Count > 0) rightNode = _rightTreeView.Nodes[0];

            _isSyncing = true;
            try
            {
                if (leftNode != null) { _leftTreeView.SelectedNode = leftNode; leftNode.EnsureVisible(); }
                if (rightNode != null) { _rightTreeView.SelectedNode = rightNode; rightNode.EnsureVisible(); }
            }
            finally
            {
                _isSyncing = false;
            }

            // 8. Refresh grids for restored selection
            if (leftNode?.Tag is NodeTag lt)
            {
                _leftPathBox.Text = lt.Path;
                ShowValues(_leftValuesGrid, lt.Key, lt.Path, isLeftGrid: true);
            }
            if (rightNode?.Tag is NodeTag rt)
            {
                _rightPathBox.Text = rt.Path;
                ShowValues(_rightValuesGrid, rt.Key, rt.Path, isLeftGrid: false);
            }
        }

        private void CollectExpandedPaths(TreeNodeCollection nodes, HashSet<string> paths)
        {
            foreach (TreeNode node in nodes)
            {
                if (node.IsExpanded)
                {
                    if (node.Tag is NodeTag tag) paths.Add(tag.Path);
                    CollectExpandedPaths(node.Nodes, paths);
                }
            }
        }

        private void FilterNodeChildren(TreeNode node, Dictionary<string, TreeNode> nodeLookup)
        {
            // Prune children where HasDifference == false (safe: means no diff in this node OR any descendant)
            for (int i = node.Nodes.Count - 1; i >= 0; i--)
            {
                var child = node.Nodes[i];
                var childTag = child.Tag as NodeTag;
                if (childTag == null || !childTag.HasDifference)
                {
                    if (childTag != null) RemoveFromLookup(child, nodeLookup);
                    node.Nodes.RemoveAt(i);
                }
                else
                {
                    FilterNodeChildren(child, nodeLookup);
                }
            }
        }

        private void RemoveFromLookup(TreeNode node, Dictionary<string, TreeNode> nodeLookup)
        {
            if (node.Tag is NodeTag tag) nodeLookup.Remove(tag.Path);
            foreach (TreeNode child in node.Nodes)
                RemoveFromLookup(child, nodeLookup);
        }

        private void RestoreExpandedPaths(TreeNodeCollection nodes, HashSet<string> expandedPaths)
        {
            foreach (TreeNode node in nodes)
            {
                if (node.Tag is NodeTag tag && expandedPaths.Contains(tag.Path))
                {
                    node.Expand();
                    RestoreExpandedPaths(node.Nodes, expandedPaths);
                }
            }
        }

        protected override void Dispose(bool disposing)
        {
            // Most cleanup is already done in FormClosing (async)
            // This is just a safety net for any remaining resources
            if (_isDisposed)
            {
                base.Dispose(disposing);
                return;
            }
                 
            _isDisposed = true;
            
            if (disposing)
            {
                // Unsubscribe from theme changes (safety net)
                if (_themeChangedHandler != null)
                {
                    ModernTheme.ThemeChanged -= _themeChangedHandler;
                    _themeChangedHandler = null;
                }
                
                // Note: Don't clear nodes here - it's slow and already done in FormClosing
                // Just let base.Dispose handle the control disposal
                _valueImageList?.Dispose();
            }
            base.Dispose(disposing);
        }

        private class NodeTag
        {
            public string Path { get; set; } = "";
            public RegistryKey Key { get; set; } = null!;
            public bool HasDifference { get; set; } = false;
            public bool IsUniqueToThisHive { get; set; } = false;  // GREEN - only exists in this hive
            public bool HasValueDifference { get; set; } = false;   // RED - exists in both but values differ
        }
    }
}

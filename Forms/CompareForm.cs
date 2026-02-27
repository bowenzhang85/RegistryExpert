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
        private Dictionary<string, RegistryKey>? _leftKeysByPath = new(StringComparer.OrdinalIgnoreCase);
        private Dictionary<string, RegistryKey>? _rightKeysByPath = new(StringComparer.OrdinalIgnoreCase);
        
        // Lookup dictionaries for O(1) node finding (instead of O(n) tree traversal)
        private Dictionary<string, TreeNode>? _leftNodesByPath = new(StringComparer.OrdinalIgnoreCase);
        private Dictionary<string, TreeNode>? _rightNodesByPath = new(StringComparer.OrdinalIgnoreCase);

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
        private CheckBox _diffOnlyCheckbox = null!;

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

        // State
        private volatile bool _isSyncing = false;
        private bool _comparisonDone = false;
        private bool _isDisposed = false;
        private bool _showDifferencesOnly = false;
        private CancellationTokenSource? _cancellationTokenSource;
        private EventHandler? _themeChangedHandler;

        // Progress overlay
        private Panel _progressOverlay = null!;
        private Label _progressLabel = null!;
        private ProgressBar _progressBar = null!;
        private ImageList? _valueImageList; // Track for disposal (value type icons)
        
        // Pre-computed diff status for every key path (lightweight — no Win32 controls)
        private Dictionary<string, DiffInfo>? _leftDiffByPath;
        private Dictionary<string, DiffInfo>? _rightDiffByPath;

        // Status bar (bottom of form - progress during hive loading)
        private Panel _statusPanel = null!;
        private Panel _statusRightPanel = null!;
        private string _statusText = "Ready";
        private Color _statusForeColor;
        private ProgressBar _loadProgressBar = null!;
        private Panel _progressWrapper = null!;
        private Button _cancelLoadButton = null!;
        private Panel _cancelWrapper = null!;
        private CancellationTokenSource? _loadCts;

        public CompareForm()
        {
            InitializeComponent();
            ApplyTheme();
            
            // Subscribe to theme changes with a stored handler so we can unsubscribe
            // Theme changes fire on the main UI thread, but this form may run on its own
            // STA thread, so marshal the call to this form's thread.
            _themeChangedHandler = (s, e) =>
            {
                if (_isDisposed) return;
                try
                {
                    if (this.IsHandleCreated && this.InvokeRequired)
                        this.BeginInvoke(() => { if (!_isDisposed) ApplyTheme(); });
                    else if (!_isDisposed)
                        ApplyTheme();
                }
                catch (ObjectDisposedException) { }
            };
            ModernTheme.ThemeChanged += _themeChangedHandler;
            
            // Ensure proper cleanup when form is closing (before it closes)
            this.FormClosing += CompareForm_FormClosing;
        }

        private void CompareForm_FormClosing(object? sender, FormClosingEventArgs e)
        {
            // Mark as disposed to prevent any further operations
            _isDisposed = true;
            
            // Cancel any ongoing operations first
            try
            {
                _cancellationTokenSource?.Cancel();
                _loadCts?.Cancel();
            }
            catch { }
            
            // Unsubscribe from theme changes (must be on UI thread)
            if (_themeChangedHandler != null)
            {
                ModernTheme.ThemeChanged -= _themeChangedHandler;
                _themeChangedHandler = null;
            }
            
            // Capture references for background disposal
            var leftParser = _leftParser;
            var rightParser = _rightParser;
            var cts = _cancellationTokenSource;
            var loadCts = _loadCts;
            var leftKeys = _leftKeysByPath;
            var rightKeys = _rightKeysByPath;
            var leftNodes = _leftNodesByPath;
            var rightNodes = _rightNodesByPath;
            var leftDiff = _leftDiffByPath;
            var rightDiff = _rightDiffByPath;
            
            // Clear references immediately
            _leftParser = null;
            _rightParser = null;
            _cancellationTokenSource = null;
            _loadCts = null;
            _leftKeysByPath = null;
            _rightKeysByPath = null;
            _leftNodesByPath = null;
            _rightNodesByPath = null;
            _leftDiffByPath = null;
            _rightDiffByPath = null;
            
            // Dispose heavy resources in background (no UI access)
            Task.Run(() =>
            {
                try { cts?.Dispose(); } catch { }
                try { loadCts?.Dispose(); } catch { }
                try { leftParser?.Dispose(); } catch { }
                try { rightParser?.Dispose(); } catch { }
                leftKeys?.Clear();
                rightKeys?.Clear();
                leftNodes?.Clear();
                rightNodes?.Clear();
                leftDiff?.Clear();
                rightDiff?.Clear();
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
            _centerButtonPanel.Size = DpiHelper.ScaleSize(180, 95);
            _compareButton.Size = DpiHelper.ScaleSize(160, 50);
            _progressOverlay.Size = DpiHelper.ScaleSize(300, 150);
        }

        /// <summary>
        /// Pre-load a hive file as the left (base) hive
        /// </summary>
        public void SetLeftHive(string filePath)
        {
            _ = LoadHiveFileAsync(filePath, isLeft: true);
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
                Size = DpiHelper.ScaleSize(180, 95),
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

            _diffOnlyCheckbox = new CheckBox
            {
                Text = "Show differences only",
                Checked = false,
                ForeColor = ModernTheme.TextPrimary,
                BackColor = Color.Transparent,
                Font = new Font("Segoe UI", 9F),
                AutoSize = true,
                Location = DpiHelper.ScalePoint(18, 60),
                Cursor = Cursors.Hand
            };
            _centerButtonPanel.Controls.Add(_diffOnlyCheckbox);

            // Status bar (bottom of form) — matches MainForm status bar design
            CreateStatusBar();

            this.Controls.Add(_centerButtonPanel);
            this.Controls.Add(_mainSplit);
            this.Controls.Add(_statusPanel); // Dock.Bottom added after Dock.Fill so it sits below

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

        private void CreateStatusBar()
        {
            _statusForeColor = ModernTheme.TextSecondary;
            
            _statusPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = DpiHelper.Scale(32),
                BackColor = ModernTheme.Surface,
                Padding = DpiHelper.ScalePadding(10, 0, 10, 0)
            };

            // Owner-draw status text: enable double-buffering and paint handler
            typeof(Panel).InvokeMember("DoubleBuffered",
                System.Reflection.BindingFlags.SetProperty | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
                null, _statusPanel, new object[] { true });
            _statusPanel.Paint += StatusPanel_Paint;
            _statusPanel.AccessibleRole = AccessibleRole.StatusBar;
            _statusPanel.AccessibleName = _statusText;

            // Fixed-width right panel
            _statusRightPanel = new Panel
            {
                Dock = DockStyle.Right,
                Width = DpiHelper.Scale(170),
                BackColor = Color.Transparent,
            };

            // Progress bar inside a wrapper for vertical centering
            _loadProgressBar = new ProgressBar
            {
                Minimum = 0,
                Maximum = 100,
                Value = 0,
                Dock = DockStyle.Fill,
                Style = ProgressBarStyle.Continuous,
            };
            _progressWrapper = new Panel
            {
                Dock = DockStyle.Right,
                Width = DpiHelper.Scale(100),
                Padding = new Padding(0, DpiHelper.Scale(8), 0, DpiHelper.Scale(8)),
                Visible = false,
                BackColor = Color.Transparent,
            };
            _progressWrapper.Controls.Add(_loadProgressBar);

            // Cancel button inside a wrapper for vertical centering
            _cancelLoadButton = new Button
            {
                Text = "Cancel",
                AccessibleName = "Cancel Loading",
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.Transparent,
                ForeColor = ModernTheme.TextSecondary,
                Font = ModernTheme.SmallFont,
                Cursor = Cursors.Hand,
                Dock = DockStyle.Fill,
            };
            _cancelLoadButton.FlatAppearance.BorderColor = ModernTheme.Border;
            _cancelLoadButton.FlatAppearance.BorderSize = 1;
            _cancelLoadButton.FlatAppearance.MouseOverBackColor = ModernTheme.Selection;
            _cancelLoadButton.FlatAppearance.MouseDownBackColor = ModernTheme.SurfaceLight;
            _cancelLoadButton.Click += CancelLoad_Click;
            _cancelWrapper = new Panel
            {
                Dock = DockStyle.Right,
                Width = DpiHelper.Scale(68),
                Padding = new Padding(DpiHelper.Scale(4), DpiHelper.Scale(4), DpiHelper.Scale(4), DpiHelper.Scale(4)),
                Visible = false,
                BackColor = Color.Transparent,
            };
            _cancelWrapper.Controls.Add(_cancelLoadButton);

            // Dock.Right controls added after Dock.Fill (higher z-order, docked first)
            _statusRightPanel.Controls.Add(_cancelWrapper);
            _statusRightPanel.Controls.Add(_progressWrapper);

            _statusPanel.Controls.Add(_statusRightPanel);
        }

        private void CancelLoad_Click(object? sender, EventArgs e)
        {
            _loadCts?.Cancel();
            _cancelLoadButton.Enabled = false;
            SetStatusText("Cancelling...");
        }

        private void SetStatusText(string text, Color? color = null)
        {
            _statusText = text;
            _statusPanel.AccessibleName = text;
            if (color.HasValue)
                _statusForeColor = color.Value;
            _statusPanel.Invalidate();
        }

        private void StatusPanel_Paint(object? sender, PaintEventArgs e)
        {
            var padding = _statusPanel.Padding;
            var textRect = new Rectangle(
                padding.Left,
                padding.Top,
                _statusPanel.ClientSize.Width - _statusRightPanel.Width - padding.Left - padding.Right,
                _statusPanel.ClientSize.Height - padding.Top - padding.Bottom);

            TextRenderer.DrawText(
                e.Graphics,
                _statusText,
                ModernTheme.RegularFont,
                textRect,
                _statusForeColor,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPadding);
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
                Cursor = Cursors.Hand,
                AccessibleName = "Back"
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
                Text = "",
                AccessibleName = isLeft ? "Left Hive Path" : "Right Hive Path"
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
                Font = ModernTheme.DataFont,
                AccessibleName = isLeft ? "Left Hive Keys" : "Right Hive Keys"
            };
            ModernTheme.ApplyTo(treeView);
            treeView.DrawNode += TreeView_DrawNode;
            treeView.BeforeExpand += TreeView_BeforeExpand;

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
            valuesGrid = CreateValuesGrid(isLeft ? "Left" : "Right");

            splitContainer.Panel1.Controls.Add(treeView);
            splitContainer.Panel2.Controls.Add(valuesGrid);

            panel.Controls.Add(splitContainer);
            panel.Controls.Add(pathBox);
            panel.Controls.Add(headerPanel);

            return panel;
        }

        private DataGridView CreateValuesGrid(string side)
        {
            var grid = new DataGridView { Dock = DockStyle.Fill, AccessibleName = $"{side} Hive Values" };
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

        private async void LoadHive(bool isLeft)
        {
            using var dialog = new OpenFileDialog
            {
                Title = isLeft ? "Open First Registry Hive" : "Open Second Registry Hive",
                Filter = "All Files|*.*|Registry Hives|NTUSER.DAT;SAM;SECURITY;SOFTWARE;SYSTEM;USRCLASS.DAT;DEFAULT;Amcache.hve;BCD",
                FilterIndex = 1
            };

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                await LoadHiveFileAsync(dialog.FileName, isLeft);
            }
        }

        private async Task LoadHiveFileAsync(string filePath, bool isLeft)
        {
            var side = isLeft ? "left" : "right";
            try
            {
                // Cancel any previous load operation
                _loadCts?.Cancel();
                _loadCts?.Dispose();
                _loadCts = new CancellationTokenSource();
                var token = _loadCts.Token;

                // Show progress UI
                SetStatusText($"Loading {side} hive...", ModernTheme.Warning);
                _loadProgressBar.Value = 0;
                _progressWrapper.Visible = true;
                _cancelWrapper.Visible = true;
                _cancelLoadButton.Enabled = true;

                // Disable load/compare buttons during loading
                _leftLoadButton.Enabled = false;
                _rightLoadButton.Enabled = false;
                _compareButton.Enabled = false;

                var parser = new OfflineRegistryParser();

                var lastPhase = string.Empty;
                var stageNumber = 0;

                var progress = new Progress<(string phase, double percent)>(update =>
                {
                    if (update.phase != lastPhase)
                    {
                        lastPhase = update.phase;
                        stageNumber++;
                    }
                    var phase = update.phase.TrimEnd('.');
                    var newText = $"Stage {stageNumber}/2: {phase} {update.percent:P0}";
                    if (_statusText != newText)
                        SetStatusText(newText);
                    var newValue = Math.Clamp((int)(update.percent * 100), 0, 100);
                    if (_loadProgressBar.Value != newValue)
                        _loadProgressBar.Value = newValue;
                });

                await Task.Run(() => parser.LoadHive(filePath, progress, token), token).ConfigureAwait(true);

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
                        SetStatusText("Hive type mismatch", ModernTheme.Warning);
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

                SetStatusText($"Loaded: {System.IO.Path.GetFileName(filePath)}", ModernTheme.Success);
                UpdateUI();
            }
            catch (OperationCanceledException)
            {
                SetStatusText("Loading cancelled", ModernTheme.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading hive:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                SetStatusText($"Failed to load {side} hive", ModernTheme.Error);
            }
            finally
            {
                _progressWrapper.Visible = false;
                _cancelWrapper.Visible = false;

                // Re-enable buttons
                _leftLoadButton.Enabled = true;
                _compareButton.Enabled = true;
                // Right load button state depends on left parser
                UpdateUI();
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

        private async void CompareButton_Click(object? sender, EventArgs e)
        {
            if (_leftParser == null || _rightParser == null)
                return;

            _showDifferencesOnly = _diffOnlyCheckbox.Checked;

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
                _progressLabel.Refresh();
                await Task.Delay(10, token).ConfigureAwait(true);

                if (token.IsCancellationRequested) return;

                // Pre-compute diff status for every key path (lightweight — no Win32 controls).
                // This replaces the old recursive CreateTreeNode which built 500K+ TreeNodes.
                var leftKeys = _leftKeysByPath;
                var rightKeys = _rightKeysByPath;
                
                await Task.Run(() =>
                {
                    if (token.IsCancellationRequested || leftKeys == null || rightKeys == null) return;
                    
                    _leftDiffByPath = new Dictionary<string, DiffInfo>(StringComparer.OrdinalIgnoreCase);
                    _rightDiffByPath = new Dictionary<string, DiffInfo>(StringComparer.OrdinalIgnoreCase);
                    
                    var leftRoot = leftParser.GetRootKey();
                    var rightRoot = rightParser.GetRootKey();
                    
                    if (leftRoot != null && !token.IsCancellationRequested)
                        ComputeDiffRecursive(leftRoot, "", rightKeys, _leftDiffByPath, token);
                    if (rightRoot != null && !token.IsCancellationRequested)
                        ComputeDiffRecursive(rightRoot, "", leftKeys, _rightDiffByPath, token);
                }, token).ConfigureAwait(true);

                if (token.IsCancellationRequested) return;

                // Create root nodes lazily (just root + first-level children with dummies).
                // With lazy-load, adding nodes is instant — no rendering delay.
                _leftNodesByPath = new Dictionary<string, TreeNode>(StringComparer.OrdinalIgnoreCase);
                _rightNodesByPath = new Dictionary<string, TreeNode>(StringComparer.OrdinalIgnoreCase);
                
                TreeNode? leftRootNode = null;
                TreeNode? rightRootNode = null;
                
                var leftRootKey = leftParser.GetRootKey();
                var rightRootKey = rightParser.GetRootKey();
                
                if (leftRootKey != null)
                    leftRootNode = CreateLazyNode(leftRootKey, leftRootKey.KeyName, isLeftTree: true);
                if (rightRootKey != null)
                    rightRootNode = CreateLazyNode(rightRootKey, rightRootKey.KeyName, isLeftTree: false);

                // Add nodes to tree views on UI thread
                _leftTreeView.BeginUpdate();
                _rightTreeView.BeginUpdate();
                
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

                // Switch to comparison view
                _comparisonDone = true;
                _centerButtonPanel.Visible = false;
                _leftLandingPanel.Visible = false;
                _rightLandingPanel.Visible = false;
                _leftComparePanel.Visible = true;
                _rightComparePanel.Visible = true;
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
                _progressBar.Style = ProgressBarStyle.Marquee;
                _progressBar.Value = 0;
                _compareButton.Enabled = true;
                _compareButton.Text = "Compare Hives";
            }
        }

        private void ResetComparison()
        {
            _comparisonDone = false;
            _showDifferencesOnly = false;
            _diffOnlyCheckbox.Checked = false;
            _leftComparePanel.Visible = false;
            _rightComparePanel.Visible = false;
            _leftLandingPanel.Visible = true;
            _rightLandingPanel.Visible = true;
            
            // Replace tree views instead of calling Nodes.Clear() — clearing sends
            // individual TVM_DELETEITEM messages per node and freezes the UI for 15-20s
            // on large hives. Replacing the control is instant; old ones are disposed
            // asynchronously via BeginInvoke so the heavy DestroyWindow happens later.
            ReplaceTreeView(ref _leftTreeView, isLeft: true);
            ReplaceTreeView(ref _rightTreeView, isLeft: false);
            
            _leftValuesGrid.Rows.Clear();
            _rightValuesGrid.Rows.Clear();
            _leftPathBox.Text = "";
            _rightPathBox.Text = "";
            
            // Clear node lookup dictionaries
            _leftNodesByPath?.Clear();
            _rightNodesByPath?.Clear();
            _leftDiffByPath?.Clear();
            _rightDiffByPath?.Clear();

            UpdateUI();
        }

        /// <summary>
        /// Replace a TreeView with a fresh empty one.
        /// With lazy-load, trees have very few nodes so disposal is instant.
        /// </summary>
        private void ReplaceTreeView(ref TreeView treeView, bool isLeft)
        {
            var parent = treeView.Parent;
            if (parent == null) return;
            
            var oldTree = treeView;
            
            // Create replacement with identical configuration
            var newTree = new TreeView
            {
                Dock = DockStyle.Fill,
                DrawMode = TreeViewDrawMode.OwnerDrawText,
                HideSelection = false,
                ShowLines = true,
                ShowPlusMinus = true,
                ShowRootLines = true,
                Font = ModernTheme.DataFont,
                AccessibleName = isLeft ? "Left Hive Keys" : "Right Hive Keys"
            };
            ModernTheme.ApplyTo(newTree);
            newTree.DrawNode += TreeView_DrawNode;
            newTree.BeforeExpand += TreeView_BeforeExpand;
            
            if (isLeft)
            {
                newTree.AfterSelect += LeftTreeView_AfterSelect;
                newTree.AfterExpand += LeftTreeView_AfterExpand;
                newTree.AfterCollapse += LeftTreeView_AfterCollapse;
            }
            else
            {
                newTree.AfterSelect += RightTreeView_AfterSelect;
                newTree.AfterExpand += RightTreeView_AfterExpand;
                newTree.AfterCollapse += RightTreeView_AfterCollapse;
            }
            
            // Swap: add new, remove old
            parent.SuspendLayout();
            parent.Controls.Add(newTree);
            parent.Controls.Remove(oldTree);
            parent.ResumeLayout();
            
            treeView = newTree;
            
            // With lazy-load, trees have very few nodes — dispose is instant
            oldTree.Dispose();
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

        /// <summary>
        /// Pre-compute diff status for every key path in a hive.
        /// Mirrors the GREEN/RED logic from the old recursive CreateTreeNode,
        /// but only walks RegistryKey objects — no TreeNodes or Win32 controls.
        /// </summary>
        private void ComputeDiffRecursive(RegistryKey key, string parentPath,
            Dictionary<string, RegistryKey> otherIndex,
            Dictionary<string, DiffInfo> result,
            CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            
            var path = string.IsNullOrEmpty(parentPath) ? key.KeyName : $"{parentPath}\\{key.KeyName}";
            
            bool isUnique = !otherIndex.ContainsKey(path);
            bool hasValueDiff = false;
            
            if (!isUnique && otherIndex.TryGetValue(path, out var otherKey))
                hasValueDiff = HasValueDifferences(key, otherKey);
            
            bool anyChildHasDiff = false;
            bool allChildrenUnique = true;
            
            if (key.SubKeys != null)
            {
                foreach (var sub in key.SubKeys)
                {
                    ComputeDiffRecursive(sub, path, otherIndex, result, token);
                    var childPath = $"{path}\\{sub.KeyName}";
                    if (result.TryGetValue(childPath, out var childDiff) && childDiff.HasDifference)
                    {
                        anyChildHasDiff = true;
                        if (!childDiff.IsUniqueToThisHive)
                            allChildrenUnique = false;
                    }
                }
            }
            
            bool hasDiff = isUnique || hasValueDiff || anyChildHasDiff;
            bool nodeIsUnique = false;
            bool nodeHasValueDiff = false;
            
            if (hasDiff)
            {
                if (isUnique && (!anyChildHasDiff || allChildrenUnique))
                    nodeIsUnique = true;
                else
                    nodeHasValueDiff = true;
            }
            
            result[path] = new DiffInfo(hasDiff, nodeIsUnique, nodeHasValueDiff);
        }

        /// <summary>
        /// Create a single tree node using pre-computed diff info.
        /// Non-recursive — children are populated lazily via BeforeExpand.
        /// If the key has visible children, a dummy sentinel node (Tag == null)
        /// is added to show the [+] expand indicator.
        /// </summary>
        private TreeNode? CreateLazyNode(RegistryKey key, string path, bool isLeftTree)
        {
            var diffIndex = isLeftTree ? _leftDiffByPath : _rightDiffByPath;
            var nodeIndex = isLeftTree ? _leftNodesByPath : _rightNodesByPath;
            
            // Get pre-computed diff info
            DiffInfo diffInfo = default;
            diffIndex?.TryGetValue(path, out diffInfo);
            
            // In differences-only mode, skip nodes with no differences
            if (_showDifferencesOnly && !diffInfo.HasDifference)
                return null;
            
            var nodeTag = new NodeTag
            {
                Path = path,
                Key = key,
                HasDifference = diffInfo.HasDifference,
                IsUniqueToThisHive = diffInfo.IsUniqueToThisHive,
                HasValueDifference = diffInfo.HasValueDifference
            };
            var node = new TreeNode(key.KeyName) { Tag = nodeTag };
            
            // Add dummy child if this node has expandable children
            if (HasVisibleChildren(key, path, isLeftTree))
                node.Nodes.Add(new TreeNode()); // Dummy sentinel (Tag == null)
            
            // Register in lookup dictionary for O(1) finding
            if (nodeIndex != null)
                nodeIndex[path] = node;
            
            return node;
        }

        /// <summary>
        /// Check if a key has any children that should be shown in the tree.
        /// In normal mode, any subkeys count. In differences-only mode,
        /// only subkeys with differences (or descendant differences) count.
        /// </summary>
        private bool HasVisibleChildren(RegistryKey key, string parentPath, bool isLeftTree)
        {
            if (key.SubKeys == null || key.SubKeys.Count == 0)
                return false;
            
            if (!_showDifferencesOnly)
                return true;
            
            var diffIndex = isLeftTree ? _leftDiffByPath : _rightDiffByPath;
            if (diffIndex == null) return false;
            
            foreach (var sub in key.SubKeys)
            {
                var subPath = $"{parentPath}\\{sub.KeyName}";
                if (diffIndex.TryGetValue(subPath, out var diff) && diff.HasDifference)
                    return true;
            }
            return false;
        }

        /// <summary>
        /// Create child nodes for a node being expanded.
        /// Called from BeforeExpand handler when lazy sentinel is detected.
        /// </summary>
        private void PopulateChildren(TreeNode parentNode, bool isLeftTree)
        {
            var tag = parentNode.Tag as NodeTag;
            if (tag == null) return;
            
            var key = tag.Key;
            var parentPath = tag.Path;
            
            if (key.SubKeys == null) return;
            
            foreach (var subKey in key.SubKeys.OrderBy(k => k.KeyName, StringComparer.OrdinalIgnoreCase))
            {
                var path = $"{parentPath}\\{subKey.KeyName}";
                var childNode = CreateLazyNode(subKey, path, isLeftTree);
                if (childNode != null)
                    parentNode.Nodes.Add(childNode);
            }
        }

        private void TreeView_BeforeExpand(object? sender, TreeViewCancelEventArgs e)
        {
            if (e.Node == null) return;
            
            // Check for dummy sentinel (Tag == null on first child)
            if (e.Node.Nodes.Count != 1 || e.Node.Nodes[0].Tag != null)
                return; // Already populated
            
            bool isLeft = (sender == _leftTreeView);
            
            var tree = sender as TreeView;
            tree?.BeginUpdate();
            try
            {
                e.Node.Nodes.Clear(); // Remove dummy (1 node, instant)
                PopulateChildren(e.Node, isLeft);
            }
            finally
            {
                tree?.EndUpdate();
            }
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
            if (otherIndex != null && otherIndex.TryGetValue(path, out var otherKey) && otherKey.Values != null)
            {
                otherValues = new Dictionary<string, KeyValue>(StringComparer.OrdinalIgnoreCase);
                foreach (var v in otherKey.Values)
                {
                    var name = v.ValueName ?? "(Default)";
                    if (!otherValues.ContainsKey(name))
                        otherValues[name] = v;
                }
            }

            // Show values — filter to differences only when that mode is active
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
                // else: identical value — default TextPrimary color
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
            bool isLeft = (tree == _leftTreeView);
            var nodeIndex = isLeft ? _leftNodesByPath : _rightNodesByPath;
            
            // Fast path: check dictionary cache
            if (nodeIndex != null && nodeIndex.TryGetValue(path, out var cached))
                return cached;
            
            // Check if the path even exists in this tree's key index
            var keyIndex = isLeft ? _leftKeysByPath : _rightKeysByPath;
            if (keyIndex == null || !keyIndex.ContainsKey(path))
                return null;
            
            // Also check if it should be visible (in differences-only mode)
            if (_showDifferencesOnly)
            {
                var diffIndex = isLeft ? _leftDiffByPath : _rightDiffByPath;
                if (diffIndex == null || !diffIndex.TryGetValue(path, out var diff) || !diff.HasDifference)
                    return null;
            }
            
            // Slow path: walk from root, expanding as needed to create lazy nodes
            if (tree.Nodes.Count == 0) return null;
            var current = tree.Nodes[0];
            var rootTag = current.Tag as NodeTag;
            if (rootTag == null) return null;
            
            if (string.Equals(path, rootTag.Path, StringComparison.OrdinalIgnoreCase))
                return current;
            
            if (!path.StartsWith(rootTag.Path + "\\", StringComparison.OrdinalIgnoreCase))
                return null;
            
            var remaining = path.Substring(rootTag.Path.Length + 1);
            var segments = remaining.Split('\\');
            
            foreach (var segment in segments)
            {
                // Ensure node is populated (expand if it has dummy child)
                if (current.Nodes.Count == 1 && current.Nodes[0].Tag == null)
                    current.Expand(); // triggers BeforeExpand, populates children
                
                // Find child matching this segment
                TreeNode? found = null;
                foreach (TreeNode child in current.Nodes)
                {
                    if (child.Text.Equals(segment, StringComparison.OrdinalIgnoreCase))
                    {
                        found = child;
                        break;
                    }
                }
                
                if (found == null) return null;
                current = found;
            }
            
            return current;
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

            _diffOnlyCheckbox.ForeColor = ModernTheme.TextPrimary;

            // Status bar
            _statusPanel.BackColor = ModernTheme.Surface;
            _statusForeColor = ModernTheme.TextSecondary;
            _statusPanel.Invalidate();
            _cancelLoadButton.ForeColor = ModernTheme.TextSecondary;
            _cancelLoadButton.FlatAppearance.BorderColor = ModernTheme.Border;
            _cancelLoadButton.FlatAppearance.MouseOverBackColor = ModernTheme.Selection;
            _cancelLoadButton.FlatAppearance.MouseDownBackColor = ModernTheme.SurfaceLight;

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

        protected override void Dispose(bool disposing)
        {
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

        /// <summary>
        /// Pre-computed diff status for a registry key path.
        /// Computed once during comparison, used by lazy tree node creation.
        /// </summary>
        private record struct DiffInfo(
            bool HasDifference,          // This key or any descendant has a difference
            bool IsUniqueToThisHive,     // GREEN: unique AND all descendants unique
            bool HasValueDifference      // RED: value diff, mixed descendants, or child diffs
        );
    }
}

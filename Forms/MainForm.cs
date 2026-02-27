using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using RegistryParser.Abstractions;

namespace RegistryExpert
{
    // Helper class to store theming controls for the Analyze form
    internal class AnalyzeFormThemeData : IDisposable
    {
        public SplitContainer? MainSplit { get; set; }
        public SplitContainer? NetworkSplit { get; set; }
        public SplitContainer? ContentDetailSplit { get; set; }  // New: for resizable detail pane
        public Panel? LeftPanel { get; set; }
        public Panel? CategoryHeader { get; set; }
        public Label? CategoryTitle { get; set; }
        public ListBox? CategoryList { get; set; }
        public Panel? RightPanel { get; set; }
        public Panel? ContentHeader { get; set; }
        public Label? ContentTitle { get; set; }
        public FlowLayoutPanel? SubCategoryPanel { get; set; }
        public DataGridView? ContentGrid { get; set; }
        public Panel? DetailPanel { get; set; }  // Renamed from RegistryInfoPanel
        public TextBox? RegistryPathLabel { get; set; }
        public RichTextBox? RegistryValueBox { get; set; }  // Changed from TextBox to RichTextBox
        public ListBox? NetworkAdaptersList { get; set; }
        public Panel? NetworkAdaptersHeader { get; set; }
        public Label? NetworkAdaptersLabel { get; set; }
        public Panel? NetworkDetailsHeader { get; set; }
        public Label? NetworkDetailsLabel { get; set; }
        public DataGridView? NetworkDetailsGrid { get; set; }
        public List<Button> SubCategoryButtons { get; set; } = new();
        public List<Button> ServiceFilterButtons { get; set; } = new();
        // Appx filter panel controls
        public FlowLayoutPanel? AppxFilterPanel { get; set; }
        public List<Button> AppxFilterButtons { get; set; } = new();
        // Storage filter buttons
        public List<Button> StorageFilterButtons { get; set; } = new();
        // Firewall panel controls
        public Panel? FirewallPanel { get; set; }
        public FlowLayoutPanel? FirewallProfileButtonsPanel { get; set; }
        public Label? FirewallProfileLabel { get; set; }
        public Panel? FirewallRulesPanel { get; set; }
        public Panel? FirewallRulesHeader { get; set; }
        public Label? FirewallRulesLabel { get; set; }
        public DataGridView? FirewallRulesGrid { get; set; }
        public List<Button> FirewallProfileButtons { get; set; } = new();
        public Action? RefreshFirewallDisplay { get; set; }
        // Fonts used in drawing (need explicit disposal)
        public Font? CategoryTextFont { get; set; }
        public ToolTip? CategoryToolTip { get; set; }
        public ToolTip? SubCategoryToolTip { get; set; }
        // Category icon images (need explicit disposal)
        public Dictionary<string, Image> CategoryIcons { get; set; } = new();

        public void Dispose()
        {
            CategoryTextFont?.Dispose();
            CategoryToolTip?.Dispose();
            SubCategoryToolTip?.Dispose();
            foreach (var img in CategoryIcons.Values)
                img.Dispose();
            CategoryIcons.Clear();
        }
    }

    public partial class MainForm : Form
    {
        private OfflineRegistryParser? _parser;
        private RegistryInfoExtractor? _infoExtractor;
        private Form? _analyzeForm; // Track the analyze window for theme changes
        private SearchForm? _searchForm; // Track the search window for theme changes
        private Form? _statisticsForm; // Track the statistics window for theme changes
        private CompareForm? _compareForm; // Track the compare window for disposal
        private Form? _timelineForm; // Track the timeline window for theme changes
        private ImageList? _imageList; // Track for disposal
        private ImageList? _valueImageList; // Track for disposal (value type icons)
        private ToolTip? _sharedToolTip; // Single shared ToolTip for the form
        private Icon? _customIcon; // Track custom icon for disposal
        private string? _currentHivePath; // Track the current loaded hive file path
        private TreeNode? _previousSelectedNode; // Track for folder icon switching
        private Dictionary<string, Image> _toolbarIcons = new(); // Track toolbar icons for disposal

        // Cached icon fonts for Paint handlers (avoid allocating on every paint)
        private static readonly Font _iconFont12 = new Font("Segoe MDL2 Assets", 12F);
        private static readonly Font _iconFont13 = new Font("Segoe MDL2 Assets", 13F);
        private static readonly Font _iconFont16 = new Font("Segoe MDL2 Assets", 16F);
        private static readonly Font _iconFont10 = new Font("Segoe MDL2 Assets", 10F);
        
        // UI Controls
        private MenuStrip _menuStrip = null!;
        private Panel _toolbarPanel = null!;
        private SplitContainer _mainSplitContainer = null!;
        private SplitContainer _rightSplitContainer = null!;
        private TreeView _treeView = null!;
        private ListView _listView = null!;
        private RichTextBox _detailsBox = null!;
        private Panel _statusPanel = null!;
        private string _statusText = "Ready";
        private Color _statusForeColor;
        private Panel _statusRightPanel = null!;
        private Label _hiveTypeLabel = null!;
        private ProgressBar _loadProgressBar = null!;
        private Panel _progressWrapper = null!;
        private Button _cancelLoadButton = null!;
        private Panel _cancelWrapper = null!;
        private CancellationTokenSource? _loadCts;
        private Panel _dropPanel = null!;
        private Panel _bookmarkBar = null!;
        private Panel _bookmarkPanel = null!;
        private int _bookmarkExpandedWidth;

        // Bookmark definitions per hive type
        private static readonly Dictionary<string, List<(string Name, string Path)>> _bookmarksByHive = new()
        {
        ["SOFTWARE"] = new()
        {
            ("Activation", @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"),
            ("Component Based Servicing", @"SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing"),
            ("Current Version", @"SOFTWARE\Microsoft\Windows NT\CurrentVersion"),
            ("Installed Program", @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
            ("Logon UI", @"SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI"),
            ("Profile List", @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"),
            ("Startup Programs", @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run"),
            ("Windows Azure", @"SOFTWARE\Microsoft\Windows Azure"),
            ("Windows Defender", @"SOFTWARE\Microsoft\Windows Defender"),
            ("Windows Update", @"SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"),
        },
        ["SYSTEM"] = new()
        {
            ("Class: Adapter", @"SYSTEM\ControlSet001\Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}"),
            ("Class: Disk", @"SYSTEM\ControlSet001\Control\Class\{4d36e967-e325-11ce-bfc1-08002be10318}"),
            ("Crash Control", @"SYSTEM\ControlSet001\Control\CrashControl"),
            ("Firewall", @"SYSTEM\ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy"),
            ("Guest Agent", @"SYSTEM\ControlSet001\Services\WindowsAzureGuestAgent"),
            ("Network Adapters", @"SYSTEM\ControlSet001\Services\Tcpip\Parameters\Interfaces"),
            ("Network Shares", @"SYSTEM\ControlSet001\Services\LanmanServer\Shares"),
            ("NTLM", @"SYSTEM\ControlSet001\Control\Lsa"),
            ("RDP-Tcp", @"SYSTEM\ControlSet001\Control\Terminal Server\WinStations\RDP-Tcp"),
            ("Services", @"SYSTEM\ControlSet001\Services"),
            ("TLS/SSL", @"SYSTEM\ControlSet001\Control\SecurityProviders\SCHANNEL\Protocols"),
        },
    };

        public MainForm()
        {
            InitializeComponent();
            _parser = new OfflineRegistryParser();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Dispose MenuStrip FIRST - critical to unsubscribe from SystemEvents.UserPreferenceChanged
                // which holds a strong reference and prevents the application from closing
                _menuStrip?.Dispose();
                
                _parser?.Dispose();
                _imageList?.Dispose();
                _valueImageList?.Dispose();
                _analyzeForm?.Dispose();
                _statisticsForm?.Dispose();
                _searchForm?.Dispose();
                _compareForm?.Dispose();
                _timelineForm?.Dispose();
                _sharedToolTip?.Dispose();
                _customIcon?.Dispose();
                _loadCts?.Dispose();
                foreach (var img in _toolbarIcons.Values)
                    img.Dispose();
                _toolbarIcons.Clear();
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
            
            // Update controls that need manual DPI adjustment
            _toolbarPanel.Height = DpiHelper.Scale(52);
            _statusPanel.Height = DpiHelper.Scale(32);
            _statusPanel.Padding = DpiHelper.ScalePadding(10, 0, 10, 0);
            _statusRightPanel.Width = DpiHelper.Scale(170);
            _progressWrapper.Width = DpiHelper.Scale(100);
            _progressWrapper.Padding = new Padding(0, DpiHelper.Scale(8), 0, DpiHelper.Scale(8));
            _cancelWrapper.Width = DpiHelper.Scale(68);
            _cancelWrapper.Padding = new Padding(DpiHelper.Scale(4), DpiHelper.Scale(4), DpiHelper.Scale(4), DpiHelper.Scale(4));
            _hiveTypeLabel.Padding = DpiHelper.ScalePadding(4, 0, 0, 0);
            
            // Refresh the toolbar to recalculate button layouts
            CreateModernToolbar();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            
            // Enable DPI scaling
            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.AutoScaleDimensions = new SizeF(96F, 96F);
            
            // Form settings
            this.Text = "Registry Expert";
            this.Size = new Size(1280, 800);
            this.MinimumSize = new Size(900, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            
            // Set application icon
            try
            {
                var iconPath = Path.Combine(AppContext.BaseDirectory, "registry_fixed.ico");
                if (File.Exists(iconPath))
                {
                    _customIcon = new Icon(iconPath);
                    this.Icon = _customIcon;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Icon loading failed: {ex.Message}");
            }
            
            ModernTheme.ApplyTo(this);
            this.AllowDrop = true;
            this.DragEnter += MainForm_DragEnter;
            this.DragDrop += MainForm_DragDrop;

            // Menu Strip
            _menuStrip = new MenuStrip();
            _menuStrip.AccessibleName = "Main Menu";
            ModernTheme.ApplyTo(_menuStrip);
            _menuStrip.Padding = new Padding(8, 4, 0, 4);
            CreateMenu();
            
            // Modern Toolbar Panel
            _toolbarPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = DpiHelper.Scale(52),
                BackColor = ModernTheme.Surface,
                Padding = new Padding(8, 8, 8, 8)
            };
            CreateModernToolbar();

            // Status Panel (bottom)
            _statusPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = DpiHelper.Scale(32),
                BackColor = ModernTheme.Surface,
                Padding = DpiHelper.ScalePadding(10, 0, 10, 0)
            };
            
            // Owner-draw status text: enable double-buffering and paint handler
            // This avoids the Label control's intermittent text clipping during rapid progress updates
            _statusForeColor = ModernTheme.TextSecondary;
            typeof(Panel).InvokeMember("DoubleBuffered",
                BindingFlags.SetProperty | BindingFlags.Instance | BindingFlags.NonPublic,
                null, _statusPanel, new object[] { true });
            _statusPanel.Paint += StatusPanel_Paint;
            _statusPanel.AccessibleRole = AccessibleRole.StatusBar;
            _statusPanel.AccessibleName = _statusText;

            // Fixed-width right panel — always visible, never changes width.
            // Inner controls toggle visibility without affecting the status label bounds.
            _statusRightPanel = new Panel
            {
                Dock = DockStyle.Right,
                Width = DpiHelper.Scale(170),
                BackColor = Color.Transparent,
            };

            _hiveTypeLabel = new Label
            {
                Text = "",
                ForeColor = ModernTheme.Accent,
                Font = ModernTheme.BoldFont,
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleRight,
                Padding = DpiHelper.ScalePadding(4, 0, 0, 0)
            };

            // Progress bar inside a wrapper panel for vertical centering (thin bar)
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

            // Cancel button inside a wrapper panel for vertical centering
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

            // Dock.Fill added first (lowest z-order, fills remaining space)
            _statusRightPanel.Controls.Add(_hiveTypeLabel);
            // Dock.Right added after (higher z-order, docked first)
            _statusRightPanel.Controls.Add(_cancelWrapper);
            _statusRightPanel.Controls.Add(_progressWrapper);

            // Status text is owner-drawn via _statusPanel.Paint — no Label control needed
            _statusPanel.Controls.Add(_statusRightPanel);   // Dock.Right - fixed width, always visible

            // Main content area
            var contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ModernTheme.Background,
                Padding = new Padding(1)
            };

            // Drop Panel (shown when no hive is loaded)
            _dropPanel = CreateDropPanel();
            _dropPanel.AccessibleName = "Drop a registry hive file here or click Open";
            _dropPanel.AccessibleRole = AccessibleRole.Pane;
            
            // Main Split Container (Tree | Right Panel)
            _mainSplitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                BackColor = ModernTheme.Border,
                Panel1MinSize = 50,
                Panel2MinSize = 50,
                SplitterWidth = 3,
                Visible = false
            };
            _mainSplitContainer.Panel1.BackColor = ModernTheme.Background;
            _mainSplitContainer.Panel2.BackColor = ModernTheme.Background;

            // Left panel with header and tree
            var leftPanel = new Panel { Dock = DockStyle.Fill, BackColor = ModernTheme.Background };
            var treeHeader = CreateSectionHeader("Registry Keys", "\uE8B7");
            
            _imageList = CreateImageList();
            _treeView = new TreeView
            {
                Dock = DockStyle.Fill,
                ImageList = _imageList,
                HideSelection = false  // Keep selection visible when focus is lost
            };
            ModernTheme.ApplyTo(_treeView);
            _treeView.AccessibleName = "Registry Keys";
            _treeView.AfterSelect += TreeView_AfterSelect;
            _treeView.BeforeExpand += TreeView_BeforeExpand;
            
            // Context menu for tree view - Copy Path
            var treeContextMenu = new ContextMenuStrip();
            treeContextMenu.Renderer = new ToolStripProfessionalRenderer(new ModernColorTable());
            var copyPathItem = new ToolStripMenuItem("Copy Path", null, (s, ev) =>
            {
                if (_treeView.SelectedNode?.Tag is RegistryKey key)
                {
                    Clipboard.SetText(key.KeyPath);
                }
            });
            copyPathItem.ShortcutKeys = Keys.Control | Keys.C;
            treeContextMenu.Items.Add(copyPathItem);
            treeContextMenu.Opening += (s, ev) =>
            {
                treeContextMenu.BackColor = ModernTheme.Surface;
                treeContextMenu.ForeColor = ModernTheme.TextPrimary;
                foreach (ToolStripItem item in treeContextMenu.Items)
                {
                    item.ForeColor = ModernTheme.TextPrimary;
                    item.BackColor = ModernTheme.Surface;
                }
                // Disable if no valid node selected
                copyPathItem.Enabled = _treeView.SelectedNode?.Tag is RegistryKey;
            };
            _treeView.NodeMouseClick += (s, ev) =>
            {
                if (ev.Button == MouseButtons.Right)
                {
                    _treeView.SelectedNode = ev.Node;
                }
            };
            _treeView.ContextMenuStrip = treeContextMenu;
            
            leftPanel.Controls.Add(_treeView);
            leftPanel.Controls.Add(CreateBookmarkPanels());
            leftPanel.Controls.Add(treeHeader);
            _mainSplitContainer.Panel1.Controls.Add(leftPanel);

            // Right Split Container (ListView | Details)
            _rightSplitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                BackColor = ModernTheme.Border,
                Panel1MinSize = 50,
                Panel2MinSize = 50,
                SplitterWidth = 3
            };
            _rightSplitContainer.Panel1.BackColor = ModernTheme.Background;
            _rightSplitContainer.Panel2.BackColor = ModernTheme.Background;

            // Values panel with header and list
            var valuesPanel = new Panel { Dock = DockStyle.Fill, BackColor = ModernTheme.Background };
            var valuesHeader = CreateSectionHeader("Values", "\uE8A5");
            
            _listView = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                HeaderStyle = ColumnHeaderStyle.Nonclickable
            };
            ModernTheme.ApplyTo(_listView);
            _listView.AccessibleName = "Registry Values";
            _valueImageList = CreateValueImageList();
            _listView.SmallImageList = _valueImageList;
            _listView.Columns.Add("Name", 220);
            _listView.Columns.Add("Type", 100);
            _listView.Columns.Add("Data", 450);
            _listView.SelectedIndexChanged += ListView_SelectedIndexChanged;
            _listView.DoubleClick += ListView_DoubleClick;
            // Dynamically expand the Data column when the window/split resizes
            _listView.Resize += (s, e) => AdjustValuesColumns();
            _rightSplitContainer.SplitterMoved += (s, e) => AdjustValuesColumns();
            _rightSplitContainer.SizeChanged += (s, e) => AdjustValuesColumns();
            
            // Context menu for list view - Copy Value
            var listContextMenu = new ContextMenuStrip();
            listContextMenu.Renderer = new ToolStripProfessionalRenderer(new ModernColorTable());
            var copyValueItem = new ToolStripMenuItem("Copy Value", null, (s, ev) =>
            {
                if (_listView.SelectedItems.Count > 0)
                {
                    var data = _listView.SelectedItems[0].SubItems[2].Text;
                    if (!string.IsNullOrEmpty(data))
                    {
                        Clipboard.SetText(data);
                    }
                }
            });
            copyValueItem.ShortcutKeys = Keys.Control | Keys.C;
            listContextMenu.Items.Add(copyValueItem);
            listContextMenu.Opening += (s, ev) =>
            {
                listContextMenu.BackColor = ModernTheme.Surface;
                listContextMenu.ForeColor = ModernTheme.TextPrimary;
                foreach (ToolStripItem item in listContextMenu.Items)
                {
                    item.ForeColor = ModernTheme.TextPrimary;
                    item.BackColor = ModernTheme.Surface;
                }
                // Disable if no value selected
                copyValueItem.Enabled = _listView.SelectedItems.Count > 0;
            };
            _listView.MouseDown += (s, ev) =>
            {
                if (ev.Button == MouseButtons.Right)
                {
                    var hitTest = _listView.HitTest(ev.X, ev.Y);
                    if (hitTest.Item != null)
                    {
                        hitTest.Item.Selected = true;
                    }
                }
            };
            _listView.ContextMenuStrip = listContextMenu;
            
            valuesPanel.Controls.Add(_listView);
            valuesPanel.Controls.Add(valuesHeader);
            _rightSplitContainer.Panel1.Controls.Add(valuesPanel);
            AdjustValuesColumns();

            // Details panel with header and text
            var detailsPanel = new Panel { Dock = DockStyle.Fill, BackColor = ModernTheme.Background };
            var detailsHeader = CreateSectionHeader("Details", "\uE946");
            
            _detailsBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                WordWrap = true,
                ScrollBars = RichTextBoxScrollBars.Vertical
            };
            ModernTheme.ApplyTo(_detailsBox);
            _detailsBox.AccessibleName = "Value Details";
            _detailsBox.Padding = new Padding(8);
            
            detailsPanel.Controls.Add(_detailsBox);
            detailsPanel.Controls.Add(detailsHeader);
            _rightSplitContainer.Panel2.Controls.Add(detailsPanel);

            _mainSplitContainer.Panel2.Controls.Add(_rightSplitContainer);
            
            contentPanel.Controls.Add(_mainSplitContainer);
            contentPanel.Controls.Add(_dropPanel);

            // Add controls to form in correct order
            this.Controls.Add(contentPanel);
            this.Controls.Add(_statusPanel);
            this.Controls.Add(_toolbarPanel);
            this.Controls.Add(_menuStrip);
            this.MainMenuStrip = _menuStrip;
            
            this.ResumeLayout(false);
            this.PerformLayout();
            
            // Set splitter distances after layout
            this.Load += async (s, e) => {
                _mainSplitContainer.SplitterDistance = 280;
                _rightSplitContainer.SplitterDistance = Math.Max(250, _rightSplitContainer.Height * 2 / 3);
                try { await CheckForUpdatesOnStartupAsync(); }
                catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"Startup update check failed: {ex.Message}"); }
            };
        }

        private Panel CreateDropPanel()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ModernTheme.Background
            };
            
            // Card-like center panel with subtle border
            var centerPanel = new Panel
            {
                Size = DpiHelper.ScaleSize(450, 400),
                BackColor = ModernTheme.Surface
            };
            centerPanel.Paint += (s, e) =>
            {
                // Draw border manually with four lines for proper corner connection
                var w = centerPanel.Width - 1;
                var h = centerPanel.Height - 1;
                using var pen = new Pen(ModernTheme.Border, 1);
                // Top line
                e.Graphics.DrawLine(pen, 0, 0, w, 0);
                // Right line
                e.Graphics.DrawLine(pen, w, 0, w, h);
                // Bottom line
                e.Graphics.DrawLine(pen, w, h, 0, h);
                // Left line
                e.Graphics.DrawLine(pen, 0, h, 0, 0);
            };
            
            // Program icon - load from PNG for best quality
            var iconBox = new PictureBox
            {
                Size = DpiHelper.ScaleSize(72, 72),
                Location = DpiHelper.ScalePoint((450 - 72) / 2, 30),
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.Transparent
            };
            try
            {
                var pngPath = Path.Combine(AppContext.BaseDirectory, "registry.png");
                if (File.Exists(pngPath))
                {
                    iconBox.Image = Image.FromFile(pngPath);
                }
            }
            catch { }
            
            var titleLabel = new Label
            {
                Text = "Registry Expert",
                Font = ModernTheme.TitleFont,
                ForeColor = ModernTheme.TextPrimary,
                TextAlign = ContentAlignment.MiddleCenter,
                Size = DpiHelper.ScaleSize(450, 45),
                Location = DpiHelper.ScalePoint(0, 115)
            };
            
            var subtitleLabel = new Label
            {
                Text = "Drag and drop a registry hive file here\nor click the button below to browse",
                Font = ModernTheme.RegularFont,
                ForeColor = ModernTheme.TextSecondary,
                TextAlign = ContentAlignment.MiddleCenter,
                Size = DpiHelper.ScaleSize(450, 50),
                Location = DpiHelper.ScalePoint(0, 160)
            };
            
            var openButton = ModernTheme.CreateButton("  Open Hive File  ", OpenHive_Click);
            openButton.Size = DpiHelper.ScaleSize(160, 40);
            openButton.Location = DpiHelper.ScalePoint(145, 225);
            
            // Compare button
            var compareButton = ModernTheme.CreateButton("  Compare Hives  ", (s, e) => ShowCompare_Click(s, e));
            compareButton.Size = DpiHelper.ScaleSize(160, 40);
            compareButton.Location = DpiHelper.ScalePoint(145, 280);
            compareButton.BackColor = ModernTheme.Surface;
            compareButton.ForeColor = ModernTheme.TextPrimary;
            compareButton.FlatAppearance.BorderColor = ModernTheme.Accent;
            compareButton.FlatAppearance.BorderSize = 1;
            
            // Hint text
            var hintLabel = new Label
            {
                Text = "Supports SYSTEM, SOFTWARE, SAM, SECURITY, NTUSER.DAT, and more",
                Font = ModernTheme.SmallFont,
                ForeColor = ModernTheme.TextDisabled,
                TextAlign = ContentAlignment.MiddleCenter,
                Size = DpiHelper.ScaleSize(450, 25),
                Location = DpiHelper.ScalePoint(0, 340)
            };
            
            centerPanel.Controls.AddRange(new Control[] { iconBox, titleLabel, subtitleLabel, openButton, compareButton, hintLabel });
            panel.Controls.Add(centerPanel);
            
            // Center the panel
            panel.Resize += (s, e) =>
            {
                centerPanel.Location = new Point(
                    (panel.Width - centerPanel.Width) / 2,
                    (panel.Height - centerPanel.Height) / 2
                );
            };
            
            return panel;
        }

        private void RefreshDropPanel()
        {
            if (_dropPanel == null) return;
            
            var parent = _dropPanel.Parent;
            var wasVisible = _dropPanel.Visible;
            var index = parent?.Controls.IndexOf(_dropPanel) ?? -1;
            
            // Remove old drop panel and dispose image resources
            parent?.Controls.Remove(_dropPanel);
            foreach (Control ctrl in _dropPanel.Controls)
            {
                if (ctrl is Panel innerPanel)
                {
                    foreach (Control innerCtrl in innerPanel.Controls)
                    {
                        if (innerCtrl is PictureBox pb && pb.Image != null)
                        {
                            pb.Image.Dispose();
                            pb.Image = null;
                        }
                    }
                }
            }
            _dropPanel.Dispose();
            
            // Create new drop panel with updated theme colors
            _dropPanel = CreateDropPanel();
            _dropPanel.Visible = wasVisible;
            
            // Add back to parent at same position
            if (parent != null && index >= 0)
            {
                parent.Controls.Add(_dropPanel);
                parent.Controls.SetChildIndex(_dropPanel, index);
            }
        }

        private Panel CreateSectionHeader(string text, string icon)
        {
            var panel = new Panel
            {
                Height = DpiHelper.Scale(40),
                Dock = DockStyle.Top,
                BackColor = ModernTheme.Surface,
                Padding = new Padding(16, 0, 16, 0)
            };
            
            // Custom paint for gradient and icon
            panel.Paint += (s, e) =>
            {
                // Subtle gradient
                using var brush = new LinearGradientBrush(panel.ClientRectangle, 
                    ModernTheme.GradientStart, ModernTheme.GradientEnd, LinearGradientMode.Vertical);
                e.Graphics.FillRectangle(brush, panel.ClientRectangle);
                
                // Icon
                var iconFont = _iconFont12;
                using var iconBrush = new SolidBrush(ModernTheme.Accent);
                e.Graphics.DrawString(icon, iconFont, iconBrush, DpiHelper.Scale(16), (panel.Height - iconFont.Height) / 2);
                
                // Text
                using var textBrush = new SolidBrush(ModernTheme.TextPrimary);
                e.Graphics.DrawString(text, ModernTheme.HeaderFont, textBrush, DpiHelper.Scale(40), (panel.Height - ModernTheme.HeaderFont.Height) / 2);
            };
            
            // Bottom border
            var border = new Panel
            {
                Height = 1,
                Dock = DockStyle.Bottom,
                BackColor = ModernTheme.Border
            };
            
            panel.Controls.Add(border);
            return panel;
        }

        /// <summary>
        /// Creates the bookmark bar (collapsed) and bookmark panel (expanded).
        /// Returns a container panel (DockStyle.Left) holding both.
        /// </summary>
        private Panel CreateBookmarkPanels()
        {
            int barWidth = DpiHelper.Scale(24);

            // Container holds both bar and panel, only one visible at a time
            var container = new Panel
            {
                Dock = DockStyle.Left,
                Width = barWidth,
                BackColor = ModernTheme.Background,
                Visible = false // Hidden until a hive with bookmarks is loaded
            };

            // --- Collapsed bar: narrow strip with rotated "Bookmarks" text + ▶ arrow ---
            _bookmarkBar = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ModernTheme.Surface,
                Cursor = Cursors.Hand,
                Visible = true
            };
            _bookmarkBar.Paint += (s, e) =>
            {
                var g = e.Graphics;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                // Draw ▶ arrow at top
                using var arrowFont = new Font("Segoe UI", 8F);
                using var arrowBrush = new SolidBrush(ModernTheme.Accent);
                g.DrawString("\u25B6", arrowFont, arrowBrush, DpiHelper.Scale(4), DpiHelper.Scale(6));

                // Draw "Bookmarks" rotated 90° (top-to-bottom)
                using var textFont = new Font("Segoe UI", 9F);
                using var textBrush = new SolidBrush(ModernTheme.TextSecondary);
                var textSize = g.MeasureString("Bookmarks", textFont);
                var state = g.Save();
                g.TranslateTransform(_bookmarkBar.Width / 2 + textSize.Height / 2 - DpiHelper.Scale(2), DpiHelper.Scale(24));
                g.RotateTransform(90);
                g.DrawString("Bookmarks", textFont, textBrush, 0, 0);
                g.Restore(state);

                // Draw bookmark icon below the rotated text (GDI+ drawn, DPI-aware)
                int iconW = DpiHelper.Scale(14);
                int iconH = DpiHelper.Scale(18);
                int iconX = (_bookmarkBar.Width - iconW) / 2;
                int iconY = DpiHelper.Scale(24) + (int)textSize.Width + DpiHelper.Scale(6);
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                using var iconBrush = new SolidBrush(ModernTheme.Accent);
                // Ribbon/bookmark shape: rectangle with a V-notch at the bottom
                var ribbon = new PointF[]
                {
                    new PointF(iconX, iconY),
                    new PointF(iconX + iconW, iconY),
                    new PointF(iconX + iconW, iconY + iconH),
                    new PointF(iconX + iconW / 2f, iconY + iconH - DpiHelper.Scale(5)),
                    new PointF(iconX, iconY + iconH),
                };
                g.FillPolygon(iconBrush, ribbon);
            };
            _bookmarkBar.Click += (s, e) => ToggleBookmarkPanel(true);

            // --- Expanded panel: collapse bar on right + bookmark items ---
            _bookmarkPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ModernTheme.Surface,
                Visible = false,
                Padding = new Padding(0)
            };

            // Right-side collapse bar (mirrors the collapsed bar style, with ◀ arrow)
            var collapseBar = new Panel
            {
                Dock = DockStyle.Right,
                Width = barWidth,
                BackColor = ModernTheme.Surface,
                Cursor = Cursors.Hand,
                Tag = "bookmarkCollapseBar"
            };
            collapseBar.Paint += (s, e) =>
            {
                var g = e.Graphics;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                // Draw ◀ arrow at top
                using var arrowFont = new Font("Segoe UI", 8F);
                using var arrowBrush = new SolidBrush(ModernTheme.Accent);
                g.DrawString("\u25C0", arrowFont, arrowBrush, DpiHelper.Scale(4), DpiHelper.Scale(6));

                // Draw "Bookmarks" rotated 90° (top-to-bottom)
                using var textFont = new Font("Segoe UI", 9F);
                using var textBrush = new SolidBrush(ModernTheme.TextSecondary);
                var textSize = g.MeasureString("Bookmarks", textFont);
                var state = g.Save();
                g.TranslateTransform(collapseBar.Width / 2 + textSize.Height / 2 - DpiHelper.Scale(2), DpiHelper.Scale(24));
                g.RotateTransform(90);
                g.DrawString("Bookmarks", textFont, textBrush, 0, 0);
                g.Restore(state);

                // Draw bookmark ribbon icon below the rotated text
                int iconW = DpiHelper.Scale(14);
                int iconH = DpiHelper.Scale(18);
                int iconX = (collapseBar.Width - iconW) / 2;
                int iconY = DpiHelper.Scale(24) + (int)textSize.Width + DpiHelper.Scale(6);
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                using var iconBrush = new SolidBrush(ModernTheme.Accent);
                var ribbon = new PointF[]
                {
                    new PointF(iconX, iconY),
                    new PointF(iconX + iconW, iconY),
                    new PointF(iconX + iconW, iconY + iconH),
                    new PointF(iconX + iconW / 2f, iconY + iconH - DpiHelper.Scale(5)),
                    new PointF(iconX, iconY + iconH),
                };
                g.FillPolygon(iconBrush, ribbon);
            };
            collapseBar.Click += (s, e) => ToggleBookmarkPanel(false);

            // Items container
            var itemsPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true,
                BackColor = ModernTheme.Surface,
                Padding = DpiHelper.ScalePadding(4, 6, 4, 4),
                Tag = "bookmarkItems"
            };

            // WinForms dock order: last added docks first
            _bookmarkPanel.Controls.Add(itemsPanel);
            _bookmarkPanel.Controls.Add(collapseBar);

            // Container holds both — bar is Fill (visible), panel is Fill (hidden)
            container.Controls.Add(_bookmarkBar);
            container.Controls.Add(_bookmarkPanel);
            container.Tag = "bookmarkContainer";

            return container;
        }

        /// <summary>
        /// Toggle the bookmark panel between collapsed (bar) and expanded (panel).
        /// </summary>
        private void ToggleBookmarkPanel(bool expand)
        {
            var container = _bookmarkBar.Parent;
            if (container == null) return;

            if (expand)
            {
                // Use adaptive width calculated from bookmark text, fallback to 200px if not yet computed
                container.Width = _bookmarkExpandedWidth > 0 ? _bookmarkExpandedWidth : DpiHelper.Scale(200);
                _bookmarkBar.Visible = false;
                _bookmarkPanel.Visible = true;
            }
            else
            {
                container.Width = DpiHelper.Scale(24);
                _bookmarkBar.Visible = true;
                _bookmarkPanel.Visible = false;
            }
        }

        /// <summary>
        /// Populate bookmark items for the current hive type. Hides the panel if no bookmarks exist.
        /// </summary>
        private void PopulateBookmarks(string hiveType)
        {
            var container = _bookmarkBar.Parent;
            if (container == null) return;

            if (!_bookmarksByHive.TryGetValue(hiveType, out var bookmarks) || bookmarks.Count == 0)
            {
                container.Visible = false;
                return;
            }

            // Find the items panel
            FlowLayoutPanel? itemsPanel = null;
            foreach (Control c in _bookmarkPanel.Controls)
            {
                if (c is FlowLayoutPanel fp && fp.Tag as string == "bookmarkItems")
                {
                    itemsPanel = fp;
                    break;
                }
            }
            if (itemsPanel == null) return;

            // Clear previous items
            itemsPanel.Controls.Clear();

            // Measure the widest bookmark text to determine adaptive panel width
            int maxTextWidth = 0;
            foreach (var (name, _) in bookmarks)
            {
                var text = $"  \u25B8  {name}";
                var textWidth = TextRenderer.MeasureText(text, ModernTheme.RegularFont).Width;
                if (textWidth > maxTextWidth)
                    maxTextWidth = textWidth;
            }

            // Item width = widest text + left/right padding + small extra margin
            int itemPadding = DpiHelper.Scale(2) * 2;
            int itemWidth = maxTextWidth + itemPadding + DpiHelper.Scale(4);

            // Expanded panel width = item width + collapse bar + flow panel padding
            int barWidth = DpiHelper.Scale(24);
            int flowPadding = DpiHelper.Scale(4) * 2;
            _bookmarkExpandedWidth = itemWidth + barWidth + flowPadding;

            // Add bookmark labels
            foreach (var (name, path) in bookmarks)
            {
                var btn = new Label
                {
                    Text = $"  \u25B8  {name}",
                    Font = ModernTheme.RegularFont,
                    ForeColor = ModernTheme.TextPrimary,
                    BackColor = ModernTheme.Surface,
                    AutoSize = false,
                    AutoEllipsis = true,
                    Width = itemWidth,
                    Height = DpiHelper.Scale(28),
                    TextAlign = ContentAlignment.MiddleLeft,
                    Cursor = Cursors.Hand,
                    Padding = DpiHelper.ScalePadding(2, 0, 2, 0),
                    Tag = path
                };
                btn.MouseEnter += (s, e) =>
                {
                    btn.BackColor = ModernTheme.Selection;
                    btn.ForeColor = ModernTheme.Accent;
                };
                btn.MouseLeave += (s, e) =>
                {
                    btn.BackColor = ModernTheme.Surface;
                    btn.ForeColor = ModernTheme.TextPrimary;
                };
                btn.Click += (s, e) =>
                {
                    if (btn.Tag is string keyPath)
                    {
                        NavigateToKey(keyPath);
                    }
                };
                itemsPanel.Controls.Add(btn);
            }

            // Show the container in collapsed state
            container.Visible = true;
            ToggleBookmarkPanel(false);
        }

        private ImageList CreateImageList()
        {
            var imageList = new ImageList
            {
                ColorDepth = ColorDepth.Depth32Bit,
                ImageSize = new Size(18, 18)
            };
            
            imageList.Images.Add("folder", ModernTheme.CreateFolderIcon(false));
            imageList.Images.Add("folder_open", ModernTheme.CreateFolderIcon(true));
            imageList.Images.Add("value", ModernTheme.CreateValueIcon());
            
            return imageList;
        }

        private ImageList CreateValueImageList() => ModernTheme.CreateValueImageList();

        private static string GetValueImageKey(string? valueType) => ModernTheme.GetValueImageKey(valueType);

        /// <summary>
        /// Strips leading emoji/symbol characters from text for use in AccessibleName properties.
        /// Emoji prefixed titles follow the pattern "&lt;emoji&gt; &lt;text&gt;" (e.g. "📁 Hive Information" → "Hive Information").
        /// </summary>
        private static string StripEmojiPrefix(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            int i = 0;
            while (i < text.Length)
            {
                char c = text[i];
                // Skip surrogate pairs (emoji like 📁, 🪟, etc.)
                if (char.IsHighSurrogate(c) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                {
                    i += 2;
                    continue;
                }
                // Skip common symbol/emoji characters in BMP (✅, ❌, ⛔, ⚡, ☁, etc.)
                if (c > '\u2000' && !char.IsLetterOrDigit(c) && c != '(')
                {
                    i++;
                    continue;
                }
                // Skip variation selectors (U+FE0E, U+FE0F) and zero-width joiners
                if (c == '\uFE0E' || c == '\uFE0F' || c == '\u200D')
                {
                    i++;
                    continue;
                }
                break;
            }
            // Skip trailing spaces after emoji
            while (i < text.Length && text[i] == ' ') i++;
            return i > 0 && i < text.Length ? text[i..] : text;
        }

        private void CreateMenu()
        {
            var fileMenu = new ToolStripMenuItem("&File");
            fileMenu.DropDownItems.Add(new ToolStripMenuItem("&Open Hive...", null, OpenHive_Click) { ShortcutKeys = Keys.Control | Keys.O });
            fileMenu.DropDownItems.Add(new ToolStripSeparator());
            fileMenu.DropDownItems.Add(new ToolStripMenuItem("&Export Key...", null, ExportKey_Click) { ShortcutKeys = Keys.Control | Keys.E });
            fileMenu.DropDownItems.Add(new ToolStripSeparator());
            fileMenu.DropDownItems.Add(new ToolStripMenuItem("E&xit", null, (s, e) => Close()) { ShortcutKeys = Keys.Alt | Keys.F4 });
            _menuStrip.Items.Add(fileMenu);

            var viewMenu = new ToolStripMenuItem("&View");
            
            var darkThemeItem = new ToolStripMenuItem("&Dark Theme", null, (s, e) => SwitchTheme(ThemeType.Dark));
            var lightThemeItem = new ToolStripMenuItem("&Light Theme", null, (s, e) => SwitchTheme(ThemeType.Light));
            darkThemeItem.Checked = ModernTheme.CurrentTheme == ThemeType.Dark;
            lightThemeItem.Checked = ModernTheme.CurrentTheme == ThemeType.Light;
            viewMenu.DropDownItems.Add(darkThemeItem);
            viewMenu.DropDownItems.Add(lightThemeItem);
            _menuStrip.Items.Add(viewMenu);

            var toolsMenu = new ToolStripMenuItem("&Tools");
            toolsMenu.DropDownItems.Add(new ToolStripMenuItem("&Search...", null, Search_Click) { ShortcutKeys = Keys.Control | Keys.F });
            toolsMenu.DropDownItems.Add(new ToolStripSeparator());
            toolsMenu.DropDownItems.Add(new ToolStripMenuItem("&Analyze...", null, ShowAnalyzeDialog_Click) { ShortcutKeys = Keys.F5 });
            toolsMenu.DropDownItems.Add(new ToolStripMenuItem("S&tatistics...", null, ShowStatistics_Click) { ShortcutKeys = Keys.Control | Keys.I });
            toolsMenu.DropDownItems.Add(new ToolStripMenuItem("Ti&meline...", null, ShowTimeline_Click) { ShortcutKeys = Keys.Control | Keys.T });
            toolsMenu.DropDownItems.Add(new ToolStripMenuItem("&Compare...", null, ShowCompare_Click) { ShortcutKeys = Keys.Control | Keys.M });
            _menuStrip.Items.Add(toolsMenu);

            var helpMenu = new ToolStripMenuItem("&Help");
            helpMenu.DropDownItems.Add("Check for &Updates...", null, CheckForUpdates_Click);
            helpMenu.DropDownItems.Add(new ToolStripSeparator());
            helpMenu.DropDownItems.Add("&About", null, About_Click);
            _menuStrip.Items.Add(helpMenu);
        }

        private void CreateModernToolbar()
        {
            // Dispose old toolbar controls before creating new ones (prevents leak on DPI change)
            foreach (Control ctrl in _toolbarPanel.Controls)
                ctrl.Dispose();
            _toolbarPanel.Controls.Clear();

            // Dispose old toolbar icons and reload at current DPI
            foreach (var img in _toolbarIcons.Values)
                img.Dispose();
            _toolbarIcons.Clear();

            var assembly = System.Reflection.Assembly.GetExecutingAssembly();
            int toolbarIconSize = DpiHelper.Scale(28);
            var toolbarIconNames = new[] { "open", "search", "analyze", "statistics", "compare", "timeline" };
            foreach (var name in toolbarIconNames)
            {
                var resourceName = $"RegistryExpert.icons.toolbar_{name}.png";
                using var stream = assembly.GetManifestResourceStream(resourceName);
                if (stream != null)
                {
                    using var original = Image.FromStream(stream);
                    var scaled = new Bitmap(toolbarIconSize, toolbarIconSize);
                    using (var g = Graphics.FromImage(scaled))
                    {
                        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                        g.DrawImage(original, 0, 0, toolbarIconSize, toolbarIconSize);
                    }
                    _toolbarIcons[name] = scaled;
                }
            }

            var flow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                BackColor = Color.Transparent,
                Padding = new Padding(4, 0, 0, 0)
            };

            flow.Controls.Add(CreateToolbarButton("Open", "open", OpenHive_Click, "Open registry hive (Ctrl+O)"));
            flow.Controls.Add(CreateToolbarSeparator());
            flow.Controls.Add(CreateToolbarButton("Search", "search", Search_Click, "Search registry (Ctrl+F)"));
            flow.Controls.Add(CreateToolbarSeparator());
            flow.Controls.Add(CreateToolbarButton("Analyze", "analyze", ShowAnalyzeDialog_Click, "Analyze registry (F5)"));
            flow.Controls.Add(CreateToolbarSeparator());
            flow.Controls.Add(CreateToolbarButton("Statistics", "statistics", ShowStatistics_Click, "View registry statistics"));
            flow.Controls.Add(CreateToolbarSeparator());
            flow.Controls.Add(CreateToolbarButton("Compare", "compare", ShowCompare_Click, "Compare two registry hives"));
            flow.Controls.Add(CreateToolbarSeparator());
            flow.Controls.Add(CreateToolbarButton("Timeline", "timeline", ShowTimeline_Click, "View registry keys by last modified time (Ctrl+T)"));
            
            _toolbarPanel.Controls.Add(flow);
        }

        private Button CreateToolbarButton(string text, string iconKey, EventHandler onClick, string tooltip)
        {
            var btn = new Button
            {
                Text = "",  // We'll draw text manually
                AccessibleName = text,
                AccessibleRole = AccessibleRole.PushButton,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.Transparent,
                ForeColor = ModernTheme.TextPrimary,
                Font = ModernTheme.RegularFont,
                Size = DpiHelper.ScaleSize(105, 36),
                Margin = DpiHelper.ScalePadding(2),
                Cursor = Cursors.Hand
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.FlatAppearance.MouseOverBackColor = ModernTheme.SurfaceHover;
            btn.FlatAppearance.MouseDownBackColor = ModernTheme.AccentDark;
            btn.Click += onClick;
            
            // Custom paint for icon image and text
            btn.Paint += (s, e) =>
            {
                var g = e.Graphics;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                
                var iconX = DpiHelper.Scale(10);
                
                // Draw icon image
                if (_toolbarIcons.TryGetValue(iconKey, out var iconImage))
                {
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    var iconY = (btn.Height - iconImage.Height) / 2;
                    g.DrawImage(iconImage, iconX, iconY, iconImage.Width, iconImage.Height);
                }
                
                // Draw text after icon with proper gap
                var textX = iconX + DpiHelper.Scale(28) + DpiHelper.Scale(4);  // icon width + 4px gap
                using var textBrush = new SolidBrush(ModernTheme.TextPrimary);
                var textY = (btn.Height - ModernTheme.RegularFont.Height) / 2;
                g.DrawString(text, ModernTheme.RegularFont, textBrush, textX, textY);
            };
            
            _sharedToolTip ??= new ToolTip();
            _sharedToolTip.SetToolTip(btn, tooltip);
            
            return btn;
        }

        private Panel CreateToolbarSeparator()
        {
            return new Panel
            {
                Width = DpiHelper.Scale(1),
                Height = DpiHelper.Scale(28),
                Margin = DpiHelper.ScalePadding(8, 4, 8, 4),
                BackColor = ModernTheme.Border
            };
        }

        private void MainForm_DragEnter(object? sender, DragEventArgs e)
        {
            if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
            {
                e.Effect = DragDropEffects.Copy;
                _dropPanel.BackColor = ModernTheme.SurfaceLight;
            }
        }

        private void MainForm_DragDrop(object? sender, DragEventArgs e)
        {
            _dropPanel.BackColor = ModernTheme.Background;
            var files = e.Data?.GetData(DataFormats.FileDrop) as string[];
            if (files?.Length > 0)
            {
                _ = LoadHiveFileAsync(files[0]);
            }
        }

        private void OpenHive_Click(object? sender, EventArgs e)
        {
            using var dialog = new OpenFileDialog
            {
                Title = "Open Registry Hive",
                Filter = "Registry Hives|NTUSER.DAT;SAM;SECURITY;SOFTWARE;SYSTEM;USRCLASS.DAT;DEFAULT;Amcache.hve;BCD|All Files|*.*",
                FilterIndex = 2
            };

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                _ = LoadHiveFileAsync(dialog.FileName);
            }
        }

        /// <summary>
        /// Sets the status bar text and optional color, then invalidates for repaint.
        /// Uses owner-draw to avoid WinForms Label clipping during rapid updates.
        /// </summary>
        private void SetStatusText(string text, Color? color = null)
        {
            _statusText = text;
            _statusPanel.AccessibleName = text;
            if (color.HasValue)
                _statusForeColor = color.Value;
            _statusPanel.Invalidate();
        }

        /// <summary>
        /// Owner-draw handler for the status panel — draws _statusText with EndEllipsis.
        /// Bypasses the Label control entirely, using actual panel bounds at paint time.
        /// </summary>
        private void StatusPanel_Paint(object? sender, PaintEventArgs e)
        {
            // Calculate text area: full panel client area minus the right panel's space and padding
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

        private void CancelLoad_Click(object? sender, EventArgs e)
        {
            _loadCts?.Cancel();
            _cancelLoadButton.Enabled = false;
            SetStatusText("Cancelling...");
        }

        private async Task LoadHiveFileAsync(string filePath)
        {
            try
            {
                _loadCts?.Cancel();
                _loadCts?.Dispose();
                _loadCts = new CancellationTokenSource();
                var token = _loadCts.Token;

                // Show progress UI — visibility changes inside _statusRightPanel
                // don't affect status text bounds (fixed-width right panel)
                _hiveTypeLabel.Visible = false;
                SetStatusText("Loading hive...", ModernTheme.Warning);
                _loadProgressBar.Value = 0;
                _progressWrapper.Visible = true;
                _cancelWrapper.Visible = true;
                _cancelLoadButton.Enabled = true;

                _parser?.Dispose();
                _parser = new OfflineRegistryParser();

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

                await Task.Run(() => _parser.LoadHive(filePath, progress, token), token).ConfigureAwait(true);

                _currentHivePath = filePath;
                _infoExtractor = new RegistryInfoExtractor(_parser);

                PopulateTreeView();
                PopulateBookmarks(_parser.CurrentHiveType.ToString());

                // Show main view, hide drop panel
                _dropPanel.Visible = false;
                _mainSplitContainer.Visible = true;
                _mainSplitContainer.SplitterDistance = _mainSplitContainer.Width * 3 / 7;

                _hiveTypeLabel.Text = $"● {_parser.CurrentHiveType}";
                SetStatusText($"Loaded: {Path.GetFileName(filePath)}", ModernTheme.Success);
                this.Text = $"Registry Expert - {Path.GetFileName(filePath)}";
            }
            catch (OperationCanceledException)
            {
                _parser?.Dispose();
                _parser = null;
                SetStatusText("Loading cancelled", ModernTheme.Warning);
            }
            catch (Exception ex)
            {
                ShowError($"Error loading hive: {ex.Message}");
                SetStatusText("Failed to load hive", ModernTheme.Error);
            }
            finally
            {
                _progressWrapper.Visible = false;
                _cancelWrapper.Visible = false;
                _hiveTypeLabel.Visible = true;
            }
        }

        private void PopulateTreeView()
        {
            _previousSelectedNode = null;  // Reset before clearing tree to prevent stale references
            _treeView.Nodes.Clear();
            _listView.Items.Clear();
            _detailsBox.Clear();

            var root = _parser?.GetRootKey();
            if (root == null || _parser == null) return;

            var rootNode = CreateTreeNode(root);
            rootNode.Text = $"{_parser.CurrentHiveType}";
            rootNode.NodeFont = ModernTheme.BoldFont;
            _treeView.Nodes.Add(rootNode);
            rootNode.Expand();
        }

        private TreeNode CreateTreeNode(RegistryKey key)
        {
            var node = new TreeNode(key.KeyName)
            {
                Tag = key,
                ImageKey = "folder",
                SelectedImageKey = "folder_open"
            };

            if (key.SubKeys?.Count > 0)
            {
                node.Nodes.Add(new TreeNode("Loading...") { Tag = "placeholder" });
            }

            return node;
        }

        private void TreeView_BeforeExpand(object? sender, TreeViewCancelEventArgs e)
        {
            var node = e.Node;
            if (node == null) return;

            if (node.Nodes.Count == 1 && node.Nodes[0].Tag?.ToString() == "placeholder")
            {
                node.Nodes.Clear();
                
                if (node.Tag is RegistryKey key && key.SubKeys != null)
                {
                    foreach (var subKey in key.SubKeys.OrderBy(k => k.KeyName))
                    {
                        node.Nodes.Add(CreateTreeNode(subKey));
                    }
                }
            }
        }

        private void TreeView_AfterSelect(object? sender, TreeViewEventArgs e)
        {
            // Reset previous node to closed folder icon
            // Check if node is still valid (in the tree) to prevent access issues
            if (_previousSelectedNode != null && 
                _previousSelectedNode != e.Node &&
                _previousSelectedNode.TreeView != null)
            {
                _previousSelectedNode.ImageKey = "folder";
                _previousSelectedNode.SelectedImageKey = "folder";
            }
            
            // Set current node to open folder icon
            if (e.Node != null)
            {
                e.Node.ImageKey = "folder_open";
                e.Node.SelectedImageKey = "folder_open";
                _previousSelectedNode = e.Node;
            }
            
            // Existing functionality
            if (e.Node?.Tag is RegistryKey key)
            {
                PopulateListView(key);
                UpdateDetailsForKey(key);
            }
        }

        private void PopulateListView(RegistryKey key)
        {
            _listView.Items.Clear();
            
            foreach (var value in key.Values.OrderBy(v => v.ValueName == "" ? "" : v.ValueName))
            {
                var name = string.IsNullOrEmpty(value.ValueName) ? "(Default)" : value.ValueName;
                var type = value.ValueType;
                var data = FormatValueData(value);
                
                var item = new ListViewItem(name, GetValueImageKey(value.ValueType));
                item.SubItems.Add(type);
                item.SubItems.Add(data);
                item.Tag = value;
                _listView.Items.Add(item);
            }
        }

        private string FormatValueData(KeyValue value)
        {
            try
            {
                if (value.ValueData == null) return "(null)";
                
                switch (value.ValueType?.ToUpperInvariant() ?? "")
                {
                    case "REGBINARY":
                        var bytes = value.ValueDataRaw;
                        if (bytes == null || bytes.Length == 0) return "(empty)";
                        var hex = BitConverter.ToString(bytes.Take(64).ToArray()).Replace("-", " ");
                        return bytes.Length > 64 ? $"{hex}... ({bytes.Length} bytes)" : hex;
                    
                    case "REGMULTISTRING":
                    case "REGMULTISZ":
                        return value.ValueData?.ToString()?.Replace("\0", " | ") ?? "";
                    
                    case "REGQWORD":
                    case "REGDWORD":
                        return $"{value.ValueData} (0x{Convert.ToInt64(value.ValueData):X})";
                    
                    default:
                        return value.ValueData?.ToString() ?? "";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error formatting value: {ex.Message}");
                return value.ValueData?.ToString() ?? "(error)";
            }
        }

        private void UpdateDetailsForKey(RegistryKey key)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"Path: {key.KeyPath}");
            sb.AppendLine($"Name: {key.KeyName}");
            sb.AppendLine($"Last Modified: {key.LastWriteTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "Unknown"}");
            sb.AppendLine($"Subkeys: {key.SubKeys?.Count ?? 0}");
            sb.AppendLine($"Values: {key.Values.Count}");
            
            _detailsBox.Text = sb.ToString();
        }

        private void ListView_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (_listView.SelectedItems.Count > 0 && _listView.SelectedItems[0].Tag is KeyValue value)
            {
                ShowValueDetails(value);
            }
        }

        private void ListView_DoubleClick(object? sender, EventArgs e)
        {
            if (_listView.SelectedItems.Count > 0 && _listView.SelectedItems[0].Tag is KeyValue value)
            {
                ShowValueInDialog(value);
            }
        }

        private void ShowValueDetails(KeyValue value)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"Name: {(string.IsNullOrEmpty(value.ValueName) ? "(Default)" : value.ValueName)}");
            sb.AppendLine($"Type: {value.ValueType}");
            sb.AppendLine($"Slack: {value.ValueSlack?.Length ?? 0} bytes");
            sb.AppendLine();
            sb.AppendLine("─── Data ───");
            sb.AppendLine(value.ValueData?.ToString() ?? "(null)");
            
            if (value.ValueDataRaw != null && value.ValueDataRaw.Length > 0)
            {
                sb.AppendLine();
                sb.AppendLine("─── Hex Dump ───");
                sb.AppendLine(FormatHexDump(value.ValueDataRaw));
            }
            
            _detailsBox.Text = sb.ToString();
        }

        private void ShowValueInDialog(KeyValue value)
        {
            using var form = new Form
            {
                Text = $"Value: {(string.IsNullOrEmpty(value.ValueName) ? "(Default)" : value.ValueName)}",
                Size = DpiHelper.ScaleSize(700, 550),
                StartPosition = FormStartPosition.CenterParent,
                ShowInTaskbar = false
            };
            ModernTheme.ApplyTo(form);

            var textBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                Text = FormatFullValueDetails(value),
                Padding = new Padding(12)
            };
            ModernTheme.ApplyTo(textBox);

            form.Controls.Add(textBox);
            form.ShowDialog(this);
        }

        private string FormatFullValueDetails(KeyValue value)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"Name: {(string.IsNullOrEmpty(value.ValueName) ? "(Default)" : value.ValueName)}");
            sb.AppendLine($"Type: {value.ValueType}");
            sb.AppendLine();
            sb.AppendLine("═══ Interpreted Data ═══");
            sb.AppendLine(value.ValueData?.ToString() ?? "(null)");
            
            if (value.ValueDataRaw != null && value.ValueDataRaw.Length > 0)
            {
                sb.AppendLine();
                sb.AppendLine("═══ Raw Hex Dump ═══");
                sb.AppendLine(FormatHexDump(value.ValueDataRaw));
            }
            
            return sb.ToString();
        }

        private string FormatHexDump(byte[] data)
        {
            var sb = new StringBuilder();
            int offset = 0;
            int bytesToShow = Math.Min(data.Length, 1024);
            
            while (offset < bytesToShow)
            {
                sb.Append($"{offset:X8}  ");
                
                for (int i = 0; i < 16; i++)
                {
                    if (offset + i < bytesToShow)
                        sb.Append($"{data[offset + i]:X2} ");
                    else
                        sb.Append("   ");
                    if (i == 7) sb.Append(" ");
                }
                
                sb.Append(" ");
                
                for (int i = 0; i < 16 && offset + i < bytesToShow; i++)
                {
                    byte b = data[offset + i];
                    sb.Append(b >= 32 && b < 127 ? (char)b : '.');
                }
                
                sb.AppendLine();
                offset += 16;
            }
            
            if (data.Length > bytesToShow)
                sb.AppendLine($"... ({data.Length - bytesToShow} more bytes)");
            
            return sb.ToString();
        }

        private void ExportKey_Click(object? sender, EventArgs e)
        {
            if (_treeView.SelectedNode?.Tag is not RegistryKey key)
            {
                ShowInfo("Please select a key to export.");
                return;
            }

            using var dialog = new SaveFileDialog
            {
                Title = "Export Registry Key",
                Filter = "Text Files|*.txt|All Files|*.*",
                FileName = $"{key.KeyName}_export.txt"
            };

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    var sb = new StringBuilder();
                    ExportKeyRecursive(key, sb, 0);
                    File.WriteAllText(dialog.FileName, sb.ToString());
                    ShowInfo("Export completed successfully.");
                }
                catch (Exception ex)
                {
                    ShowError($"Export failed: {ex.Message}");
                }
            }
        }

        private void AdjustValuesColumns()
        {
            try
            {
                if (_listView == null || _listView.Columns.Count < 3)
                    return;

                // Keep Name and Type fixed, expand Data to fill remaining space
                int clientWidth = _listView.ClientSize.Width;
                int fixedWidth = _listView.Columns[0].Width + _listView.Columns[1].Width;
                int padding = 4; // minimal padding
                int remaining = clientWidth - fixedWidth - padding;
                _listView.Columns[2].Width = Math.Max(150, remaining);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error adjusting columns: {ex.Message}");
            }
        }

        // Pre-computed indent strings for export (avoids per-call allocations)
        private static readonly string[] _indentCache = Enumerable.Range(0, 50).Select(i => new string(' ', i * 2)).ToArray();
        
        private static string GetIndent(int indent)
        {
            if (indent < _indentCache.Length)
                return _indentCache[indent];
            return new string(' ', indent * 2); // Fallback for deeply nested keys
        }

        private void ExportKeyRecursive(RegistryKey key, StringBuilder sb, int indent)
        {
            string indentStr = GetIndent(indent);
            sb.AppendLine($"{indentStr}[{key.KeyPath}]");
            sb.AppendLine($"{indentStr}Last Write: {key.LastWriteTime}");
            
            string valueIndent = GetIndent(indent) + "  ";
            foreach (var value in key.Values)
            {
                var name = string.IsNullOrEmpty(value.ValueName) ? "@" : $"\"{value.ValueName}\"";
                sb.AppendLine($"{valueIndent}{name} = {value.ValueData} ({value.ValueType})");
            }
            
            sb.AppendLine();
            
            if (key.SubKeys != null)
            {
                foreach (var subKey in key.SubKeys)
                {
                    ExportKeyRecursive(subKey, sb, indent + 1);
                }
            }
        }

        private void ShowStatistics_Click(object? sender, EventArgs e)
        {
            if (_parser == null || !_parser.IsLoaded)
            {
                ShowInfo("Please load a registry hive first.");
                return;
            }

            ShowStatisticsDialog();
        }

        private void ShowCompare_Click(object? sender, EventArgs e)
        {
            // Close existing compare form if open
            if (_compareForm != null && !_compareForm.IsDisposed)
            {
                _compareForm.Close();
            }
            
            _compareForm = new CompareForm();
            _compareForm.Icon = this.Icon;
            _compareForm.FormClosed += (s, ev) => { _compareForm = null; };
            
            // If a hive is already loaded, pre-load it as the left (base) hive
            if (_parser != null && _parser.IsLoaded && !string.IsNullOrEmpty(_currentHivePath))
            {
                _compareForm.SetLeftHive(_currentHivePath);
            }
            
            // Show as non-modal so user can still interact with main form
            _compareForm.Show();
        }

        private void ShowTimeline_Click(object? sender, EventArgs e)
        {
            if (_parser == null || !_parser.IsLoaded)
            {
                ShowInfo("Please load a registry hive first.");
                return;
            }

            // Close existing timeline form if open, or bring to front
            if (_timelineForm != null && !_timelineForm.IsDisposed)
            {
                _timelineForm.BringToFront();
                _timelineForm.Activate();
                return;
            }

            _timelineForm = new TimelineForm(_parser, this);
            _timelineForm.FormClosed += (s, ev) => { _timelineForm = null; };
            _timelineForm.Show();
        }

        private void ShowStatisticsDialog()
        {
            if (_parser == null) return;
            
            // Close existing statistics form if open
            if (_statisticsForm != null && !_statisticsForm.IsDisposed)
            {
                _statisticsForm.Close();
            }

            var form = new Form
            {
                Text = $"Registry Statistics - {_parser.CurrentHiveType}",
                Size = DpiHelper.ScaleSize(900, 650),
                StartPosition = FormStartPosition.CenterScreen,
                MinimumSize = DpiHelper.ScaleSize(700, 500),
                ShowInTaskbar = true
            };
            ModernTheme.ApplyTo(form);
            _statisticsForm = form; // Track the form;

            // Main panel with padding
            var mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ModernTheme.Background,
                Padding = new Padding(20)
            };

            // Header section
            var headerPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = DpiHelper.Scale(110),  // Increased for stat cards
                BackColor = ModernTheme.Surface
            };
            headerPanel.Paint += (s, ev) =>
            {
                using var pen = new Pen(ModernTheme.Border);
                ev.Graphics.DrawRectangle(pen, 0, 0, headerPanel.Width - 1, headerPanel.Height - 1);
            };

            var stats = _parser.GetStatistics();
            
            // Summary labels in header
            var titleLabel = new Label
            {
                Text = "\uE9D9  Registry Statistics",
                Font = ModernTheme.HeaderFont,
                ForeColor = ModernTheme.TextPrimary,
                Location = DpiHelper.ScalePoint(20, 15),
                AutoSize = true
            };
            titleLabel.Paint += (s, ev) =>
            {
                var iconFont = _iconFont16;
                using var iconBrush = new SolidBrush(ModernTheme.Accent);
                ev.Graphics.DrawString("\uE9D9", iconFont, iconBrush, 0, 2);
            };

            var summaryFlow = new FlowLayoutPanel
            {
                Location = DpiHelper.ScalePoint(20, 50),
                Size = DpiHelper.ScaleSize(850, 50),  // Increased height for stat cards
                FlowDirection = FlowDirection.LeftToRight,
                BackColor = Color.Transparent
            };

            summaryFlow.Controls.Add(CreateStatCard("File Size", stats.FormattedFileSize, ModernTheme.Accent));
            summaryFlow.Controls.Add(CreateStatCard("Total Keys", stats.TotalKeys.ToString("N0"), ModernTheme.Success));
            summaryFlow.Controls.Add(CreateStatCard("Total Values", stats.TotalValues.ToString("N0"), ModernTheme.Warning));
            summaryFlow.Controls.Add(CreateStatCard("Hive Type", stats.HiveType.ToString(), ModernTheme.Info));

            headerPanel.Controls.Add(titleLabel);
            headerPanel.Controls.Add(summaryFlow);

            // Tab control for different views
            var tabControl = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = ModernTheme.RegularFont
            };
            ModernTheme.ApplyTo(tabControl);

            // Key Count Tab
            var keyCountTab = new TabPage("Key Counts (Bloat Detection)")
            {
                BackColor = ModernTheme.Background,
                Padding = new Padding(10)
            };

            // Key Size Tab (shows data content size, not including registry structure overhead)
            var keySizeTab = new TabPage("Data Sizes")
            {
                BackColor = ModernTheme.Background,
                Padding = new Padding(10)
            };

            // Analyze registry in background
            var statusLabel = new Label
            {
                Text = "Analyzing registry structure...",
                Font = ModernTheme.RegularFont,
                ForeColor = ModernTheme.TextSecondary,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter
            };
            keyCountTab.Controls.Add(statusLabel);

            var statusLabel2 = new Label
            {
                Text = "Analyzing registry structure...",
                Font = ModernTheme.RegularFont,
                ForeColor = ModernTheme.TextSecondary,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter
            };
            keySizeTab.Controls.Add(statusLabel2);

            tabControl.TabPages.Add(keyCountTab);
            tabControl.TabPages.Add(keySizeTab);

            mainPanel.Controls.Add(tabControl);
            mainPanel.Controls.Add(headerPanel);
            form.Controls.Add(mainPanel);

            form.FormClosed += (s, ev) => { _statisticsForm = null; };
            form.Show();

            // Capture parser reference to avoid race condition if user loads new hive
            var parser = _parser;
            if (parser == null) return;

            // Analyze in background - only top level initially for speed
            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    var keyStats = AnalyzeTopLevelKeys(parser);
                    
                    if (form.IsDisposed) return;
                    
                    form.Invoke(() =>
                    {
                        if (form.IsDisposed) return;
                        
                        keyCountTab.Controls.Clear();
                        keySizeTab.Controls.Clear();
                        
                        // Create expandable Key Count tree panel
                        var keyCountPanel = CreateExpandableStatsPanel(
                            keyStats.OrderByDescending(k => k.SubKeyCount).ToList(),
                            "Subkey Count",
                            k => k.SubKeyCount,
                            ModernTheme.Accent);
                        keyCountPanel.Dock = DockStyle.Fill;
                        keyCountTab.Controls.Add(keyCountPanel);

                        // Create expandable Key Size tree panel
                        var keySizePanel = CreateExpandableStatsPanel(
                            keyStats.OrderByDescending(k => k.TotalSize).ToList(),
                            "Size (bytes)",
                            k => k.TotalSize,
                            ModernTheme.Success);
                        keySizePanel.Dock = DockStyle.Fill;
                        keySizeTab.Controls.Add(keySizePanel);
                    });
                }
                catch (ObjectDisposedException) { }
                catch (InvalidOperationException) { }
            });
        }

        private Panel CreateExpandableStatsPanel(List<KeyStatistics> data, string valueLabel, Func<KeyStatistics, long> valueSelector, Color barColor)
        {
            var panel = new Panel
            {
                BackColor = ModernTheme.Background,
                Padding = new Padding(5)
            };

            if (data.Count == 0)
            {
                var emptyLabel = new Label
                {
                    Text = "No data to display",
                    Dock = DockStyle.Fill,
                    TextAlign = ContentAlignment.MiddleCenter,
                    ForeColor = ModernTheme.TextSecondary,
                    Font = ModernTheme.RegularFont
                };
                panel.Controls.Add(emptyLabel);
                return panel;
            }

            long maxValue = data.Max(d => valueSelector(d));
            if (maxValue == 0) maxValue = 1;

            // Create TreeView for expandable display
            var tree = new TreeView
            {
                Dock = DockStyle.Fill,
                BackColor = ModernTheme.Background,
                ForeColor = ModernTheme.TextPrimary,
                Font = ModernTheme.DataFont,
                BorderStyle = BorderStyle.None,
                ShowLines = true,
                ShowPlusMinus = true,
                ShowRootLines = true,
                FullRowSelect = true,
                ItemHeight = DpiHelper.Scale(26),
                DrawMode = TreeViewDrawMode.OwnerDrawText
            };

            // Store value info for custom drawing
            var nodeValues = new Dictionary<TreeNode, (long value, long maxVal, Color color)>();

            // Add top-level nodes
            foreach (var item in data)
            {
                var node = new TreeNode(item.KeyPath)
                {
                    Tag = item.KeyPath // Store path for lazy loading
                };
                nodeValues[node] = (valueSelector(item), maxValue, barColor);
                
                // Add dummy child to show expand button (will be replaced on expand)
                // SubKeyCount includes the key itself, so > 1 means it has actual children
                if (item.SubKeyCount > 1)
                {
                    node.Nodes.Add(new TreeNode("Loading...") { Tag = "dummy" });
                }
                
                tree.Nodes.Add(node);
            }

            // Custom draw for bar chart effect (OwnerDrawText: system draws +/- and indentation,
            // we draw the text area with key name, bar chart, and value)
            tree.DrawNode += (s, e) =>
            {
                if (e.Node == null || e.Bounds.IsEmpty) return;

                var g = e.Graphics;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                
                // Background - fill from text bounds onward (system draws +/- buttons to the left)
                var fullBounds = new Rectangle(e.Bounds.X, e.Bounds.Y, tree.ClientSize.Width - e.Bounds.X, e.Bounds.Height);
                var bgColor = e.Node.IsSelected ? ModernTheme.Selection : 
                             (e.Node.Index % 2 == 0 ? ModernTheme.Background : ModernTheme.ListViewAltRow);
                using var bgBrush = new SolidBrush(bgColor);
                g.FillRectangle(bgBrush, fullBounds);

                // Key path text - starts at e.Bounds.X (system already handled indentation)
                int textX = e.Bounds.X;
                int maxTextWidth = Math.Min(DpiHelper.Scale(400), tree.ClientSize.Width / 2 - textX);
                var displayText = e.Node.Text;
                if (displayText.Length > 50)
                    displayText = "..." + displayText.Substring(displayText.Length - 47);
                
                using var textBrush = new SolidBrush(e.Node.IsSelected ? ModernTheme.Accent : ModernTheme.TextPrimary);
                var textRect = new RectangleF(textX, e.Bounds.Y + DpiHelper.Scale(4), maxTextWidth, e.Bounds.Height - DpiHelper.Scale(8));
                using var sf = new StringFormat { Trimming = StringTrimming.EllipsisPath };
                g.DrawString(displayText, ModernTheme.DataFont, textBrush, textRect, sf);

                // Draw bar and value if this node has stats
                if (nodeValues.TryGetValue(e.Node, out var valInfo))
                {
                    int barGap = DpiHelper.Scale(20);
                    int barStartX = textX + maxTextWidth + barGap;
                    int barEndMargin = DpiHelper.Scale(100);
                    int barMaxWidth = Math.Max(DpiHelper.Scale(100), tree.ClientSize.Width - barStartX - barEndMargin);
                    int barWidth = (int)((double)valInfo.value / valInfo.maxVal * barMaxWidth);
                    if (barWidth < DpiHelper.Scale(3)) barWidth = DpiHelper.Scale(3);

                    // Bar - vertically centered
                    int barHeight = DpiHelper.Scale(14);
                    var barRect = new Rectangle(barStartX, e.Bounds.Y + (e.Bounds.Height - barHeight) / 2, barWidth, barHeight);
                    using var barBrush = new SolidBrush(valInfo.color);
                    g.FillRectangle(barBrush, barRect);

                    // Value text
                    var valueText = FormatSize(valInfo.value);
                    using var valueBrush = new SolidBrush(ModernTheme.TextSecondary);
                    g.DrawString(valueText, ModernTheme.DataFont, valueBrush, barStartX + barMaxWidth + DpiHelper.Scale(10), e.Bounds.Y + (e.Bounds.Height - ModernTheme.DataFont.Height) / 2);
                }
            };

            // Lazy load children on expand
            tree.BeforeExpand += (s, e) =>
            {
                if (e.Node == null) return;
                
                // Check if this is a "more items" node
                if (e.Node.Tag is List<KeyStatistics> moreItems)
                {
                    var oldCursor = tree.Cursor;
                    tree.Cursor = Cursors.WaitCursor;
                    
                    try
                    {
                        // Cancel the expand — we're adding siblings, not children
                        e.Cancel = true;
                        
                        // Determine the parent node collection and index of the "more" node
                        var parentNodes = e.Node.Parent?.Nodes ?? tree.Nodes;
                        int moreIndex = e.Node.Index;
                        
                        long childMax = moreItems.Max(c => valueLabel == "Size (bytes)" ? c.TotalSize : c.SubKeyCount);
                        if (childMax == 0) childMax = 1;
                        
                        // Show next 100 items
                        var batch = moreItems.Take(100).ToList();
                        var remaining = moreItems.Skip(100).ToList();
                        
                        // Remove the "more" node
                        parentNodes.RemoveAt(moreIndex);
                        
                        // Insert batch items at the same level as siblings
                        int insertIndex = moreIndex;
                        foreach (var child in batch)
                        {
                            var childValue = valueLabel == "Size (bytes)" ? child.TotalSize : (long)child.SubKeyCount;
                            var childNode = new TreeNode(child.KeyPath.Contains('\\') 
                                ? child.KeyPath.Substring(child.KeyPath.LastIndexOf('\\') + 1)
                                : child.KeyPath)
                            {
                                Tag = child.KeyPath
                            };
                            nodeValues[childNode] = (childValue, childMax, barColor);
                            
                            // SubKeyCount includes the key itself, so > 1 means it has actual children
                            if (child.SubKeyCount > 1)
                            {
                                childNode.Nodes.Add(new TreeNode("Loading...") { Tag = "dummy" });
                            }
                            
                            parentNodes.Insert(insertIndex++, childNode);
                        }
                        
                        // Add another "more" node at the same level if there are still remaining items
                        if (remaining.Count > 0)
                        {
                            var moreNode = new TreeNode($"... and {remaining.Count} more")
                            {
                                ForeColor = ModernTheme.Accent,
                                Tag = remaining // Store remaining items for next expansion
                            };
                            moreNode.Nodes.Add(new TreeNode("Loading...") { Tag = "dummy" });
                            parentNodes.Insert(insertIndex, moreNode);
                        }
                    }
                    finally
                    {
                        tree.Cursor = oldCursor;
                    }
                    return;
                }
                
                // Check if we need to load children (has dummy node)
                if (e.Node.Nodes.Count == 1 && e.Node.Nodes[0].Tag as string == "dummy")
                {
                    var keyPath = e.Node.Tag as string;
                    if (string.IsNullOrEmpty(keyPath)) return;

                    // Show wait cursor during load
                    var oldCursor = tree.Cursor;
                    tree.Cursor = Cursors.WaitCursor;
                    
                    try
                    {
                        // Load children synchronously (faster perceived performance than async)
                        var childStats = GetChildKeyStats(keyPath, valueLabel == "Size (bytes)");
                        
                        e.Node.Nodes.Clear();
                        
                        if (childStats.Count == 0)
                        {
                            e.Node.Nodes.Add(new TreeNode("(No subkeys)") { ForeColor = ModernTheme.TextDisabled });
                            return;
                        }

                        var sortedChildren = valueLabel == "Size (bytes)" 
                            ? childStats.OrderByDescending(c => c.TotalSize).ToList()
                            : childStats.OrderByDescending(c => c.SubKeyCount).ToList();

                        long childMax = sortedChildren.Max(c => valueLabel == "Size (bytes)" ? c.TotalSize : c.SubKeyCount);
                        if (childMax == 0) childMax = 1;

                        foreach (var child in sortedChildren.Take(50)) // Show first 50
                        {
                            var childValue = valueLabel == "Size (bytes)" ? child.TotalSize : (long)child.SubKeyCount;
                            var childNode = new TreeNode(child.KeyPath.Contains('\\') 
                                ? child.KeyPath.Substring(child.KeyPath.LastIndexOf('\\') + 1)
                                : child.KeyPath)
                            {
                                Tag = child.KeyPath
                            };
                            nodeValues[childNode] = (childValue, childMax, barColor);
                            
                            // Add dummy for further expansion if has subkeys
                            // SubKeyCount includes the key itself, so > 1 means it has actual children
                            if (child.SubKeyCount > 1)
                            {
                                childNode.Nodes.Add(new TreeNode("Loading...") { Tag = "dummy" });
                            }
                            
                            e.Node.Nodes.Add(childNode);
                        }
                        
                        // Add expandable "more" node if there are remaining items
                        if (sortedChildren.Count > 50)
                        {
                            var remaining = sortedChildren.Skip(50).ToList();
                            var moreNode = new TreeNode($"... and {remaining.Count} more")
                            {
                                ForeColor = ModernTheme.Accent,
                                Tag = remaining // Store remaining items for lazy loading
                            };
                            moreNode.Nodes.Add(new TreeNode("Loading...") { Tag = "dummy" });
                            e.Node.Nodes.Add(moreNode);
                        }
                    }
                    finally
                    {
                        tree.Cursor = oldCursor;
                    }
                }
            };

            // Single ToolTip for the tree - created once, reused
            var treeToolTip = new ToolTip();
            
            // Tooltip for full path
            tree.NodeMouseHover += (s, e) =>
            {
                if (e.Node?.Tag is string path)
                {
                    if (nodeValues.TryGetValue(e.Node, out var valInfo))
                    {
                        treeToolTip.SetToolTip(tree, $"{path}\n{valueLabel}: {valInfo.value:N0}");
                    }
                    else
                    {
                        treeToolTip.SetToolTip(tree, path);
                    }
                }
            };
            
            // Dispose tooltip when panel is disposed
            panel.Disposed += (s, e) => treeToolTip.Dispose();

            panel.Controls.Add(tree);
            return panel;
        }

        private List<KeyStatistics> AnalyzeTopLevelKeys(OfflineRegistryParser parser)
        {
            var results = new List<KeyStatistics>();
            var rootKey = parser.GetRootKey();
            if (rootKey?.SubKeys == null) return results;

            foreach (var topKey in rootKey.SubKeys)
            {
                // Combined single-pass traversal for all three metrics
                var (subKeyCount, valueCount, totalSize) = CalculateKeyStatisticsRecursive(topKey);
                var stat = new KeyStatistics
                {
                    KeyPath = topKey.KeyName,
                    SubKeyCount = subKeyCount,
                    ValueCount = valueCount,
                    TotalSize = totalSize
                };
                results.Add(stat);
            }

            return results;
        }

        private List<KeyStatistics> GetChildKeyStats(string parentPath, bool calculateSize)
        {
            var results = new List<KeyStatistics>();
            if (_parser == null) return results;
            
            var parentKey = _parser.GetKey(parentPath);
            if (parentKey?.SubKeys == null) return results;

            foreach (var subKey in parentKey.SubKeys)
            {
                var childPath = $"{parentPath}\\{subKey.KeyName}";
                // Get the fully loaded key to ensure SubKeys are populated
                var fullChildKey = _parser.GetKey(childPath);
                
                var (subKeyCount, valueCount, totalSize) = CalculateKeyStatisticsRecursive(fullChildKey);
                var stat = new KeyStatistics
                {
                    KeyPath = childPath,
                    SubKeyCount = subKeyCount,
                    ValueCount = valueCount,
                    TotalSize = calculateSize ? totalSize : 0
                };
                results.Add(stat);
            }

            return results;
        }

        private string FormatSize(long bytes)
        {
            if (bytes >= 1_000_000_000) return $"{bytes / 1_000_000_000.0:F1}G";
            if (bytes >= 1_000_000) return $"{bytes / 1_000_000.0:F1}M";
            if (bytes >= 1_000) return $"{bytes / 1_000.0:F1}K";
            return bytes.ToString("N0");
        }

        private Panel CreateStatCard(string label, string value, Color accentColor)
        {
            var panel = new Panel
            {
                Size = DpiHelper.ScaleSize(180, 45),  // Increased height for DPI scaling
                Margin = DpiHelper.ScalePadding(0, 0, 15, 0),
                BackColor = Color.Transparent
            };
            
            panel.Paint += (s, e) =>
            {
                var g = e.Graphics;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                
                // Draw accent line
                using var accentBrush = new SolidBrush(accentColor);
                g.FillRectangle(accentBrush, 0, 0, DpiHelper.Scale(3), panel.Height);
                
                // Draw label
                int textX = DpiHelper.Scale(10);
                using var labelBrush = new SolidBrush(ModernTheme.TextSecondary);
                g.DrawString(label, ModernTheme.SmallFont, labelBrush, textX, DpiHelper.Scale(2));
                
                // Draw value - position below label
                using var valueBrush = new SolidBrush(ModernTheme.TextPrimary);
                g.DrawString(value, ModernTheme.BoldFont, valueBrush, textX, DpiHelper.Scale(18));
            };
            
            return panel;
        }

        private class KeyStatistics
        {
            public string KeyPath { get; set; } = "";
            public int SubKeyCount { get; set; }
            public int ValueCount { get; set; }
            public long TotalSize { get; set; }
        }

        /// <summary>
        /// Single-pass recursive traversal that calculates subkey count, value count, and total size at once.
        /// </summary>
        private (int subKeyCount, int valueCount, long totalSize) CalculateKeyStatisticsRecursive(RegistryParser.Abstractions.RegistryKey? key)
        {
            if (key == null) return (0, 0, 0);
            
            // Calculate size for this key - only actual data, no overhead
            long size = ((long)(key.KeyName?.Length ?? 0)) * 2;  // Key name only
            int valueCount = key.Values?.Count ?? 0;
            
            // Add size of values
            if (key.Values != null)
            {
                foreach (var val in key.Values)
                {
                    size += ((long)(val.ValueName?.Length ?? 0)) * 2;
                    size += GetValueDataSize(val);
                }
            }
            
            // Leaf key (no children)
            if (key.SubKeys == null || key.SubKeys.Count == 0)
            {
                return (1, valueCount, size);
            }
            
            // Non-leaf: recursively accumulate from children (start at 1 to count this key itself)
            int subKeyCount = 1;
            foreach (var subKey in key.SubKeys)
            {
                var (childSubKeys, childValues, childSize) = CalculateKeyStatisticsRecursive(subKey);
                subKeyCount += childSubKeys;
                valueCount += childValues;
                size += childSize;
            }
            
            return (subKeyCount, valueCount, size);
        }

        private long GetValueDataSize(KeyValue val)
        {
            if (val.ValueData == null) return 0;
            
            // Try to get actual data size based on type
            var dataStr = val.ValueData.ToString() ?? "";
            
            // Check for binary data (hex string format from Registry library)
            if (val.ValueType == "RegBinary" || dataStr.Contains(" ") && IsHexString(dataStr))
            {
                // Binary data is typically shown as hex bytes separated by spaces
                var parts = dataStr.Split(new[] { ' ', '-' }, StringSplitOptions.RemoveEmptyEntries);
                return parts.Length;
            }
            
            // REG_DWORD
            if (val.ValueType == "RegDword")
                return 4;
            
            // REG_QWORD
            if (val.ValueType == "RegQword")
                return 8;
            
            // REG_SZ, REG_EXPAND_SZ - Unicode string
            if (val.ValueType == "RegSz" || val.ValueType == "RegExpandSz")
                return (dataStr.Length + 1) * 2; // Unicode + null terminator
            
            // REG_MULTI_SZ - Multiple strings
            if (val.ValueType == "RegMultiSz")
                return (dataStr.Length + 2) * 2; // Unicode + double null terminator
            
            // Default: estimate based on string representation
            return dataStr.Length;
        }

        private bool IsHexString(string str)
        {
            if (string.IsNullOrEmpty(str)) return false;
            var trimmed = str.Replace(" ", "").Replace("-", "");
            return trimmed.Length > 0 && trimmed.All(c => 
                (c >= '0' && c <= '9') || 
                (c >= 'A' && c <= 'F') || 
                (c >= 'a' && c <= 'f'));
        }

        private void Search_Click(object? sender, EventArgs e)
        {
            if (_parser == null || !_parser.IsLoaded)
            {
                ShowInfo("Please load a registry hive first.");
                return;
            }

            // Close existing search window if open
            if (_searchForm != null && !_searchForm.IsDisposed)
            {
                _searchForm.Close();
            }

            _searchForm = new SearchForm(_parser, this);
            _searchForm.FormClosed += (s, ev) => { _searchForm = null; };
            _searchForm.Show();
        }

        private void ShowAnalyzeDialog_Click(object? sender, EventArgs e)
        {
            if (_infoExtractor == null || _parser == null || !_parser.IsLoaded)
            {
                ShowInfo("Please load a registry hive first.");
                return;
            }

            // If existing analyze window is open, bring it to front instead of recreating
            if (_analyzeForm != null && !_analyzeForm.IsDisposed)
            {
                _analyzeForm.BringToFront();
                _analyzeForm.Activate();
                if (_analyzeForm.WindowState == FormWindowState.Minimized)
                    _analyzeForm.WindowState = FormWindowState.Normal;
                return;
            }

            var form = new Form
            {
                Text = $"Analyze Registry - {_parser.CurrentHiveType}",
                Size = DpiHelper.ScaleSize(1100, 700),
                StartPosition = FormStartPosition.CenterScreen,
                MinimumSize = DpiHelper.ScaleSize(900, 500),
                ShowInTaskbar = true
            };
            _analyzeForm = form; // Track the form
            form.Icon = this.Icon;
            ModernTheme.ApplyTo(form);

            // Create theme data to track controls
            var themeData = new AnalyzeFormThemeData();
            form.Tag = themeData;

            // Main split container
            var splitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterWidth = 3,
                BackColor = ModernTheme.Border,
                Panel1MinSize = 50,
                Panel2MinSize = 50,
                AccessibleName = "Analyze Registry",
                AccessibleRole = AccessibleRole.Pane
            };
            splitContainer.Panel1.BackColor = ModernTheme.Background;
            splitContainer.Panel2.BackColor = ModernTheme.Background;
            themeData.MainSplit = splitContainer;

            // Set splitter distance after form loads
            form.Load += (s, ev) =>
            {
                if (splitContainer.Width > 300)
                    splitContainer.SplitterDistance = DpiHelper.Scale(250);
            };

            // Left panel - Categories
            var leftPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ModernTheme.Background,
                Padding = new Padding(0),
                AccessibleName = "Category Panel",
                AccessibleRole = AccessibleRole.Pane
            };
            themeData.LeftPanel = leftPanel;

            var categoryHeader = new Panel
            {
                Dock = DockStyle.Top,
                Height = DpiHelper.Scale(50),
                BackColor = ModernTheme.Surface,
                Padding = DpiHelper.ScalePadding(15, 0, 15, 0)
            };
            themeData.CategoryHeader = categoryHeader;

            var categoryTitle = new Label
            {
                Text = "Categories",
                Font = ModernTheme.HeaderFont,
                ForeColor = ModernTheme.TextPrimary,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };
            categoryHeader.Controls.Add(categoryTitle);
            themeData.CategoryTitle = categoryTitle;

            var categoryList = new ListBox
            {
                Dock = DockStyle.Fill,
                BackColor = ModernTheme.TreeViewBack,
                ForeColor = ModernTheme.TextPrimary,
                Font = ModernTheme.RegularFont,
                BorderStyle = BorderStyle.None,
                ItemHeight = DpiHelper.Scale(48),
                DrawMode = DrawMode.OwnerDrawFixed,
                AccessibleName = "Categories",
                AccessibleRole = AccessibleRole.List
            };
            themeData.CategoryList = categoryList;

            // Add categories with PNG icons loaded from embedded resources
            var categoryKeys = new (string text, string key)[]
            {
                ("System", "System"),
                ("Profiles", "Profiles"),
                ("Services", "Services"),
                ("Software", "Software"),
                ("Storage", "Storage"),
                ("Network", "Network"),
                ("RDP", "RDP"),
                ("Update", "Update")
            };

            // Load category icon images from embedded resources
            var categoryIcons = new Dictionary<string, Image>();
            var assembly = System.Reflection.Assembly.GetExecutingAssembly();
            int iconDisplaySize = DpiHelper.Scale(32);
            foreach (var (text, key) in categoryKeys)
            {
                var resourceName = $"RegistryExpert.icons.category_{key.ToLowerInvariant()}.png";
                using var stream = assembly.GetManifestResourceStream(resourceName);
                if (stream != null)
                {
                    using var original = Image.FromStream(stream);
                    var scaled = new Bitmap(iconDisplaySize, iconDisplaySize);
                    using (var g = Graphics.FromImage(scaled))
                    {
                        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                        g.DrawImage(original, 0, 0, iconDisplaySize, iconDisplaySize);
                    }
                    categoryIcons[key] = scaled;
                }
            }
            themeData.CategoryIcons = categoryIcons;

            // Determine which categories are available based on hive type
            var currentHiveType = _parser.CurrentHiveType;
            var enabledCategories = new HashSet<string>();
            
            if (currentHiveType == OfflineRegistryParser.HiveType.SYSTEM)
            {
                enabledCategories.Add("System");
                enabledCategories.Add("Services");
                enabledCategories.Add("Storage");
                enabledCategories.Add("Network");
                enabledCategories.Add("RDP");
            }
            else if (currentHiveType == OfflineRegistryParser.HiveType.SOFTWARE)
            {
                enabledCategories.Add("Profiles");
                enabledCategories.Add("System"); // Activation subcategory is available in SOFTWARE hive
                enabledCategories.Add("Software"); // Installed Programs and Appx
                enabledCategories.Add("Update"); // Windows Update settings
            }
            else if (currentHiveType == OfflineRegistryParser.HiveType.COMPONENTS)
            {
                // COMPONENTS hive only enables Update category (for CBS Components)
                enabledCategories.Add("Update");
            }
            else if (currentHiveType == OfflineRegistryParser.HiveType.SAM)
            {
                // SAM hive contains user account information
                enabledCategories.Add("Profiles");
            }
            else if (currentHiveType == OfflineRegistryParser.HiveType.NTUSER)
            {
                // NTUSER.DAT contains user-specific settings
                enabledCategories.Add("Profiles");
                enabledCategories.Add("Software"); // User-specific software settings
            }
            // For BCD, SECURITY, USRCLASS, DEFAULT, AMCACHE, Unknown - no categories are enabled
            // as the Analyze feature doesn't have specific support for these hive types

            // Add categories with enabled ones first, then disabled ones
            var availableCats = new List<(string text, string key)>();
            var unavailableCats = new List<(string text, string key)>();
            foreach (var cat in categoryKeys)
            {
                if (enabledCategories.Contains(cat.key))
                    availableCats.Add(cat);
                else
                    unavailableCats.Add(cat);
            }
            foreach (var cat in availableCats)
                categoryList.Items.Add(cat);
            foreach (var cat in unavailableCats)
                categoryList.Items.Add(cat);

            // Font point sizes are already DPI-aware, so don't scale them
            var textFont = new Font("Segoe UI Semibold", 10.5F);
            themeData.CategoryTextFont = textFont;

            // Pre-calculate DPI-scaled values for drawing positions (to match scaled ItemHeight)
            int iconCenterXOffset = DpiHelper.Scale(32);
            int iconHalfSize = DpiHelper.Scale(16);
            int textXOffset = DpiHelper.Scale(60);
            int textRightPadding = DpiHelper.Scale(65);
            int accentBarWidth = DpiHelper.Scale(3);
            int accentBarPadding = DpiHelper.Scale(8);

            // Custom draw for category items - modern card-like style with PNG icons
            categoryList.DrawItem += (s, ev) =>
            {
                if (ev.Index < 0) return;
                var item = ((string text, string key))categoryList.Items[ev.Index];
                bool isEnabled = enabledCategories.Contains(item.key);
                
                bool isSelected = (ev.State & DrawItemState.Selected) == DrawItemState.Selected;
                
                // Background with subtle selection indicator
                Color backColor = isSelected && isEnabled 
                    ? ModernTheme.Selection 
                    : ModernTheme.TreeViewBack;
                
                using var brush = new SolidBrush(backColor);
                ev.Graphics.FillRectangle(brush, ev.Bounds);
                
                // Draw accent bar on left for selected item
                if (isSelected && isEnabled)
                {
                    using var accentBrush = new SolidBrush(ModernTheme.Accent);
                    ev.Graphics.FillRectangle(accentBrush, ev.Bounds.X, ev.Bounds.Y + accentBarPadding, accentBarWidth, ev.Bounds.Height - accentBarPadding * 2);
                }
                
                Color textColor = isEnabled 
                    ? (isSelected ? ModernTheme.Accent : ModernTheme.TextPrimary) 
                    : ModernTheme.TextDisabled;

                // Draw category icon image
                var iconCenterX = ev.Bounds.X + iconCenterXOffset;
                var iconCenterY = ev.Bounds.Y + ev.Bounds.Height / 2;
                
                if (categoryIcons.TryGetValue(item.key, out var iconImage))
                {
                    var iconX = iconCenterX - iconHalfSize;
                    var iconY = iconCenterY - iconHalfSize;
                    
                    if (isEnabled)
                    {
                        ev.Graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        ev.Graphics.DrawImage(iconImage, iconX, iconY, iconDisplaySize, iconDisplaySize);
                    }
                    else
                    {
                        // Draw greyed-out icon for disabled categories
                        using var grayAttributes = new System.Drawing.Imaging.ImageAttributes();
                        var grayMatrix = new System.Drawing.Imaging.ColorMatrix(new float[][]
                        {
                            new float[] { 0.3f, 0.3f, 0.3f, 0, 0 },
                            new float[] { 0.3f, 0.3f, 0.3f, 0, 0 },
                            new float[] { 0.3f, 0.3f, 0.3f, 0, 0 },
                            new float[] { 0, 0, 0, 0.4f, 0 },
                            new float[] { 0, 0, 0, 0, 1 }
                        });
                        grayAttributes.SetColorMatrix(grayMatrix);
                        ev.Graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        ev.Graphics.DrawImage(iconImage,
                            new Rectangle(iconX, iconY, iconDisplaySize, iconDisplaySize),
                            0, 0, iconImage.Width, iconImage.Height,
                            GraphicsUnit.Pixel, grayAttributes);
                    }
                }

                // Draw text
                using var textBrush = new SolidBrush(textColor);
                using var sf = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center };
                var textRect = new Rectangle(ev.Bounds.X + textXOffset, ev.Bounds.Y, ev.Bounds.Width - textRightPadding, ev.Bounds.Height);
                ev.Graphics.DrawString(item.text, textFont, textBrush, textRect, sf);
            };

            // Prevent selection of disabled categories
            categoryList.SelectedIndexChanged += (s, ev) =>
            {
                if (categoryList.SelectedIndex >= 0 && categoryList.SelectedItem != null)
                {
                    var selected = ((string text, string key))categoryList.SelectedItem;
                    if (!enabledCategories.Contains(selected.key))
                    {
                        // Find first enabled category
                        for (int i = 0; i < categoryList.Items.Count; i++)
                        {
                            var item = ((string text, string key))categoryList.Items[i];
                            if (enabledCategories.Contains(item.key))
                            {
                                categoryList.SelectedIndex = i;
                                return;
                            }
                        }
                        categoryList.SelectedIndex = -1;
                    }
                }
            };


            // Add tooltip for disabled categories to show required hive
            var categoryToolTip = new ToolTip
            {
                InitialDelay = 300,
                ReshowDelay = 100,
                AutoPopDelay = 5000,
                BackColor = ModernTheme.Surface,
                ForeColor = ModernTheme.TextPrimary
            };
            themeData.CategoryToolTip = categoryToolTip;
            int lastHoveredIndex = -1;
            
            // Map categories to their required hive types
            var categoryHiveRequirements = new Dictionary<string, string>
            {
                { "System", "SYSTEM" },
                { "Services", "SYSTEM" },
                { "Storage", "SYSTEM" },
                { "Network", "SYSTEM" },
                { "RDP", "SYSTEM" },
                { "Profiles", "SOFTWARE" },
                { "Software", "SOFTWARE" },
                { "Update", "SOFTWARE" }
            };
            
            categoryList.MouseMove += (s, ev) =>
            {
                int index = categoryList.IndexFromPoint(ev.Location);
                if (index != lastHoveredIndex)
                {
                    lastHoveredIndex = index;
                    if (index >= 0 && index < categoryList.Items.Count)
                    {
                        var item = ((string text, string key))categoryList.Items[index];
                        if (!enabledCategories.Contains(item.key) && categoryHiveRequirements.TryGetValue(item.key, out var requiredHive))
                        {
                            categoryToolTip.SetToolTip(categoryList, $"This feature requires the {requiredHive} hive to be loaded");
                        }
                        else
                        {
                            categoryToolTip.SetToolTip(categoryList, "");
                        }
                    }
                    else
                    {
                        categoryToolTip.SetToolTip(categoryList, "");
                    }
                }
            };
            
            categoryList.MouseLeave += (s, ev) =>
            {
                lastHoveredIndex = -1;
                categoryToolTip.SetToolTip(categoryList, "");
            };

            // Show notice when clicking a disabled (greyed-out) category
            categoryList.MouseDown += (s, ev) =>
            {
                int index = categoryList.IndexFromPoint(ev.Location);
                if (index >= 0 && index < categoryList.Items.Count)
                {
                    var item = ((string text, string key))categoryList.Items[index];
                    if (!enabledCategories.Contains(item.key) && categoryHiveRequirements.TryGetValue(item.key, out var requiredHive))
                    {
                        MessageBox.Show(
                            $"This feature requires {requiredHive} hive to be loaded.\n\n" +
                            $"Common location: C:\\Windows\\System32\\config\\{requiredHive}",
                            $"{requiredHive} Hive Required",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                }
            };

            leftPanel.Controls.Add(categoryList);
            leftPanel.Controls.Add(categoryHeader);

            // Right panel - Content
            var rightPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ModernTheme.Background,
                Padding = new Padding(0),
                AccessibleName = "Content Panel",
                AccessibleRole = AccessibleRole.Pane
            };
            themeData.RightPanel = rightPanel;

            // Forward declaration for contentDetailSplit (created later, but captured by lambdas)
            SplitContainer? contentDetailSplit = null;

            var contentHeader = new Panel
            {
                Dock = DockStyle.Top,
                Height = DpiHelper.Scale(50),
                BackColor = ModernTheme.Surface,
                Padding = DpiHelper.ScalePadding(15, 0, 15, 0)
            };
            themeData.ContentHeader = contentHeader;

            var contentTitle = new Label
            {
                Text = "Select a category",
                Font = ModernTheme.HeaderFont,
                ForeColor = ModernTheme.TextPrimary,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };
            contentHeader.Controls.Add(contentTitle);
            themeData.ContentTitle = contentTitle;

            // Subcategory tabs panel
            var subCategoryPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                MinimumSize = DpiHelper.ScaleSize(0, 48),
                MaximumSize = DpiHelper.ScaleSize(0, 140), // Allow up to 3 rows
                BackColor = ModernTheme.Surface,
                Padding = DpiHelper.ScalePadding(5, 10, 5, 10),
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                AccessibleName = "Subcategory Tabs",
                AccessibleRole = AccessibleRole.ToolBar
            };
            themeData.SubCategoryPanel = subCategoryPanel;

            // Appx filter panel (secondary filter row, shown when Appx Packages is selected)
            var appxFilterPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                MinimumSize = DpiHelper.ScaleSize(0, 0),
                BackColor = ModernTheme.TreeViewBack,
                Padding = DpiHelper.ScalePadding(10, 5, 5, 5),
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                Visible = false // Hidden by default
            };
            themeData.AppxFilterPanel = appxFilterPanel;

            // Main content area - DataGridView for better display
            var contentGrid = new DataGridView { Dock = DockStyle.Fill, AccessibleName = "Analysis Results" };
            ModernTheme.ApplyTo(contentGrid);
            contentGrid.ColumnHeadersDefaultCellStyle.BackColor = ModernTheme.Surface;  // Override
            contentGrid.ColumnHeadersDefaultCellStyle.SelectionBackColor = ModernTheme.Surface;
            contentGrid.ColumnHeadersDefaultCellStyle.Padding = DpiHelper.ScalePadding(5, 0, 5, 0);
            themeData.ContentGrid = contentGrid;

            // Network Interfaces master-detail panel
            var networkSplitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterWidth = 3,
                BackColor = ModernTheme.Border,
                Panel1MinSize = 100,
                Panel2MinSize = 100,
                Visible = false
            };
            networkSplitContainer.Panel1.BackColor = ModernTheme.Background;
            networkSplitContainer.Panel2.BackColor = ModernTheme.Background;
            themeData.NetworkSplit = networkSplitContainer;
            
            // Set splitter distance when visible
            networkSplitContainer.VisibleChanged += (s, ev) =>
            {
                if (networkSplitContainer.Visible && networkSplitContainer.Width > 300)
                {
                    try { networkSplitContainer.SplitterDistance = DpiHelper.Scale(250); } catch { }
                }
            };

            // Network adapters list (left)
            var networkAdaptersList = new ListBox
            {
                Dock = DockStyle.Fill,
                BackColor = ModernTheme.TreeViewBack,
                ForeColor = ModernTheme.TextPrimary,
                Font = ModernTheme.RegularFont,
                BorderStyle = BorderStyle.None,
                ItemHeight = DpiHelper.Scale(32),
                DrawMode = DrawMode.OwnerDrawFixed,
                AccessibleName = "Network Adapters",
                AccessibleRole = AccessibleRole.List
            };
            themeData.NetworkAdaptersList = networkAdaptersList;

            // Custom draw for adapter items
            networkAdaptersList.DrawItem += (s, ev) =>
            {
                if (ev.Index < 0) return;
                var item = networkAdaptersList.Items[ev.Index] as NetworkAdapterItem;
                if (item == null) return;

                bool isSelected = (ev.State & DrawItemState.Selected) == DrawItemState.Selected;
                Color backColor = isSelected ? ModernTheme.Accent : (ev.Index % 2 == 0 ? ModernTheme.TreeViewBack : ModernTheme.ListViewAltRow);
                Color textColor = isSelected ? Color.White : ModernTheme.TextPrimary;

                using var brush = new SolidBrush(backColor);
                ev.Graphics.FillRectangle(brush, ev.Bounds);

                using var textBrush = new SolidBrush(textColor);
                using var sf = new StringFormat { LineAlignment = StringAlignment.Center, Trimming = StringTrimming.EllipsisCharacter };
                var textRect = new Rectangle(ev.Bounds.X + DpiHelper.Scale(10), ev.Bounds.Y, ev.Bounds.Width - DpiHelper.Scale(10), ev.Bounds.Height);
                ev.Graphics.DrawString(item.DisplayName, ev.Font ?? networkAdaptersList.Font, textBrush, textRect, sf);
            };

            var networkAdaptersHeader = new Panel
            {
                Dock = DockStyle.Top,
                Height = DpiHelper.Scale(32),
                BackColor = ModernTheme.Surface,
                Padding = DpiHelper.ScalePadding(10, 0, 10, 0)
            };
            themeData.NetworkAdaptersHeader = networkAdaptersHeader;
            var networkAdaptersLabel = new Label
            {
                Text = "Adapters",
                Font = new Font("Segoe UI Semibold", 9F),
                ForeColor = ModernTheme.TextSecondary,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };
            networkAdaptersHeader.Controls.Add(networkAdaptersLabel);
            themeData.NetworkAdaptersLabel = networkAdaptersLabel;

            networkSplitContainer.Panel1.Controls.Add(networkAdaptersList);
            networkSplitContainer.Panel1.Controls.Add(networkAdaptersHeader);

            // Network details panel (right) - contains grid and registry info
            var networkDetailsPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ModernTheme.Background
            };

            var networkDetailsHeader = new Panel
            {
                Dock = DockStyle.Top,
                Height = DpiHelper.Scale(32),
                BackColor = ModernTheme.Surface,
                Padding = DpiHelper.ScalePadding(10, 0, 10, 0)
            };
            themeData.NetworkDetailsHeader = networkDetailsHeader;
            var networkDetailsLabel = new Label
            {
                Text = "Adapter Details",
                Font = new Font("Segoe UI Semibold", 9F),
                ForeColor = ModernTheme.TextSecondary,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };
            networkDetailsHeader.Controls.Add(networkDetailsLabel);
            themeData.NetworkDetailsLabel = networkDetailsLabel;

            var networkDetailsGrid = new DataGridView { Dock = DockStyle.Fill };
            ModernTheme.ApplyTo(networkDetailsGrid);
            themeData.NetworkDetailsGrid = networkDetailsGrid;

            networkDetailsPanel.Controls.Add(networkDetailsGrid);
            networkDetailsPanel.Controls.Add(networkDetailsHeader);

            networkSplitContainer.Panel2.Controls.Add(networkDetailsPanel);

            // ==================== Firewall Rules Panel ====================
            var firewallPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ModernTheme.Background,
                Visible = false
            };

            // Top panel - Profile buttons (fixed height)
            var firewallProfileButtonsPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = DpiHelper.Scale(44),
                BackColor = ModernTheme.Surface,
                Padding = DpiHelper.ScalePadding(10, 8, 10, 8),
                FlowDirection = FlowDirection.LeftToRight
            };

            var firewallProfileLabel = new Label
            {
                Text = "Select Profile:",
                ForeColor = ModernTheme.TextSecondary,
                Font = ModernTheme.BoldFont,
                AutoSize = true,
                Margin = DpiHelper.ScalePadding(0, 5, 10, 0)
            };
            firewallProfileButtonsPanel.Controls.Add(firewallProfileLabel);

            // Rules grid panel (fills remaining space)
            var firewallRulesPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ModernTheme.Background
            };

            var firewallRulesHeader = new Panel
            {
                Dock = DockStyle.Top,
                Height = DpiHelper.Scale(32),
                BackColor = ModernTheme.Surface,
                Padding = DpiHelper.ScalePadding(10, 0, 10, 0)
            };

            var firewallRulesLabel = new Label
            {
                Text = "Firewall Rules",
                Font = new Font("Segoe UI Semibold", 9F),
                ForeColor = ModernTheme.TextSecondary,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };
            firewallRulesHeader.Controls.Add(firewallRulesLabel);

            var firewallRulesGrid = new DataGridView { Dock = DockStyle.Fill };
            ModernTheme.ApplyTo(firewallRulesGrid);
            firewallRulesGrid.DefaultCellStyle.Padding = DpiHelper.ScalePadding(5, 1, 5, 1);  // Compact rows

            firewallRulesPanel.Controls.Add(firewallRulesGrid);
            firewallRulesPanel.Controls.Add(firewallRulesHeader);

            // Add controls to main firewall panel (order matters: Fill must be added first)
            firewallPanel.Controls.Add(firewallRulesPanel);
            firewallPanel.Controls.Add(firewallProfileButtonsPanel);

            // Register firewall controls for theme updates
            themeData.FirewallPanel = firewallPanel;
            themeData.FirewallProfileButtonsPanel = firewallProfileButtonsPanel;
            themeData.FirewallProfileLabel = firewallProfileLabel;
            themeData.FirewallRulesPanel = firewallRulesPanel;
            themeData.FirewallRulesHeader = firewallRulesHeader;
            themeData.FirewallRulesLabel = firewallRulesLabel;
            themeData.FirewallRulesGrid = firewallRulesGrid;

            // Firewall state
            var firewallProfileButtons = new List<Button>();
            themeData.FirewallProfileButtons = firewallProfileButtons;
            string currentFirewallProfile = "";
            List<FirewallRuleInfo> currentFirewallRules = new();

            // Function to display firewall rules for a profile
            Action<string, string> displayFirewallRules = (profileKey, profileDisplayName) =>
            {
                currentFirewallProfile = profileKey;
                firewallRulesLabel.Text = $"Firewall Rules - {profileDisplayName}";

                // Get rules for this profile
                currentFirewallRules = _infoExtractor.GetFirewallRulesForProfile(profileKey);

                // Sort: Active Block rules first, then Active Allow, then Inactive
                var sortedRules = currentFirewallRules
                    .OrderByDescending(r => r.IsActive && r.Action.Equals("Block", StringComparison.OrdinalIgnoreCase))
                    .ThenByDescending(r => r.IsActive)
                    .ThenBy(r => r.Name)
                    .ToList();

                // Populate grid
                firewallRulesGrid.Columns.Clear();
                firewallRulesGrid.Rows.Clear();

                firewallRulesGrid.Columns.Add("status", "Status");
                firewallRulesGrid.Columns.Add("action", "Action");
                firewallRulesGrid.Columns.Add("direction", "Dir");
                firewallRulesGrid.Columns.Add("name", "Name");
                firewallRulesGrid.Columns.Add("protocol", "Protocol");
                firewallRulesGrid.Columns.Add("lport", "Local Port");
                firewallRulesGrid.Columns.Add("rport", "Remote Port");
                firewallRulesGrid.Columns.Add("application", "Application");

                firewallRulesGrid.Columns["status"].FillWeight = 8;
                firewallRulesGrid.Columns["action"].FillWeight = 8;
                firewallRulesGrid.Columns["direction"].FillWeight = 8;
                firewallRulesGrid.Columns["name"].FillWeight = 28;
                firewallRulesGrid.Columns["protocol"].FillWeight = 10;
                firewallRulesGrid.Columns["lport"].FillWeight = 10;
                firewallRulesGrid.Columns["rport"].FillWeight = 10;
                firewallRulesGrid.Columns["application"].FillWeight = 18;

                foreach (var rule in sortedRules)
                {
                    var statusIcon = rule.IsActive ? "✅" : "⬜";
                    var actionDisplay = rule.Action.Equals("Block", StringComparison.OrdinalIgnoreCase) ? "🚫 Block" : "✓ Allow";
                    var dirDisplay = rule.Direction.Equals("Inbound", StringComparison.OrdinalIgnoreCase) ? "⬇ In" : "⬆ Out";
                    var localPort = !string.IsNullOrEmpty(rule.LocalPorts) ? rule.LocalPorts : "Any";
                    var remotePort = !string.IsNullOrEmpty(rule.RemotePorts) ? rule.RemotePorts : "Any";
                    var app = !string.IsNullOrEmpty(rule.Application) ? Path.GetFileName(rule.Application) : 
                             (!string.IsNullOrEmpty(rule.Service) ? $"[{rule.Service}]" : 
                             (!string.IsNullOrEmpty(rule.PackageFamilyName) ? rule.PackageFamilyName : ""));

                    var rowIndex = firewallRulesGrid.Rows.Add(statusIcon, actionDisplay, dirDisplay, rule.Name, rule.Protocol, localPort, remotePort, app);

                    // Color coding for block rules
                    if (rule.IsActive && rule.Action.Equals("Block", StringComparison.OrdinalIgnoreCase))
                    {
                        firewallRulesGrid.Rows[rowIndex].DefaultCellStyle.BackColor = ModernTheme.BlockRowBackground;
                        firewallRulesGrid.Rows[rowIndex].DefaultCellStyle.ForeColor = ModernTheme.BlockRowForeground;
                    }
                    else if (!rule.IsActive)
                    {
                        firewallRulesGrid.Rows[rowIndex].DefaultCellStyle.ForeColor = ModernTheme.TextDisabled;
                    }
                }

                // Update button states
                foreach (var btn in firewallProfileButtons)
                {
                    var btnProfile = btn.Tag?.ToString() ?? "";
                    btn.BackColor = btnProfile == profileKey ? ModernTheme.Accent : ModernTheme.Surface;
                    btn.ForeColor = btnProfile == profileKey ? Color.White : ModernTheme.TextPrimary;
                }
            };

            // Store refresh action for theme changes
            themeData.RefreshFirewallDisplay = () =>
            {
                if (!string.IsNullOrEmpty(currentFirewallProfile) && firewallPanel.Visible)
                {
                    // Re-apply colors to existing rows
                    var sortedRules = currentFirewallRules
                        .OrderByDescending(r => r.IsActive && r.Action.Equals("Block", StringComparison.OrdinalIgnoreCase))
                        .ThenByDescending(r => r.IsActive)
                        .ThenBy(r => r.Name)
                        .ToList();

                    for (int i = 0; i < firewallRulesGrid.Rows.Count && i < sortedRules.Count; i++)
                    {
                        var rule = sortedRules[i];
                        if (rule.IsActive && rule.Action.Equals("Block", StringComparison.OrdinalIgnoreCase))
                        {
                            firewallRulesGrid.Rows[i].DefaultCellStyle.BackColor = ModernTheme.BlockRowBackground;
                            firewallRulesGrid.Rows[i].DefaultCellStyle.ForeColor = ModernTheme.BlockRowForeground;
                        }
                        else if (!rule.IsActive)
                        {
                            firewallRulesGrid.Rows[i].DefaultCellStyle.BackColor = ModernTheme.Surface;
                            firewallRulesGrid.Rows[i].DefaultCellStyle.ForeColor = ModernTheme.TextDisabled;
                        }
                        else
                        {
                            firewallRulesGrid.Rows[i].DefaultCellStyle.BackColor = ModernTheme.Surface;
                            firewallRulesGrid.Rows[i].DefaultCellStyle.ForeColor = ModernTheme.TextPrimary;
                        }
                    }
                }
            };

            // Cache for network adapters
            List<NetworkAdapterItem> networkAdaptersCache = new();
            NetworkAdapterItem? selectedNetworkAdapter = null;

            // Network adapter and details selection handlers will be set up after registryPathLabel/registryValueBox are created

            // Track current state
            List<AnalysisSection> currentSections = new();
            AnalysisSection? currentSection = null;
            var subCategoryButtons = new List<Button>();
            themeData.SubCategoryButtons = subCategoryButtons;

            // Appx display state (declared before displaySection which uses these)
            var appxFilterButtons = new List<Button>();
            themeData.AppxFilterButtons = appxFilterButtons;
            string currentAppxFilter = "InBox";

            // Function to display Appx packages with filter
            Action<string> displayAppxWithFilter = null!;
            displayAppxWithFilter = (filter) =>
            {
                currentAppxFilter = filter;
                contentGrid.Columns.Clear();
                contentGrid.Rows.Clear();

                contentGrid.Columns.Add("name", "Package Name");
                contentGrid.Columns.Add("version", "Version");
                contentGrid.Columns.Add("arch", "Architecture");
                contentGrid.Columns["name"].FillWeight = 55;
                contentGrid.Columns["version"].FillWeight = 25;
                contentGrid.Columns["arch"].FillWeight = 20;

                var packages = _infoExtractor.GetAppxPackages(filter);

                foreach (var pkg in packages)
                {
                    var rowIdx = contentGrid.Rows.Add(pkg.PackageName, pkg.Version, pkg.Architecture);
                    contentGrid.Rows[rowIdx].Tag = pkg;  // Store package info for SelectionChanged
                }

                // Update filter button states
                foreach (var btn in appxFilterButtons)
                {
                    var btnFilter = btn.Tag?.ToString() ?? "";
                    btn.BackColor = btnFilter == filter ? ModernTheme.Accent : ModernTheme.Surface;
                    btn.ForeColor = btnFilter == filter ? Color.White : ModernTheme.TextPrimary;
                }
            };

            // Function to create Appx filter buttons
            Action createAppxFilterButtons = null!;
            createAppxFilterButtons = () =>
            {
                appxFilterPanel.Controls.Clear();
                appxFilterButtons.Clear();

                var inboxPackages = _infoExtractor.GetAppxPackages("InBox");
                var userPackages = _infoExtractor.GetAppxPackages("UserInstalled");

                var filters = new[]
                {
                    ($"📦 InBox Preinstalled ({inboxPackages.Count})", "InBox"),
                    ($"📲 User Installed ({userPackages.Count})", "UserInstalled")
                };

                foreach (var (text, filterKey) in filters)
                {
                    var btn = new Button
                    {
                        Text = text,
                        FlatStyle = FlatStyle.Flat,
                        BackColor = filterKey == "InBox" ? ModernTheme.Accent : ModernTheme.Surface,
                        ForeColor = filterKey == "InBox" ? Color.White : ModernTheme.TextPrimary,
                        Font = ModernTheme.RegularFont,
                        Height = DpiHelper.Scale(28),
                        AutoSize = true,
                        Padding = DpiHelper.ScalePadding(8, 0, 8, 0),
                        Margin = DpiHelper.ScalePadding(2),
                        Cursor = Cursors.Hand,
                        Tag = filterKey,
                        AccessibleName = StripEmojiPrefix(text)
                    };
                    btn.FlatAppearance.BorderColor = ModernTheme.Border;
                    btn.FlatAppearance.BorderSize = 1;
                    btn.FlatAppearance.MouseOverBackColor = ModernTheme.Selection;

                    btn.Click += (sender, args) => displayAppxWithFilter(filterKey);

                    appxFilterButtons.Add(btn);
                    appxFilterPanel.Controls.Add(btn);
                }

                // Show the filter panel
                appxFilterPanel.Visible = true;
            };

            // Storage filter state (for Disk/Volume filters)
            var storageFilterButtons = new List<Button>();
            themeData.StorageFilterButtons = storageFilterButtons;
            string currentStorageFilter = "Disk";

            // Function to display storage filters with filter type
            Action<string> displayStorageWithFilter = null!;
            displayStorageWithFilter = (filter) =>
            {
                currentStorageFilter = filter;
                contentGrid.Columns.Clear();
                contentGrid.Rows.Clear();

                contentGrid.Columns.Add("name", "Property");
                contentGrid.Columns.Add("value", "Value");
                contentGrid.Columns["name"].FillWeight = 30;
                contentGrid.Columns["value"].FillWeight = 70;

                var items = filter == "Disk" 
                    ? _infoExtractor.GetDiskFilters() 
                    : _infoExtractor.GetVolumeFilters();

                foreach (var item in items)
                {
                    var rowIndex = contentGrid.Rows.Add(item.Name, item.Value);
                    contentGrid.Rows[rowIndex].Tag = item;
                }

                // Update filter button states
                foreach (var btn in storageFilterButtons)
                {
                    var btnFilter = btn.Tag?.ToString() ?? "";
                    btn.BackColor = btnFilter == filter ? ModernTheme.Accent : ModernTheme.Surface;
                    btn.ForeColor = btnFilter == filter ? Color.White : ModernTheme.TextPrimary;
                }
            };

            // Function to create storage filter buttons
            Action createStorageFilterButtons = null!;
            createStorageFilterButtons = () =>
            {
                appxFilterPanel.Controls.Clear();
                storageFilterButtons.Clear();

                var filters = new[]
                {
                    ("💾 Disk Filters", "Disk"),
                    ("📀 Volume Filters", "Volume")
                };

                foreach (var (text, filterKey) in filters)
                {
                    var btn = new Button
                    {
                        Text = text,
                        FlatStyle = FlatStyle.Flat,
                        BackColor = filterKey == "Disk" ? ModernTheme.Accent : ModernTheme.Surface,
                        ForeColor = filterKey == "Disk" ? Color.White : ModernTheme.TextPrimary,
                        Font = ModernTheme.RegularFont,
                        Height = DpiHelper.Scale(28),
                        AutoSize = true,
                        Padding = DpiHelper.ScalePadding(8, 0, 8, 0),
                        Margin = DpiHelper.ScalePadding(2),
                        Cursor = Cursors.Hand,
                        Tag = filterKey,
                        AccessibleName = StripEmojiPrefix(text)
                    };
                    btn.FlatAppearance.BorderColor = ModernTheme.Border;
                    btn.FlatAppearance.BorderSize = 1;
                    btn.FlatAppearance.MouseOverBackColor = ModernTheme.Selection;

                    btn.Click += (sender, args) => displayStorageWithFilter(filterKey);

                    storageFilterButtons.Add(btn);
                    appxFilterPanel.Controls.Add(btn);
                }

                // Show the filter panel
                appxFilterPanel.Visible = true;
            };

            // CBS Packages display state
            List<AnalysisSection>? cbsPackagesSections = null;
            string currentCbsSubView = "AllPackages"; // Track current CBS sub-view
            List<(string group, string package, string state, string installed, string user, int visibility, AnalysisItem item)> allPackagesData = new();
            TextBox? cbsSearchBox = null;
            Label? cbsSearchLabel = null;
            CheckBox? cbsDismCheckBox = null;

            // Function to display CBS sub-buttons and content
            Action<string> displayCbsSubView = null!;
            
            // Function to filter All Packages view (declared early, assigned later)
            Action<string> filterAllPackages = null!;
            
            // Function to create CBS sub-buttons
            Action createCbsSubButtons = null!;
            createCbsSubButtons = () =>
            {
                appxFilterPanel.Controls.Clear();
                appxFilterPanel.Visible = true;

                var subButtons = new[]
                {
                    ("AllPackages", "All Packages"),
                    ("PendingSessions", "Pending Sessions"),
                    ("PendingPackages", "Pending Packages"),
                    ("RebootStatus", "Reboot Status")
                };

                foreach (var (viewKey, viewLabel) in subButtons)
                {
                    var btn = new Button
                    {
                        Text = viewLabel,
                        FlatStyle = FlatStyle.Flat,
                        BackColor = viewKey == currentCbsSubView ? ModernTheme.Accent : ModernTheme.Surface,
                        ForeColor = viewKey == currentCbsSubView ? Color.White : ModernTheme.TextPrimary,
                        Font = ModernTheme.RegularFont,
                        Height = DpiHelper.Scale(26),
                        AutoSize = true,
                        Padding = DpiHelper.ScalePadding(8, 0, 8, 0),
                        Margin = DpiHelper.ScalePadding(2),
                        Cursor = Cursors.Hand,
                        Tag = viewKey,
                        AccessibleName = viewLabel
                    };
                    btn.FlatAppearance.BorderColor = ModernTheme.Border;
                    btn.FlatAppearance.BorderSize = 1;
                    btn.FlatAppearance.MouseOverBackColor = ModernTheme.Selection;

                    var capturedKey = viewKey;
                    btn.Click += (s, e) => displayCbsSubView(capturedKey);

                    appxFilterPanel.Controls.Add(btn);
                }

                // Add search label and textbox for All Packages view
                cbsSearchLabel = new Label
                {
                    Text = "Search:",
                    AutoSize = true,
                    ForeColor = ModernTheme.TextPrimary,
                    Font = ModernTheme.RegularFont,
                    Margin = DpiHelper.ScalePadding(15, 6, 5, 0),
                    Visible = currentCbsSubView == "AllPackages"
                };
                appxFilterPanel.Controls.Add(cbsSearchLabel);

                cbsSearchBox = new TextBox
                {
                    Width = DpiHelper.Scale(200),
                    Height = DpiHelper.Scale(26),
                    Font = ModernTheme.RegularFont,
                    BackColor = ModernTheme.Surface,
                    ForeColor = ModernTheme.TextPrimary,
                    BorderStyle = BorderStyle.FixedSingle,
                    Margin = DpiHelper.ScalePadding(2),
                    Visible = currentCbsSubView == "AllPackages"
                };
                ModernTheme.ApplyTo(cbsSearchBox);
                cbsSearchBox.TextChanged += (s, e) => filterAllPackages(cbsSearchBox.Text);
                appxFilterPanel.Controls.Add(cbsSearchBox);

                // DISM /Get-Package filter checkbox (only visible on All Packages view)
                cbsDismCheckBox = new CheckBox
                {
                    Text = "DISM /Get-Package",
                    AutoSize = true,
                    ForeColor = ModernTheme.TextPrimary,
                    Font = ModernTheme.RegularFont,
                    Margin = DpiHelper.ScalePadding(15, 5, 5, 0),
                    Visible = currentCbsSubView == "AllPackages"
                };
                cbsDismCheckBox.CheckedChanged += (s, e) => filterAllPackages(cbsSearchBox?.Text ?? "");
                appxFilterPanel.Controls.Add(cbsDismCheckBox);
            };

            // Function to update CBS sub-button states
            Action updateCbsSubButtonStates = () =>
            {
                foreach (Control ctrl in appxFilterPanel.Controls)
                {
                    if (ctrl is Button btn && btn.Tag is string viewKey)
                    {
                        btn.BackColor = viewKey == currentCbsSubView ? ModernTheme.Accent : ModernTheme.Surface;
                        btn.ForeColor = viewKey == currentCbsSubView ? Color.White : ModernTheme.TextPrimary;
                    }
                }
                // Update search box visibility based on current view
                if (cbsSearchBox != null) cbsSearchBox.Visible = currentCbsSubView == "AllPackages";
                if (cbsSearchLabel != null) cbsSearchLabel.Visible = currentCbsSubView == "AllPackages";
                if (cbsDismCheckBox != null) cbsDismCheckBox.Visible = currentCbsSubView == "AllPackages";
            };

            // Assign function to filter All Packages view
            filterAllPackages = (searchText) =>
            {
                if (currentCbsSubView != "AllPackages") return;
                
                contentGrid.Rows.Clear();
                var search = searchText.Trim().ToLowerInvariant();
                bool dismFilter = cbsDismCheckBox?.Checked == true;
                
                foreach (var (group, package, state, installed, user, visibility, item) in allPackagesData)
                {
                    // Apply DISM /Get-Package filter: only show packages with Visibility == 1
                    if (dismFilter && visibility != 1)
                        continue;

                    if (string.IsNullOrEmpty(search) || 
                        group.ToLowerInvariant().Contains(search) || 
                        package.ToLowerInvariant().Contains(search))
                    {
                        var rowIndex = contentGrid.Rows.Add(group, package, state, installed, user);
                        contentGrid.Rows[rowIndex].Tag = item;
                    }
                }

                // Re-select the first row to trigger detail pane update (Tag is now set)
                if (contentGrid.Rows.Count > 0)
                {
                    contentGrid.ClearSelection();
                    contentGrid.Rows[0].Selected = true;
                }
            };

            // Function to display All Packages view
            Action displayAllPackages = () =>
            {
                contentGrid.Columns.Clear();
                contentGrid.Rows.Clear();
                allPackagesData.Clear();

                // Clear search box when refreshing
                if (cbsSearchBox != null) cbsSearchBox.Text = "";
                if (cbsDismCheckBox != null) cbsDismCheckBox.Checked = false;

                // Load packages data if not cached
                if (cbsPackagesSections == null)
                {
                    cbsPackagesSections = _infoExtractor!.GetPackagesAnalysis();
                }

                if (cbsPackagesSections.Count == 0)
                {
                    contentGrid.Columns.Add("info", "Information");
                    contentGrid.Rows.Add("No packages found");
                    return;
                }

                // Show all packages directly (skip summary section)
                contentGrid.Columns.Add("group", "Package Group");
                contentGrid.Columns.Add("package", "Package Version");
                contentGrid.Columns.Add("state", "State");
                contentGrid.Columns.Add("installed", "Installed");
                contentGrid.Columns.Add("user", "User");
                contentGrid.Columns["group"].FillWeight = 25;
                contentGrid.Columns["package"].FillWeight = 30;
                contentGrid.Columns["state"].FillWeight = 15;
                contentGrid.Columns["installed"].FillWeight = 18;
                contentGrid.Columns["user"].FillWeight = 12;

                foreach (var section in cbsPackagesSections.Where(s => !s.Title.Contains("Summary")))
                {
                    // Extract group name from title (remove count and emoji)
                    var groupName = section.Title;
                    if (groupName.StartsWith("📦 "))
                        groupName = groupName.Substring(3);
                    var parenIndex = groupName.LastIndexOf(" (");
                    if (parenIndex > 0)
                        groupName = groupName.Substring(0, parenIndex);

                    foreach (var item in section.Items)
                    {
                        // Parse value to extract parts: "State: X | Installed: Y | User: Z"
                        var valueParts = item.Value?.Split('|') ?? Array.Empty<string>();
                        var state = "";
                        var installed = "";
                        var user = "";

                        int visibility = 0;
                        foreach (var part in valueParts)
                        {
                            var trimmed = part.Trim();
                            if (trimmed.StartsWith("State:"))
                                state = trimmed.Substring(6).Trim();
                            else if (trimmed.StartsWith("Installed:"))
                                installed = trimmed.Substring(10).Trim();
                            else if (trimmed.StartsWith("User:"))
                                user = trimmed.Substring(5).Trim();
                            else if (trimmed.StartsWith("Visibility:"))
                                int.TryParse(trimmed.Substring(11).Trim(), out visibility);
                        }

                        // Store in allPackagesData for filtering
                        allPackagesData.Add((groupName, item.Name, state, installed, user, visibility, item));

                        var rowIndex = contentGrid.Rows.Add(groupName, item.Name, state, installed, user);
                        contentGrid.Rows[rowIndex].Tag = item;
                    }
                }

                // Re-select the first row to trigger detail pane update (Tag is now set)
                if (contentGrid.Rows.Count > 0)
                {
                    contentGrid.ClearSelection();
                    contentGrid.Rows[0].Selected = true;
                }
            };

            // Function to display Pending Sessions view
            Action displayPendingSessions = () =>
            {
                contentGrid.Columns.Clear();
                contentGrid.Rows.Clear();

                var sessions = _infoExtractor!.GetCbsPendingSessionsAnalysis();

                if (sessions.Count == 0 || sessions.All(s => s.Items.Count == 0))
                {
                    contentGrid.Columns.Add("info", "Information");
                    contentGrid.Rows.Add("No pending sessions found");
                    return;
                }

                contentGrid.Columns.Add("session", "Session");
                contentGrid.Columns.Add("status", "Status");
                contentGrid.Columns.Add("details", "Details");
                contentGrid.Columns["session"].FillWeight = 30;
                contentGrid.Columns["status"].FillWeight = 20;
                contentGrid.Columns["details"].FillWeight = 50;

                foreach (var section in sessions)
                {
                    foreach (var item in section.Items)
                    {
                        var rowIndex = contentGrid.Rows.Add(item.Name, item.Value, section.Title);
                        contentGrid.Rows[rowIndex].Tag = item;
                    }
                }

                // Re-select the first row to trigger detail pane update (Tag is now set)
                if (contentGrid.Rows.Count > 0)
                {
                    contentGrid.ClearSelection();
                    contentGrid.Rows[0].Selected = true;
                }
            };

            // Function to display Pending Packages view
            Action displayPendingPackages = () =>
            {
                contentGrid.Columns.Clear();
                contentGrid.Rows.Clear();

                var packages = _infoExtractor!.GetCbsPendingPackagesAnalysis();

                if (packages.Count == 0 || packages.All(s => s.Items.Count == 0))
                {
                    contentGrid.Columns.Add("info", "Information");
                    contentGrid.Rows.Add("No pending packages found");
                    return;
                }

                contentGrid.Columns.Add("package", "Package Name");
                contentGrid.Columns.Add("status", "Status");
                contentGrid.Columns.Add("details", "Details");
                contentGrid.Columns["package"].FillWeight = 40;
                contentGrid.Columns["status"].FillWeight = 20;
                contentGrid.Columns["details"].FillWeight = 40;

                foreach (var section in packages)
                {
                    foreach (var item in section.Items)
                    {
                        var rowIndex = contentGrid.Rows.Add(item.Name, item.Value, section.Title);
                        contentGrid.Rows[rowIndex].Tag = item;
                    }
                }

                // Re-select the first row to trigger detail pane update (Tag is now set)
                if (contentGrid.Rows.Count > 0)
                {
                    contentGrid.ClearSelection();
                    contentGrid.Rows[0].Selected = true;
                }
            };

            // Function to display Reboot Status view
            Action displayRebootStatus = () =>
            {
                contentGrid.Columns.Clear();
                contentGrid.Rows.Clear();

                var status = _infoExtractor!.GetCbsRebootStatusAnalysis();

                if (status.Count == 0 || status.All(s => s.Items.Count == 0))
                {
                    contentGrid.Columns.Add("info", "Information");
                    contentGrid.Rows.Add("No reboot status information found");
                    return;
                }

                contentGrid.Columns.Add("property", "Property");
                contentGrid.Columns.Add("value", "Value");
                contentGrid.Columns.Add("details", "Details");
                contentGrid.Columns["property"].FillWeight = 25;
                contentGrid.Columns["value"].FillWeight = 20;
                contentGrid.Columns["details"].FillWeight = 55;

                foreach (var section in status)
                {
                    foreach (var item in section.Items)
                    {
                        // Extract a meaningful summary for the Details column
                        // Use the first line of RegistryValue or a cleaned up version
                        var details = item.RegistryValue ?? "";
                        
                        // For multi-line values, get a summary
                        if (details.Contains('\n'))
                        {
                            // For items with "value = X" followed by description, extract the description
                            var lines = details.Split('\n');
                            if (lines.Length > 1)
                            {
                                // Skip lines that just show the raw value, get the descriptive part
                                details = string.Join(" ", lines.Where(l => 
                                    !l.Trim().StartsWith("ServicingInProgress =") &&
                                    !l.Trim().StartsWith("RebootPending =") &&
                                    !l.Trim().StartsWith("RebootInProgress =") &&
                                    !l.Trim().StartsWith("SessionsPendingExclusive =") &&
                                    !l.Trim().StartsWith("LastTrustTime =") &&
                                    !string.IsNullOrWhiteSpace(l)
                                ).Take(2)).Trim();
                            }
                        }
                        
                        // Truncate if too long for display
                        if (details.Length > 150)
                            details = details.Substring(0, 147) + "...";

                        var rowIndex = contentGrid.Rows.Add(item.Name, item.Value, details);
                        contentGrid.Rows[rowIndex].Tag = item;
                    }
                }

                // Re-select the first row to trigger detail pane update (Tag is now set)
                if (contentGrid.Rows.Count > 0)
                {
                    contentGrid.ClearSelection();
                    contentGrid.Rows[0].Selected = true;
                }
            };

            // Main function to display CBS sub-view
            displayCbsSubView = (viewKey) =>
            {
                currentCbsSubView = viewKey;
                updateCbsSubButtonStates();

                contentGrid.Visible = true;
                networkSplitContainer.Visible = false;
                firewallPanel.Visible = false;

                switch (viewKey)
                {
                    case "AllPackages":
                        displayAllPackages();
                        break;
                    case "PendingSessions":
                        displayPendingSessions();
                        break;
                    case "PendingPackages":
                        displayPendingPackages();
                        break;
                    case "RebootStatus":
                        displayRebootStatus();
                        break;
                }
            };

            // Function to display CBS packages sections (entry point - shows sub-buttons)
            Action displayCbsPackages = null!;
            displayCbsPackages = () =>
            {
                contentGrid.Columns.Clear();
                contentGrid.Rows.Clear();
                networkSplitContainer.Visible = false;
                firewallPanel.Visible = false;
                contentGrid.Visible = true;

                // Create CBS sub-buttons
                createCbsSubButtons();

                // Display default view (All Packages)
                currentCbsSubView = "AllPackages";
                displayAllPackages();
            };

            // User Profiles display state
            AnalysisSection? cachedProfilesSection = null;
            string currentProfileFilter = "All";

            // Function to display User Profiles with filter
            Action<string> displayProfilesWithFilter = null!;
            displayProfilesWithFilter = (filter) =>
            {
                currentProfileFilter = filter;
                contentGrid.Columns.Clear();
                contentGrid.Rows.Clear();

                contentGrid.Columns.Add("name", "Name");
                contentGrid.Columns.Add("sid", "SID");
                contentGrid.Columns.Add("path", "Path");
                contentGrid.Columns.Add("lastLogon", "Last Logon");
                contentGrid.Columns.Add("lastLogoff", "Last Logoff");
                contentGrid.Columns[0].FillWeight = 20;
                contentGrid.Columns[1].FillWeight = 25;
                contentGrid.Columns[2].FillWeight = 25;
                contentGrid.Columns[3].FillWeight = 15;
                contentGrid.Columns[4].FillWeight = 15;

                if (cachedProfilesSection == null) return;

                foreach (var item in cachedProfilesSection.Items)
                {
                    if (!item.IsSubSection || item.SubItems == null || item.SubItems.Count == 0)
                        continue;

                    var sid = item.SubItems[0].Value ?? "";

                    // Filter: "Temp" shows only SIDs ending with .bak
                    if (filter == "Temp" && !sid.EndsWith(".bak", StringComparison.OrdinalIgnoreCase))
                        continue;

                    var path = item.SubItems.Count > 1 ? item.SubItems[1].Value : "";
                    var lastLogon = item.SubItems.Count > 2 ? item.SubItems[2].Value : "";
                    var lastLogoff = item.SubItems.Count > 3 ? item.SubItems[3].Value : "";
                    var rowIdx = contentGrid.Rows.Add(item.Name, sid, path, lastLogon, lastLogoff);
                    contentGrid.Rows[rowIdx].Tag = item;
                }

                // Update filter button states
                foreach (Control ctrl in appxFilterPanel.Controls)
                {
                    if (ctrl is Button btn && btn.Tag is string btnFilter)
                    {
                        btn.BackColor = btnFilter == filter ? ModernTheme.Accent : ModernTheme.Surface;
                        btn.ForeColor = btnFilter == filter ? Color.White : ModernTheme.TextPrimary;
                    }
                }
            };

            // Function to create User Profiles filter buttons
            Action createProfileFilterButtons = null!;
            createProfileFilterButtons = () =>
            {
                appxFilterPanel.Controls.Clear();
                appxFilterPanel.Visible = true;

                var filters = new[]
                {
                    ("All", "All Profiles"),
                    ("Temp", "Temp Profiles")
                };

                foreach (var (filterKey, label) in filters)
                {
                    var btn = new Button
                    {
                        Text = label,
                        FlatStyle = FlatStyle.Flat,
                        BackColor = filterKey == currentProfileFilter ? ModernTheme.Accent : ModernTheme.Surface,
                        ForeColor = filterKey == currentProfileFilter ? Color.White : ModernTheme.TextPrimary,
                        Font = ModernTheme.RegularFont,
                        Height = DpiHelper.Scale(26),
                        AutoSize = true,
                        Padding = DpiHelper.ScalePadding(8, 0, 8, 0),
                        Margin = DpiHelper.ScalePadding(2),
                        Cursor = Cursors.Hand,
                        Tag = filterKey,
                        AccessibleName = label
                    };
                    btn.FlatAppearance.BorderColor = ModernTheme.Border;
                    btn.FlatAppearance.BorderSize = 1;
                    btn.FlatAppearance.MouseOverBackColor = ModernTheme.Selection;

                    var capturedKey = filterKey;
                    btn.Click += (s, e) => displayProfilesWithFilter(capturedKey);

                    appxFilterPanel.Controls.Add(btn);
                }
            };

            // Guest Agent sub-view state (for Extensions sub-tab)
            string currentGuestAgentSubView = ""; // Empty = overview shown, "Extensions" = extensions shown
            
            // Forward declarations for Guest Agent display functions
            Action createGuestAgentSubButtons = null!;
            Action<string> displayGuestAgentSubView = null!;
            Action updateGuestAgentSubButtonStates = null!;
            
            // Function to create Guest Agent sub-buttons (Extensions tab)
            createGuestAgentSubButtons = () =>
            {
                var isSoftware = _parser.CurrentHiveType == OfflineRegistryParser.HiveType.SOFTWARE;
                
                appxFilterPanel.Controls.Clear();
                appxFilterPanel.Visible = true;
                
                var extensionsBtn = new Button
                {
                    Text = "Extensions",
                    FlatStyle = FlatStyle.Flat,
                    BackColor = isSoftware ? 
                        (currentGuestAgentSubView == "Extensions" ? ModernTheme.Accent : ModernTheme.Surface) : 
                        ModernTheme.TreeViewBack,
                    ForeColor = isSoftware ? 
                        (currentGuestAgentSubView == "Extensions" ? Color.White : ModernTheme.TextPrimary) : 
                        ModernTheme.TextDisabled,
                    Font = ModernTheme.RegularFont,
                    Height = DpiHelper.Scale(26),
                    AutoSize = true,
                    Padding = DpiHelper.ScalePadding(8, 0, 8, 0),
                    Margin = DpiHelper.ScalePadding(2),
                    Cursor = isSoftware ? Cursors.Hand : Cursors.Default,
                    Tag = "Extensions",
                    AccessibleName = "Extensions"
                };
                extensionsBtn.FlatAppearance.BorderColor = isSoftware ? ModernTheme.Border : ModernTheme.TextDisabled;
                extensionsBtn.FlatAppearance.BorderSize = 1;
                extensionsBtn.FlatAppearance.MouseOverBackColor = isSoftware ? ModernTheme.Selection : ModernTheme.TreeViewBack;
                
                // Use shared tooltip from themeData
                themeData.SubCategoryToolTip ??= new ToolTip { InitialDelay = 200, ReshowDelay = 100, AutoPopDelay = 5000 };
                if (!isSoftware)
                    themeData.SubCategoryToolTip.SetToolTip(extensionsBtn, "This feature requires SOFTWARE hive to be loaded");
                else
                    themeData.SubCategoryToolTip.SetToolTip(extensionsBtn, "View Azure VM Extensions installed on this system");
                
                var capturedIsSoftware = isSoftware; // Capture for closure
                extensionsBtn.Click += (s, e) =>
                {
                    if (capturedIsSoftware)
                        displayGuestAgentSubView("Extensions");
                    else
                        MessageBox.Show(
                            "This feature requires SOFTWARE hive to be loaded.\n\n" +
                            "Azure VM Extensions are stored in the SOFTWARE hive.\n\n" +
                            "Common location: C:\\Windows\\System32\\config\\SOFTWARE",
                            "SOFTWARE Hive Required",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                };
                
                appxFilterPanel.Controls.Add(extensionsBtn);
            };
            
            // Function to update Guest Agent sub-button states
            updateGuestAgentSubButtonStates = () =>
            {
                var isSoftware = _parser.CurrentHiveType == OfflineRegistryParser.HiveType.SOFTWARE;
                
                foreach (Control ctrl in appxFilterPanel.Controls)
                {
                    if (ctrl is Button btn && btn.Tag is string viewKey)
                    {
                        if (isSoftware)
                        {
                            btn.BackColor = viewKey == currentGuestAgentSubView ? ModernTheme.Accent : ModernTheme.Surface;
                            btn.ForeColor = viewKey == currentGuestAgentSubView ? Color.White : ModernTheme.TextPrimary;
                        }
                    }
                }
            };
            
            // Function to display Guest Agent sub-view (Extensions)
            displayGuestAgentSubView = (viewKey) =>
            {
                currentGuestAgentSubView = viewKey;
                updateGuestAgentSubButtonStates();
                
                if (viewKey == "Extensions")
                {
                    // Display Extensions in the grid
                    var extensionsSection = _infoExtractor.GetAzureExtensionsAnalysis();
                    
                    currentSection = extensionsSection;
                    contentGrid.Columns.Clear();
                    contentGrid.Rows.Clear();
                    contentGrid.Visible = true;
                    networkSplitContainer.Visible = false;
                    firewallPanel.Visible = false;
                    
                    // Keep appxFilterPanel visible for sub-tabs
                    // appxFilterPanel.Visible = true; // Already visible from createGuestAgentSubButtons
                    
                    // Setup grid columns for Extensions
                    contentGrid.Columns.Add("name", "Extension");
                    contentGrid.Columns.Add("value", "Status");
                    contentGrid.Columns["name"]!.Width = 350;
                    contentGrid.Columns["value"]!.Width = 150;
                    
                    // Populate grid with extensions
                    foreach (var item in extensionsSection.Items)
                    {
                        var rowIndex = contentGrid.Rows.Add(item.Name, item.Value);
                        contentGrid.Rows[rowIndex].Tag = item;
                    }
                }
            };

            // CBS Components display state (for COMPONENTS hive)
            List<AnalysisItem>? cbsComponentsFullList = null; // Full list for search filtering
            TextBox? cbsComponentsSearchBox = null;
            Label? cbsComponentsSearchLabel = null;
            Panel? cbsComponentsSearchPanel = null;
            
            // CBS Identities display state
            List<AnalysisItem>? cbsIdentitiesFullList = null;
            Panel? cbsIdentitiesInfoPanel = null;
            
            // Components overview panel (landing page for COMPONENTS hive)
            Panel? componentsOverviewPanel = null;

            // Forward declarations for CBS display functions (defined later, but referenced in displayComponentsOverview)
            Action displayCbsComponents = null!;
            Action displayCbsIdentities = null!;

            // Function to display Components overview (landing page with sub-options)
            Action displayComponentsOverview = null!;
            displayComponentsOverview = () =>
            {
                contentGrid.Columns.Clear();
                contentGrid.Rows.Clear();
                contentGrid.Visible = false;
                networkSplitContainer.Visible = false;
                firewallPanel.Visible = false;
                appxFilterPanel.Visible = false;
                currentSection = null;
                
                // Hide CBS search/info panels
                if (cbsComponentsSearchPanel != null) cbsComponentsSearchPanel.Visible = false;
                if (cbsIdentitiesInfoPanel != null) cbsIdentitiesInfoPanel.Visible = false;
                
                // Create overview panel if not exists
                if (componentsOverviewPanel == null)
                {
                    componentsOverviewPanel = new Panel
                    {
                        Dock = DockStyle.Fill,
                        BackColor = ModernTheme.Background,
                        Padding = DpiHelper.ScalePadding(40)
                    };
                    
                    var titleLabel = new Label
                    {
                        Text = "Component Store Analysis",
                        Font = new Font("Segoe UI", 18F, FontStyle.Bold),
                        ForeColor = ModernTheme.TextPrimary,
                        AutoSize = true,
                        Location = DpiHelper.ScalePoint(40, 30)
                    };
                    componentsOverviewPanel.Controls.Add(titleLabel);
                    
                    var descLabel = new Label
                    {
                        Text = "Select an analysis option below:",
                        Font = ModernTheme.RegularFont,
                        ForeColor = ModernTheme.TextSecondary,
                        AutoSize = true,
                        Location = DpiHelper.ScalePoint(40, 70)
                    };
                    componentsOverviewPanel.Controls.Add(descLabel);
                    
                    // CBS Packages button card
                    var packagesCard = new Panel
                    {
                        Size = DpiHelper.ScaleSize(350, 100),
                        Location = DpiHelper.ScalePoint(40, 110),
                        BackColor = ModernTheme.Surface,
                        Cursor = Cursors.Hand
                    };
                    packagesCard.Paint += (s, ev) =>
                    {
                        using var pen = new Pen(ModernTheme.Border);
                        ev.Graphics.DrawRectangle(pen, 0, 0, packagesCard.Width - 1, packagesCard.Height - 1);
                    };
                    
                    var packagesIcon = new Label
                    {
                        Text = "\uE8F1", // Package icon
                        Font = new Font("Segoe MDL2 Assets", 24F),
                        ForeColor = ModernTheme.Accent,
                        AutoSize = true,
                        Location = DpiHelper.ScalePoint(15, 30)
                    };
                    packagesCard.Controls.Add(packagesIcon);
                    
                    var packagesTitle = new Label
                    {
                        Text = "CBS Packages",
                        Font = new Font("Segoe UI Semibold", 12F),
                        ForeColor = ModernTheme.TextPrimary,
                        AutoSize = true,
                        Location = DpiHelper.ScalePoint(60, 20)
                    };
                    packagesCard.Controls.Add(packagesTitle);
                    
                    var packagesDesc = new Label
                    {
                        Text = "View all components in DerivedData\\Components.\nBrowse and search through installed packages.",
                        Font = ModernTheme.RegularFont,
                        ForeColor = ModernTheme.TextSecondary,
                        AutoSize = true,
                        Location = DpiHelper.ScalePoint(60, 45)
                    };
                    packagesCard.Controls.Add(packagesDesc);
                    
                    packagesCard.Click += (s, e) => displayCbsComponents();
                    foreach (Control c in packagesCard.Controls)
                        c.Click += (s, e) => displayCbsComponents();
                    
                    packagesCard.MouseEnter += (s, e) => packagesCard.BackColor = ModernTheme.Selection;
                    packagesCard.MouseLeave += (s, e) => packagesCard.BackColor = ModernTheme.Surface;
                    foreach (Control c in packagesCard.Controls)
                    {
                        c.MouseEnter += (s, e) => packagesCard.BackColor = ModernTheme.Selection;
                        c.MouseLeave += (s, e) => packagesCard.BackColor = ModernTheme.Surface;
                    }
                    
                    componentsOverviewPanel.Controls.Add(packagesCard);
                    
                    // CBS Identities button card
                    var identitiesCard = new Panel
                    {
                        Size = DpiHelper.ScaleSize(350, 100),
                        Location = DpiHelper.ScalePoint(40, 220),
                        BackColor = ModernTheme.Surface,
                        Cursor = Cursors.Hand
                    };
                    identitiesCard.Paint += (s, ev) =>
                    {
                        using var pen = new Pen(ModernTheme.Border);
                        ev.Graphics.DrawRectangle(pen, 0, 0, identitiesCard.Width - 1, identitiesCard.Height - 1);
                    };
                    
                    var identitiesIcon = new Label
                    {
                        Text = "\uE9D9", // Scan/search icon
                        Font = new Font("Segoe MDL2 Assets", 24F),
                        ForeColor = ModernTheme.Accent,
                        AutoSize = true,
                        Location = DpiHelper.ScalePoint(15, 30)
                    };
                    identitiesCard.Controls.Add(identitiesIcon);
                    
                    var identitiesTitle = new Label
                    {
                        Text = "CBS Identities Scan",
                        Font = new Font("Segoe UI Semibold", 12F),
                        ForeColor = ModernTheme.TextPrimary,
                        AutoSize = true,
                        Location = DpiHelper.ScalePoint(60, 20)
                    };
                    identitiesCard.Controls.Add(identitiesTitle);
                    
                    var identitiesDesc = new Label
                    {
                        Text = "Scan component identities for invalid characters.\nDetects non-ASCII bytes that may cause issues.",
                        Font = ModernTheme.RegularFont,
                        ForeColor = ModernTheme.TextSecondary,
                        AutoSize = true,
                        Location = DpiHelper.ScalePoint(60, 45)
                    };
                    identitiesCard.Controls.Add(identitiesDesc);
                    
                    identitiesCard.Click += (s, e) => displayCbsIdentities();
                    foreach (Control c in identitiesCard.Controls)
                        c.Click += (s, e) => displayCbsIdentities();
                    
                    identitiesCard.MouseEnter += (s, e) => identitiesCard.BackColor = ModernTheme.Selection;
                    identitiesCard.MouseLeave += (s, e) => identitiesCard.BackColor = ModernTheme.Surface;
                    foreach (Control c in identitiesCard.Controls)
                    {
                        c.MouseEnter += (s, e) => identitiesCard.BackColor = ModernTheme.Selection;
                        c.MouseLeave += (s, e) => identitiesCard.BackColor = ModernTheme.Surface;
                    }
                    
                    componentsOverviewPanel.Controls.Add(identitiesCard);
                    
                    rightPanel.Controls.Add(componentsOverviewPanel);
                }
                
                componentsOverviewPanel.Visible = true;
                componentsOverviewPanel.BringToFront();
            };

            // Function to filter CBS Components
            Action<string> filterCbsComponents = null!;
            filterCbsComponents = (searchText) =>
            {
                if (cbsComponentsFullList == null) return;
                
                contentGrid.Rows.Clear();
                var search = searchText.Trim().ToLowerInvariant();
                
                foreach (var item in cbsComponentsFullList)
                {
                    if (string.IsNullOrEmpty(search) || 
                        item.Name.ToLowerInvariant().Contains(search))
                    {
                        var rowIndex = contentGrid.Rows.Add(item.Name);
                        contentGrid.Rows[rowIndex].Tag = item;
                    }
                }
            };

            // Function to display CBS Components (from COMPONENTS hive)
            displayCbsComponents = () =>
            {
                contentGrid.Columns.Clear();
                contentGrid.Rows.Clear();
                networkSplitContainer.Visible = false;
                firewallPanel.Visible = false;
                contentGrid.Visible = true;
                appxFilterPanel.Visible = false;
                currentSection = null; // Clear current section since this is a custom view
                
                // Hide overview panel
                if (componentsOverviewPanel != null) componentsOverviewPanel.Visible = false;
                // Hide CBS Identities info panel
                if (cbsIdentitiesInfoPanel != null) cbsIdentitiesInfoPanel.Visible = false;

                // Load components data if not cached
                if (cbsComponentsFullList == null)
                {
                    cbsComponentsFullList = _infoExtractor!.GetAllComponentsList();
                }

                if (cbsComponentsFullList.Count == 0)
                {
                    contentGrid.Columns.Add("info", "Information");
                    contentGrid.Rows.Add("No component data found in DerivedData\\Components");
                    return;
                }

                // Create search panel if not exists
                if (cbsComponentsSearchPanel == null)
                {
                    cbsComponentsSearchPanel = new FlowLayoutPanel
                    {
                        Dock = DockStyle.Top,
                        Height = DpiHelper.Scale(35),
                        BackColor = ModernTheme.TreeViewBack,
                        Padding = DpiHelper.ScalePadding(5),
                        AutoSize = false,
                        FlowDirection = FlowDirection.LeftToRight
                    };

                    cbsComponentsSearchLabel = new Label
                    {
                        Text = "Search:",
                        AutoSize = true,
                        ForeColor = ModernTheme.TextPrimary,
                        Font = ModernTheme.RegularFont,
                        Margin = DpiHelper.ScalePadding(5, 6, 5, 0)
                    };
                    cbsComponentsSearchPanel.Controls.Add(cbsComponentsSearchLabel);

                    cbsComponentsSearchBox = new TextBox
                    {
                        Width = DpiHelper.Scale(300),
                        Height = DpiHelper.Scale(26),
                        Font = ModernTheme.RegularFont,
                        BackColor = ModernTheme.Surface,
                        ForeColor = ModernTheme.TextPrimary,
                        BorderStyle = BorderStyle.FixedSingle,
                        Margin = DpiHelper.ScalePadding(2)
                    };
                    ModernTheme.ApplyTo(cbsComponentsSearchBox);
                    cbsComponentsSearchBox.TextChanged += (s, e) => filterCbsComponents(cbsComponentsSearchBox.Text);
                    cbsComponentsSearchPanel.Controls.Add(cbsComponentsSearchBox);

                    // Add total count label
                    var totalLabel = new Label
                    {
                        Text = $"Total: {cbsComponentsFullList.Count:N0} components",
                        AutoSize = true,
                        ForeColor = ModernTheme.TextSecondary,
                        Font = ModernTheme.RegularFont,
                        Margin = DpiHelper.ScalePadding(15, 6, 5, 0)
                    };
                    cbsComponentsSearchPanel.Controls.Add(totalLabel);

                    // Insert search panel above content grid (in contentDetailSplit.Panel1)
                    var contentPanel = contentDetailSplit?.Panel1 ?? rightPanel;
                    contentPanel.Controls.Add(cbsComponentsSearchPanel);
                }

                // Show search panel and clear search box
                cbsComponentsSearchPanel.Visible = true;
                // Ensure proper docking order: contentGrid (Fill) must be at index 0 (processed first)
                // Search panel (Top) must be after contentGrid so it docks above it
                var containerPanel = contentDetailSplit?.Panel1 ?? rightPanel;
                containerPanel.Controls.SetChildIndex(contentGrid, 0);
                containerPanel.Controls.SetChildIndex(cbsComponentsSearchPanel, 1);
                if (cbsComponentsSearchBox != null) cbsComponentsSearchBox.Text = "";

                // Show all components in a single column grid
                contentGrid.Columns.Add("component", "Component Name");
                contentGrid.Columns["component"].FillWeight = 100;

                foreach (var item in cbsComponentsFullList)
                {
                    var rowIndex = contentGrid.Rows.Add(item.Name);
                    contentGrid.Rows[rowIndex].Tag = item;
                }

                // Update subcategory button states - Components button stays highlighted
                foreach (var btn in subCategoryButtons)
                {
                    if (btn.Tag is string s && s == "Components")
                    {
                        btn.BackColor = ModernTheme.Accent;
                        btn.Invalidate();
                    }
                }
            };

            // Function to display CBS Identities scan results (from COMPONENTS hive)
            displayCbsIdentities = () =>
            {
                contentGrid.Columns.Clear();
                contentGrid.Rows.Clear();
                networkSplitContainer.Visible = false;
                firewallPanel.Visible = false;
                contentGrid.Visible = true;
                appxFilterPanel.Visible = false;
                currentSection = null; // Clear current section since this is a custom view

                // Hide overview panel and CBS Components search panel
                if (componentsOverviewPanel != null) componentsOverviewPanel.Visible = false;
                if (cbsComponentsSearchPanel != null) cbsComponentsSearchPanel.Visible = false;

                // Load identities scan data if not cached
                if (cbsIdentitiesFullList == null)
                {
                    // Show loading cursor during potentially long scan
                    var previousCursor = Cursor.Current;
                    Cursor.Current = Cursors.WaitCursor;
                    try
                    {
                        cbsIdentitiesFullList = _infoExtractor!.GetComponentIdentitiesAnalysis();
                    }
                    finally
                    {
                        Cursor.Current = previousCursor;
                    }
                }

                if (cbsIdentitiesFullList.Count == 0)
                {
                    contentGrid.Columns.Add("info", "Information");
                    contentGrid.Rows.Add("No component identity data found in DerivedData\\Components");
                    return;
                }

                // Create info panel if not exists
                if (cbsIdentitiesInfoPanel == null)
                {
                    cbsIdentitiesInfoPanel = new FlowLayoutPanel
                    {
                        Dock = DockStyle.Top,
                        Height = DpiHelper.Scale(35),
                        BackColor = ModernTheme.TreeViewBack,
                        Padding = DpiHelper.ScalePadding(5),
                        AutoSize = false,
                        FlowDirection = FlowDirection.LeftToRight
                    };

                    var infoLabel = new Label
                    {
                        Text = "Identity Scan: Checking for invalid (non-ASCII) characters in component identities",
                        AutoSize = true,
                        ForeColor = ModernTheme.TextSecondary,
                        Font = ModernTheme.RegularFont,
                        Margin = DpiHelper.ScalePadding(5, 6, 5, 0)
                    };
                    cbsIdentitiesInfoPanel.Controls.Add(infoLabel);

                    // Insert info panel above content grid
                    var gridIndex = rightPanel.Controls.IndexOf(contentGrid);
                    rightPanel.Controls.Add(cbsIdentitiesInfoPanel);
                    rightPanel.Controls.SetChildIndex(cbsIdentitiesInfoPanel, gridIndex);
                }

                // Show info panel with proper docking order
                cbsIdentitiesInfoPanel.Visible = true;
                // Ensure proper docking order: contentGrid (Fill) must be at index 0 (processed first)
                // Info panel (Top) must be after contentGrid so it docks above it
                rightPanel.Controls.SetChildIndex(contentGrid, 0);
                rightPanel.Controls.SetChildIndex(cbsIdentitiesInfoPanel, 1);

                // Show identities in a two-column grid (Name, Status)
                contentGrid.Columns.Add("component", "Component / Summary");
                contentGrid.Columns.Add("status", "Status");
                contentGrid.Columns["component"].FillWeight = 80;
                contentGrid.Columns["status"].FillWeight = 20;

                foreach (var item in cbsIdentitiesFullList)
                {
                    var rowIndex = contentGrid.Rows.Add(item.Name, item.Value);
                    contentGrid.Rows[rowIndex].Tag = item;
                    
                    // Highlight rows with issues
                    if (item.Value.Contains("invalid") || item.Value.Contains("issue"))
                    {
                        contentGrid.Rows[rowIndex].DefaultCellStyle.ForeColor = Color.FromArgb(255, 100, 100); // Red-ish
                    }
                    else if (item.Name == "Scan Summary" && item.Value.Contains("No issues"))
                    {
                        contentGrid.Rows[rowIndex].DefaultCellStyle.ForeColor = Color.FromArgb(100, 200, 100); // Green-ish
                    }
                }

                // Update subcategory button states - Components button stays highlighted
                foreach (var btn in subCategoryButtons)
                {
                    if (btn.Tag is string s && s == "Components")
                    {
                        btn.BackColor = ModernTheme.Accent;
                        btn.Invalidate();
                    }
                }
            };

            // Function to display a section in the grid
            Action<AnalysisSection> displaySection = (section) =>
            {
                currentSection = section;
                contentGrid.Columns.Clear();
                contentGrid.Rows.Clear();

                // Hide special panels when showing other sections
                networkSplitContainer.Visible = false;
                firewallPanel.Visible = false;
                appxFilterPanel.Visible = false;
                contentGrid.Visible = true;
                
                // Hide CBS Components search panel
                if (cbsComponentsSearchPanel != null) cbsComponentsSearchPanel.Visible = false;
                
                // Hide CBS Identities info panel
                if (cbsIdentitiesInfoPanel != null) cbsIdentitiesInfoPanel.Visible = false;
                
                // Special handling for Guest Agent - show Extensions sub-tab
                if (section.Title == "☁️ Guest Agent")
                {
                    currentGuestAgentSubView = ""; // Reset - showing overview, not Extensions
                    createGuestAgentSubButtons();
                    // Continue to display the Guest Agent overview content below
                }

                // Special handling for Network Interfaces - use master-detail view
                if (section.Title.Contains("Network Interfaces"))
                {
                    contentGrid.Visible = false;
                    networkSplitContainer.Visible = true;

                    // Populate adapters list (fresh state)
                    networkAdaptersList.Items.Clear();
                    networkAdaptersCache.Clear();

                    foreach (var item in section.Items)
                    {
                        var adapter = new NetworkAdapterItem
                        {
                            DisplayName = item.Name,
                            RegistryPath = item.RegistryPath ?? "",
                            FullGuid = item.RegistryValue ?? ""
                        };

                        if (item.SubItems != null)
                        {
                            foreach (var sub in item.SubItems)
                            {
                                adapter.Properties.Add(new NetworkPropertyItem
                                {
                                    Name = sub.Name,
                                    Value = sub.Value ?? "",
                                    RegistryValueName = sub.RegistryValue ?? (sub.Name switch
                                    {
                                        "DHCP" => "EnableDHCP",
                                        "IP Address" => "DhcpIPAddress / IPAddress",
                                        "Subnet" => "DhcpSubnetMask / SubnetMask",
                                        "Gateway" => "DhcpDefaultGateway / DefaultGateway",
                                        "DNS" => "DhcpNameServer / NameServer",
                                        "DHCP Server" => "DhcpServer",
                                        "Domain" => "DhcpDomain / Domain",
                                        _ => sub.Name
                                    }),
                                    RegistryPath = sub.RegistryPath ?? item.RegistryPath ?? ""
                                });
                            }
                        }

                        networkAdaptersCache.Add(adapter);
                        networkAdaptersList.Items.Add(adapter);
                    }

                    // Select first adapter if available
                    if (networkAdaptersList.Items.Count > 0)
                    {
                        networkAdaptersList.SelectedIndex = 0;
                        networkDetailsGrid.ClearSelection();
                    }
                    else
                    {
                        // Nothing to select, ensure grid is empty
                        networkDetailsGrid.Rows.Clear();
                        networkDetailsGrid.Columns.Clear();
                    }

                    // Update subcategory button states to reflect active tab
                    foreach (var btn in subCategoryButtons)
                    {
                        btn.BackColor = btn.Tag == section ? ModernTheme.Accent : ModernTheme.Surface;
                        btn.ForeColor = btn.Tag == section ? Color.White : ModernTheme.TextPrimary;
                    }

                    return;
                }

                // Special handling for Windows Firewall - use profile-based view
                if (section.Title.Contains("Windows Firewall"))
                {
                    contentGrid.Visible = false;
                    firewallPanel.Visible = true;

                    // Clear existing profile buttons
                    var buttonsToRemove = firewallProfileButtonsPanel.Controls.OfType<Button>().ToList();
                    foreach (var btn in buttonsToRemove)
                    {
                        firewallProfileButtonsPanel.Controls.Remove(btn);
                        btn.Dispose();
                    }
                    firewallProfileButtons.Clear();

                    // Define profiles (registry key name, display name for rules, display name for button)
                    var profiles = new[]
                    {
                        ("DomainProfile", "Domain", "Domain"),
                        ("StandardProfile", "Private", "Private"),
                        ("PublicProfile", "Public", "Public")
                    };

                    // Create profile buttons with enabled/disabled status
                    foreach (var (registryKey, profileKey, displayName) in profiles)
                    {
                        var isEnabled = _infoExtractor.IsFirewallProfileEnabled(registryKey);
                        var statusIcon = isEnabled ? "✅" : "❌";
                        var statusText = isEnabled ? "Enabled" : "Disabled";
                        
                        var profileBtn = new Button
                        {
                            Text = $"{statusIcon} {displayName}: {statusText}",
                            BackColor = ModernTheme.Surface,
                            ForeColor = ModernTheme.TextPrimary,
                            FlatStyle = FlatStyle.Flat,
                            Height = DpiHelper.Scale(28),
                            AutoSize = true,
                            MinimumSize = DpiHelper.ScaleSize(140, 28),
                            Font = ModernTheme.RegularFont,
                            Cursor = Cursors.Hand,
                            Margin = DpiHelper.ScalePadding(0, 0, 8, 0),
                            Tag = profileKey,
                            AccessibleName = $"{displayName}: {statusText}"
                        };
                        profileBtn.FlatAppearance.BorderColor = ModernTheme.Border;
                        profileBtn.FlatAppearance.MouseOverBackColor = ModernTheme.Selection;

                        var capturedKey = profileKey;
                        var capturedDisplay = displayName;
                        profileBtn.Click += (s, ev) => displayFirewallRules(capturedKey, capturedDisplay);

                        firewallProfileButtons.Add(profileBtn);
                        firewallProfileButtonsPanel.Controls.Add(profileBtn);
                    }

                    // Auto-select first profile
                    if (profiles.Length > 0)
                    {
                        displayFirewallRules(profiles[0].Item2, profiles[0].Item3);
                    }

                    // Update subcategory button states
                    foreach (var btn in subCategoryButtons)
                    {
                        btn.BackColor = btn.Tag == section ? ModernTheme.Accent : ModernTheme.Surface;
                        btn.ForeColor = btn.Tag == section ? Color.White : ModernTheme.TextPrimary;
                    }

                    return;
                }

                // Special handling for Appx Packages - use filter buttons like Services
                if (section.Title.Contains("Appx Packages"))
                {
                    contentGrid.Visible = true;
                    networkSplitContainer.Visible = false;
                    firewallPanel.Visible = false;

                    // Create filter buttons in the subcategory panel
                    createAppxFilterButtons();
                    
                    // Display InBox packages by default
                    displayAppxWithFilter("InBox");

                    // Update subcategory button states
                    foreach (var btn in subCategoryButtons)
                    {
                        btn.BackColor = btn.Tag == section ? ModernTheme.Accent : ModernTheme.Surface;
                        btn.ForeColor = btn.Tag == section ? Color.White : ModernTheme.TextPrimary;
                    }

                    return;
                }

                // Special handling for CBS Packages - show package groups with filtering
                if (section.Title.Contains("CBS Packages"))
                {
                    contentGrid.Visible = true;
                    networkSplitContainer.Visible = false;
                    firewallPanel.Visible = false;

                    // Reset cache to reload fresh data
                    cbsPackagesSections = null;

                    // Display CBS packages with summary view
                    displayCbsPackages();

                    // Update subcategory button states
                    foreach (var btn in subCategoryButtons)
                    {
                        btn.BackColor = btn.Tag == section ? ModernTheme.Accent : ModernTheme.Surface;
                        btn.ForeColor = btn.Tag == section ? Color.White : ModernTheme.TextPrimary;
                    }

                    return;
                }

                // Special handling for Storage Filters - use filter buttons for Disk/Volume
                if (section.Title == "🔧 Filters")
                {
                    contentGrid.Visible = true;
                    networkSplitContainer.Visible = false;
                    firewallPanel.Visible = false;

                    // Create filter buttons
                    createStorageFilterButtons();
                    
                    // Display Disk filters by default
                    displayStorageWithFilter("Disk");

                    // Update subcategory button states
                    foreach (var btn in subCategoryButtons)
                    {
                        btn.BackColor = btn.Tag == section ? ModernTheme.Accent : ModernTheme.Surface;
                        btn.ForeColor = btn.Tag == section ? Color.White : ModernTheme.TextPrimary;
                    }

                    return;
                }

                // Special handling for User Profiles - use filter buttons for All/Temp
                if (section.Title.Contains("User Profiles"))
                {
                    contentGrid.Visible = true;
                    networkSplitContainer.Visible = false;
                    firewallPanel.Visible = false;

                    cachedProfilesSection = section;
                    createProfileFilterButtons();
                    displayProfilesWithFilter("All");

                    // Update subcategory button states
                    foreach (var btn in subCategoryButtons)
                    {
                        btn.BackColor = btn.Tag == section ? ModernTheme.Accent : ModernTheme.Surface;
                        btn.ForeColor = btn.Tag == section ? Color.White : ModernTheme.TextPrimary;
                    }

                    return;
                }

                // Determine columns based on content
                bool hasSubItems = section.Items.Any(i => i.IsSubSection && i.SubItems?.Count > 0);
                bool hasValues = section.Items.Any(i => !string.IsNullOrEmpty(i.Value));
                bool isTlsSection = section.Title.Contains("TLS") || section.Title.Contains("SSL");

                if (hasSubItems)
                {
                    // For items with subitems (e.g., Installed Programs with version/publisher)
                    contentGrid.Columns.Add("name", "Name");
                    
                    if (isTlsSection)
                    {
                        contentGrid.Columns.Add("client", "Client");
                        contentGrid.Columns.Add("server", "Server");
                    }
                    else if (section.Title.Contains("Installed Programs"))
                    {
                        contentGrid.Columns.Add("version", "Version");
                        contentGrid.Columns.Add("publisher", "Publisher");
                        contentGrid.Columns.Add("installed", "Installed");
                        contentGrid.Columns[0].FillWeight = 35;
                        contentGrid.Columns[1].FillWeight = 20;
                        contentGrid.Columns[2].FillWeight = 30;
                        contentGrid.Columns[3].FillWeight = 15;
                    }
                    else
                    {
                        contentGrid.Columns.Add("detail1", "Detail 1");
                        contentGrid.Columns.Add("detail2", "Detail 2");
                    }
                    
                    if (!section.Title.Contains("Installed Programs"))
                    {
                        contentGrid.Columns[0].FillWeight = 40;
                        contentGrid.Columns[1].FillWeight = 30;
                        contentGrid.Columns[2].FillWeight = 30;
                    }

                    foreach (var item in section.Items)
                    {
                        int rowIdx;
                        if (item.IsSubSection && item.SubItems != null && item.SubItems.Count > 0)
                        {
                            if (section.Title.Contains("Installed Programs"))
                            {
                                // Installed Programs: Version, Publisher, Install Date
                                var version = item.SubItems.Count > 0 ? item.SubItems[0].Value : "";
                                var publisher = item.SubItems.Count > 1 ? item.SubItems[1].Value : "";
                                var installDate = item.SubItems.Count > 2 ? item.SubItems[2].Value : "";
                                rowIdx = contentGrid.Rows.Add(item.Name, version, publisher, installDate);
                            }
                            else if (isTlsSection)
                            {
                                // For TLS, just show the value without the name prefix
                                var detail1 = item.SubItems.Count > 0 ? item.SubItems[0].Value : "";
                                var detail2 = item.SubItems.Count > 1 ? item.SubItems[1].Value : "";
                                rowIdx = contentGrid.Rows.Add(item.Name, detail1, detail2);
                            }
                            else
                            {
                                var detail1 = item.SubItems.Count > 0 ? $"{item.SubItems[0].Name}: {item.SubItems[0].Value}" : "";
                                var detail2 = item.SubItems.Count > 1 ? $"{item.SubItems[1].Name}: {item.SubItems[1].Value}" : "";
                                rowIdx = contentGrid.Rows.Add(item.Name, detail1, detail2);
                            }
                        }
                        else
                        {
                            if (section.Title.Contains("Installed Programs"))
                                rowIdx = contentGrid.Rows.Add(item.Name, item.Value, "", "");
                            else
                                rowIdx = contentGrid.Rows.Add(item.Name, item.Value, "");
                        }
                        contentGrid.Rows[rowIdx].Tag = item;  // Store item for SelectionChanged
                    }
                }
                else if (hasValues)
                {
                    // Simple name-value pairs
                    contentGrid.Columns.Add("name", "Property");
                    contentGrid.Columns.Add("value", "Value");
                    contentGrid.Columns["name"].FillWeight = 35;
                    contentGrid.Columns["value"].FillWeight = 65;

                    foreach (var item in section.Items)
                    {
                        var rowIdx = contentGrid.Rows.Add(item.Name, item.Value);
                        contentGrid.Rows[rowIdx].Tag = item;  // Store item for SelectionChanged
                    }
                }
                else
                {
                    // Just names (list of items)
                    contentGrid.Columns.Add("name", "Name");

                    foreach (var item in section.Items)
                    {
                        var rowIdx = contentGrid.Rows.Add(item.Name);
                        contentGrid.Rows[rowIdx].Tag = item;  // Store item for SelectionChanged
                    }
                }

                // Update button states
                foreach (var btn in subCategoryButtons)
                {
                    btn.BackColor = btn.Tag == section ? ModernTheme.Accent : ModernTheme.Surface;
                    btn.ForeColor = btn.Tag == section ? Color.White : ModernTheme.TextPrimary;
                }
            };

            // Services display state
            List<ServiceInfo> allServicesCache = new();
            var serviceFilterButtons = new List<Button>();
            themeData.ServiceFilterButtons = serviceFilterButtons;
            string currentServiceFilter = "All";

            // Function to display services with filter
            Action<string> displayServicesWithFilter = (filter) =>
            {
                currentServiceFilter = filter;
                contentGrid.Columns.Clear();
                contentGrid.Rows.Clear();

                contentGrid.Columns.Add("name", "Service Name");
                contentGrid.Columns.Add("startType", "Start Type");
                contentGrid.Columns.Add("imagePath", "Image Path");
                contentGrid.Columns["name"].FillWeight = 25;
                contentGrid.Columns["startType"].FillWeight = 15;
                contentGrid.Columns["imagePath"].FillWeight = 60;

                var filtered = filter switch
                {
                    "Disabled" => allServicesCache.Where(s => s.IsDisabled),
                    "Auto" => allServicesCache.Where(s => s.IsAutoStart),
                    "Boot/System" => allServicesCache.Where(s => s.IsBoot || s.IsSystem),
                    "Manual" => allServicesCache.Where(s => s.IsManual),
                    _ => allServicesCache.AsEnumerable()
                };

                // Sort alphabetically by service name
                var sorted = filtered
                    .OrderBy(s => s.ServiceName)
                    .ToList();

                foreach (var svc in sorted)
                {
                    var rowIdx = contentGrid.Rows.Add(svc.ServiceName, svc.StartTypeName, svc.ImagePath);
                    contentGrid.Rows[rowIdx].Tag = svc;  // Store service info for SelectionChanged
                }

                // Update filter button states
                foreach (var btn in serviceFilterButtons)
                {
                    var btnFilter = btn.Tag?.ToString() ?? "";
                    btn.BackColor = btnFilter == filter ? ModernTheme.Accent : ModernTheme.Surface;
                    btn.ForeColor = btnFilter == filter ? Color.White : ModernTheme.TextPrimary;
                }
            };

            // Function to create service filter buttons
            Action createServiceFilterButtons = () =>
            {
                subCategoryPanel.Controls.Clear();
                serviceFilterButtons.Clear();

                var disabledCount = allServicesCache.Count(s => s.IsDisabled);
                var autoCount = allServicesCache.Count(s => s.IsAutoStart);
                var bootCount = allServicesCache.Count(s => s.IsBoot || s.IsSystem);
                var manualCount = allServicesCache.Count(s => s.IsManual);

                var filters = new[]
                {
                    ($"All ({allServicesCache.Count})", "All"),
                    ($"⛔ Disabled ({disabledCount})", "Disabled"),
                    ($"🚀 Auto-Start ({autoCount})", "Auto"),
                    ($"⚡ Boot/System ({bootCount})", "Boot/System"),
                    ($"✋ Manual ({manualCount})", "Manual")
                };

                foreach (var (text, filterKey) in filters)
                {
                    var btn = new Button
                    {
                        Text = text,
                        FlatStyle = FlatStyle.Flat,
                        BackColor = filterKey == "All" ? ModernTheme.Accent : ModernTheme.Surface,
                        ForeColor = filterKey == "All" ? Color.White : ModernTheme.TextPrimary,
                        Font = ModernTheme.RegularFont,
                        Height = DpiHelper.Scale(28),
                        AutoSize = true,
                        Padding = DpiHelper.ScalePadding(8, 0, 8, 0),
                        Margin = DpiHelper.ScalePadding(2),
                        Cursor = Cursors.Hand,
                        Tag = filterKey,
                        AccessibleName = StripEmojiPrefix(text)
                    };
                    btn.FlatAppearance.BorderColor = ModernTheme.Border;
                    btn.FlatAppearance.BorderSize = 1;
                    btn.FlatAppearance.MouseOverBackColor = ModernTheme.Selection;

                    btn.Click += (sender, args) => displayServicesWithFilter(filterKey);

                    serviceFilterButtons.Add(btn);
                    subCategoryPanel.Controls.Add(btn);
                }
            };

            // Load settings for detail panel height
            var analyzeSettings = AppSettings.Load();

            // Registry detail panel - now using SplitContainer for resizable pane
            // Assign to forward-declared variable (captured by lambdas above)
            contentDetailSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                BackColor = ModernTheme.Border,
                Panel1MinSize = 100,   // Minimum content area
                Panel2MinSize = 80,    // Minimum detail pane
                SplitterWidth = 4,     // Slightly thicker for easier grabbing
                FixedPanel = FixedPanel.Panel2  // Keep detail panel size fixed when resizing window
            };
            contentDetailSplit.Panel1.BackColor = ModernTheme.Background;
            contentDetailSplit.Panel2.BackColor = ModernTheme.TreeViewBack;
            themeData.ContentDetailSplit = contentDetailSplit;

            // Detail panel in Panel2
            var detailPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ModernTheme.TreeViewBack,
                Padding = DpiHelper.ScalePadding(10, 5, 10, 5)
            };
            themeData.DetailPanel = detailPanel;

            var registryPathLabel = new TextBox
            {
                Text = "Registry Path:",
                ForeColor = ModernTheme.TextSecondary,
                BackColor = ModernTheme.TreeViewBack,
                Font = ModernTheme.SmallFont,
                Dock = DockStyle.Top,
                Height = DpiHelper.Scale(20),
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                AccessibleName = "Registry Path"
            };
            themeData.RegistryPathLabel = registryPathLabel;

            var registryValueBox = new RichTextBox
            {
                Text = "Select an item to view registry details",
                ForeColor = ModernTheme.TextPrimary,
                BackColor = ModernTheme.TreeViewBack,
                Font = new Font("Consolas", 9.5F),
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                WordWrap = true,
                ScrollBars = RichTextBoxScrollBars.Vertical,
                AccessibleName = "Registry Value Details"
            };
            themeData.RegistryValueBox = registryValueBox;

            detailPanel.Controls.Add(registryValueBox);
            detailPanel.Controls.Add(registryPathLabel);
            contentDetailSplit.Panel2.Controls.Add(detailPanel);

            // Set splitter distance after form loads (calculated from total height)
            form.Load += (s, ev) =>
            {
                try
                {
                    // Calculate splitter distance (distance from top to splitter)
                    // Panel2 (detail) should be at the bottom with saved height
                    int totalHeight = contentDetailSplit.Height;
                    int detailHeight = Math.Min(Math.Max(analyzeSettings.DetailPanelHeight, 80), totalHeight - 100);
                    contentDetailSplit.SplitterDistance = totalHeight - detailHeight;
                }
                catch { }
            };

            // Save splitter position when changed
            contentDetailSplit.SplitterMoved += (s, ev) =>
            {
                try
                {
                    // Save the Panel2 height (detail panel)
                    int detailHeight = contentDetailSplit.Height - contentDetailSplit.SplitterDistance;
                    analyzeSettings.DetailPanelHeight = detailHeight;
                    analyzeSettings.Save();
                }
                catch { }
            };

            // Handle network adapter selection - using main registry labels
            networkAdaptersList.SelectedIndexChanged += (s, ev) =>
            {
                if (networkAdaptersList.SelectedIndex < 0) return;
                selectedNetworkAdapter = networkAdaptersList.SelectedItem as NetworkAdapterItem;
                if (selectedNetworkAdapter == null) return;

                networkDetailsLabel.Text = $"Details - {selectedNetworkAdapter.DisplayName}";
                networkDetailsGrid.Columns.Clear();
                networkDetailsGrid.Rows.Clear();

                networkDetailsGrid.Columns.Add("property", "Property");
                networkDetailsGrid.Columns.Add("value", "Value");
                networkDetailsGrid.Columns["property"].FillWeight = 30;
                networkDetailsGrid.Columns["value"].FillWeight = 70;

                foreach (var prop in selectedNetworkAdapter.Properties)
                {
                    networkDetailsGrid.Rows.Add(prop.Name, prop.Value);
                }

                registryPathLabel.Text = $"Registry Path: {selectedNetworkAdapter.RegistryPath}";
                registryValueBox.Text = "Select a property to view registry details";
            };

            // Handle network details grid selection
            networkDetailsGrid.SelectionChanged += (s, ev) =>
            {
                if (networkDetailsGrid.SelectedRows.Count > 0 && selectedNetworkAdapter != null)
                {
                    var rowIndex = networkDetailsGrid.SelectedRows[0].Index;
                    if (rowIndex >= 0 && rowIndex < selectedNetworkAdapter.Properties.Count)
                    {
                        var prop = selectedNetworkAdapter.Properties[rowIndex];
                        // Use property's own RegistryPath if available, otherwise fall back to adapter's path
                        var pathToShow = !string.IsNullOrEmpty(prop.RegistryPath) 
                            ? prop.RegistryPath 
                            : selectedNetworkAdapter.RegistryPath;
                        registryPathLabel.Text = $"Registry Path: {pathToShow}";
                        registryValueBox.Text = $"{prop.Name} = {prop.Value} | Registry Value: {prop.RegistryValueName}";
                    }
                }
            };

            // Track current category
            string currentCategory = "";

            // Handle firewall rules grid selection for registry info
            firewallRulesGrid.SelectionChanged += (s, ev) =>
            {
                if (firewallRulesGrid.SelectedRows.Count > 0)
                {
                    var rowIndex = firewallRulesGrid.SelectedRows[0].Index;
                    // Need to map back to sorted rules
                    var sortedRules = currentFirewallRules
                        .OrderByDescending(r => r.IsActive && r.Action.Equals("Block", StringComparison.OrdinalIgnoreCase))
                        .ThenByDescending(r => r.IsActive)
                        .ThenBy(r => r.Name)
                        .ToList();

                    if (rowIndex >= 0 && rowIndex < sortedRules.Count)
                    {
                        var rule = sortedRules[rowIndex];
                        registryPathLabel.Text = $"Registry Path: {rule.RegistryPath}";
                        
                        // Show detailed parsed fields and raw data
                        var details = new System.Text.StringBuilder();
                        details.AppendLine($"Rule Name: {rule.Name}");
                        if (!string.IsNullOrEmpty(rule.Description))
                            details.AppendLine($"Description: {rule.Description}");
                        details.AppendLine($"Action: {rule.Action}");
                        details.AppendLine($"Direction: {rule.Direction}");
                        details.AppendLine($"Active: {(rule.IsActive ? "Yes" : "No")}");
                        if (!string.IsNullOrEmpty(rule.Protocol))
                            details.AppendLine($"Protocol: {rule.Protocol}");
                        if (!string.IsNullOrEmpty(rule.Profiles))
                            details.AppendLine($"Profiles: {rule.Profiles}");
                        if (!string.IsNullOrEmpty(rule.LocalPorts))
                            details.AppendLine($"Local Ports: {rule.LocalPorts}");
                        if (!string.IsNullOrEmpty(rule.RemotePorts))
                            details.AppendLine($"Remote Ports: {rule.RemotePorts}");
                        if (!string.IsNullOrEmpty(rule.LocalAddresses))
                            details.AppendLine($"Local Addresses: {rule.LocalAddresses}");
                        if (!string.IsNullOrEmpty(rule.RemoteAddresses))
                            details.AppendLine($"Remote Addresses: {rule.RemoteAddresses}");
                        if (!string.IsNullOrEmpty(rule.Application))
                            details.AppendLine($"Application: {rule.Application}");
                        if (!string.IsNullOrEmpty(rule.Service))
                            details.AppendLine($"Service: {rule.Service}");
                        if (!string.IsNullOrEmpty(rule.PackageFamilyName))
                            details.AppendLine($"Package Family Name: {rule.PackageFamilyName}");
                        if (!string.IsNullOrEmpty(rule.EmbedContext))
                            details.AppendLine($"Context: {rule.EmbedContext}");
                        details.AppendLine();
                        details.AppendLine($"Raw Data: {rule.RawData}");
                        
                        registryValueBox.Text = details.ToString();
                    }
                }
            };

            // Handle grid selection to show registry path
            contentGrid.SelectionChanged += (s, ev) =>
            {
                if (contentGrid.SelectedRows.Count > 0)
                {
                    var rowIndex = contentGrid.SelectedRows[0].Index;
                    var selectedRow = contentGrid.SelectedRows[0];
                    
                    // Handle Services category specially - use Tag to get correct service after sort
                    if (currentCategory == "Services")
                    {
                        if (selectedRow.Tag is ServiceInfo service)
                        {
                            registryPathLabel.Text = $"Registry Path: {service.RegistryPath}";
                            var details = new List<string>();
                            if (!string.IsNullOrEmpty(service.DisplayName) && service.DisplayName != service.ServiceName)
                                details.Add($"Display Name: {service.DisplayName}");
                            if (!string.IsNullOrEmpty(service.Description))
                                details.Add($"Description: {service.Description}");
                            if (!string.IsNullOrEmpty(service.ImagePath))
                                details.Add($"Image Path: {service.ImagePath}");
                            registryValueBox.Text = details.Count > 0 ? string.Join(" | ", details) : service.ServiceName;
                        }
                        return;
                    }
                    
                    // Handle Appx Packages specially - use Tag to get correct package tuple after sort
                    if (currentSection?.Title.Contains("Appx Packages") == true)
                    {
                        if (selectedRow.Tag is ValueTuple<string, string, string, string> pkg)
                        {
                            registryPathLabel.Text = $"Registry Path: {pkg.Item4}";
                            registryValueBox.Text = $"Package: {pkg.Item1} | Version: {pkg.Item2} | Arch: {pkg.Item3}";
                        }
                        return;
                    }
                    
                    // Handle Storage Filters specially - use Tag to get AnalysisItem
                    if (currentCategory == "Storage" && contentGrid.Columns.Count == 2 && 
                        contentGrid.Columns.Contains("name") && contentGrid.Columns.Contains("value"))
                    {
                        if (selectedRow.Tag is AnalysisItem storageItem)
                        {
                            registryPathLabel.Text = $"Registry Path: {storageItem.RegistryPath ?? "N/A"}";
                            registryValueBox.Text = storageItem.RegistryValue ?? storageItem.Value ?? "No details available";
                            return;
                        }
                    }
                    
                    // Handle CBS Packages specially - get item from row Tag
                    // Check for any CBS sub-view columns (group for All Packages, session for Pending Sessions, etc.)
                    if (currentCbsSubView != null && (
                        contentGrid.Columns.Contains("group") ||     // All Packages view
                        contentGrid.Columns.Contains("session") ||   // Pending Sessions view
                        contentGrid.Columns.Contains("package") ||   // Pending Packages view
                        contentGrid.Columns.Contains("property")))   // Reboot Status view
                    {
                        if (selectedRow.Tag is AnalysisItem cbsItem)
                        {
                            registryPathLabel.Text = $"Registry Path: {cbsItem.RegistryPath ?? "N/A"}";
                            registryValueBox.Text = cbsItem.RegistryValue ?? cbsItem.Value ?? "No details available";
                            return;
                        }
                    }
                    
                    // Handle regular categories - try Tag first, then fall back to index
                    if (selectedRow.Tag is AnalysisItem tagItem)
                    {
                        if (!string.IsNullOrEmpty(tagItem.RegistryPath))
                        {
                            registryPathLabel.Text = $"Registry Path: {tagItem.RegistryPath}";
                            registryValueBox.Text = tagItem.RegistryValue ?? tagItem.Value ?? "No details available";
                        }
                        else
                        {
                            registryPathLabel.Text = "Registry Path:";
                            registryValueBox.Text = tagItem.Value ?? "No registry information available for this item";
                        }
                    }
                    else if (currentSection != null && rowIndex >= 0 && rowIndex < currentSection.Items.Count)
                    {
                        var item = currentSection.Items[rowIndex];
                        if (!string.IsNullOrEmpty(item.RegistryPath))
                        {
                            registryPathLabel.Text = $"Registry Path: {item.RegistryPath}";
                            registryValueBox.Text = item.RegistryValue;
                        }
                        else
                        {
                            registryPathLabel.Text = "Registry Path:";
                            registryValueBox.Text = "No registry information available for this item";
                        }
                    }
                }
            };

            // Add content controls to Panel1 of the split container
            contentDetailSplit.Panel1.Controls.Add(firewallPanel);
            contentDetailSplit.Panel1.Controls.Add(networkSplitContainer);
            contentDetailSplit.Panel1.Controls.Add(contentGrid);

            // Add the split container and other controls to rightPanel
            rightPanel.Controls.Add(contentDetailSplit);
            rightPanel.Controls.Add(appxFilterPanel);
            rightPanel.Controls.Add(subCategoryPanel);
            rightPanel.Controls.Add(contentHeader);

            // Cache for loaded content
            var contentCache = new Dictionary<string, List<AnalysisSection>>();

            // Handle category selection
            categoryList.SelectedIndexChanged += (s, ev) =>
            {
                if (categoryList.SelectedIndex < 0 || categoryList.SelectedItem == null) return;
                var selected = ((string text, string key))categoryList.SelectedItem;
                var key = selected.key;

                // Skip if category is disabled (MouseDown handler shows the message)
                if (!enabledCategories.Contains(key)) return;

                contentTitle.Text = selected.text;
                currentCategory = key;

                // Reset registry info panel
                registryPathLabel.Text = "Registry Path:";
                registryValueBox.Text = "Select an item to view registry details";

                // Hide appx filter panel when changing categories
                appxFilterPanel.Visible = false;
                
                // Hide CBS Components search panel when changing categories
                if (cbsComponentsSearchPanel != null) cbsComponentsSearchPanel.Visible = false;

                // Special handling for Services
                if (key == "Services")
                {
                    currentSection = null;
                    
                    // Hide network panel and show content grid
                    networkSplitContainer.Visible = false;
                    contentGrid.Visible = true;
                    
                    if (allServicesCache.Count == 0)
                    {
                        allServicesCache = _infoExtractor.GetAllServices();
                    }

                    createServiceFilterButtons();
                    displayServicesWithFilter("All");
                    return;
                }

                // Regular category handling
                if (!contentCache.ContainsKey(key))
                {
                    List<AnalysisSection> sections = key switch
                    {
                        "System" => _infoExtractor.GetSystemAnalysis(),
                        "Profiles" => _infoExtractor.GetUserAnalysis(),
                        "Network" => _infoExtractor.GetNetworkAnalysis(),
                        "RDP" => _infoExtractor.GetRdpAnalysis(),
                        "Update" => _infoExtractor.GetUpdateAnalysis(),
                        "Storage" => _infoExtractor.GetStorageAnalysis(),
                        "Software" => _infoExtractor.GetSoftwareAnalysis(),
                        _ => new List<AnalysisSection>()
                    };
                    contentCache[key] = sections;
                }

                currentSections = contentCache[key];

                // Create subcategory buttons
                subCategoryPanel.Controls.Clear();
                subCategoryButtons.Clear();

                // Reuse single shared ToolTip for all subcategory buttons (more efficient, prevents leaks)
                themeData.SubCategoryToolTip ??= new ToolTip { InitialDelay = 200, ReshowDelay = 100, AutoPopDelay = 5000 };
                var subCategoryToolTip = themeData.SubCategoryToolTip;
                subCategoryToolTip.RemoveAll();  // Clear old associations

                // Define which subcategories require which hive type (for System category)
                // Hive Information is available for both
                var softwareHiveSubcategories = new HashSet<string> { "🪟 Build Information" };
                var systemHiveSubcategories = new HashSet<string> { 
                    "💻 Computer Information", "🔄 CPU Hyper-Threading", 
                    "💥 Crash Dump Configuration", "🕐 System Time Config"
                };
                var bothHiveSubcategories = new HashSet<string> { "📁 Hive Information", "☁️ Guest Agent" };
                
                // Define SOFTWARE-based Update subcategories (shown as grayed when COMPONENTS hive loaded)
                var updateSoftwareSubcategories = new List<string> {
                    "📋 Update Policy",
                    "🏢 Windows Update for Business",
                    "📦 Delivery Optimization",
                    "📜 Update Configuration",
                    "🔧 Servicing Stack Update (SSU)",
                    "📦 CBS Packages"
                };
                
                var isSoftwareHive = _parser.CurrentHiveType == OfflineRegistryParser.HiveType.SOFTWARE;
                var isSystemHive = _parser.CurrentHiveType == OfflineRegistryParser.HiveType.SYSTEM;
                var isComponentsHive = _parser.CurrentHiveType == OfflineRegistryParser.HiveType.COMPONENTS;
                // Use theme-aware color for grayed out buttons
                var grayedOutColor = ModernTheme.TextDisabled;

                // Create list of buttons with availability info for sorting
                var availableButtons = new List<Button>();
                var unavailableButtons = new List<Button>();

                // Helper function to determine availability
                bool IsSubcategoryAvailable(string title, out string reqHive)
                {
                    reqHive = "";
                    
                    // For Update category: all subcategories require SOFTWARE hive (except CBS Components for COMPONENTS hive)
                    if (key == "Update")
                    {
                        if (isComponentsHive)
                        {
                            // COMPONENTS hive: all SOFTWARE-based subcategories are unavailable
                            reqHive = "SOFTWARE";
                            return false;
                        }
                        return isSoftwareHive;
                    }
                    
                    if (key != "System") return true;
                    
                    if (bothHiveSubcategories.Contains(title))
                        return true;
                    if (softwareHiveSubcategories.Contains(title))
                    {
                        reqHive = "SOFTWARE";
                        return isSoftwareHive;
                    }
                    if (systemHiveSubcategories.Contains(title))
                    {
                        reqHive = "SYSTEM";
                        return isSystemHive;
                    }
                    return true;
                }

                foreach (var section in currentSections)
                {
                    // Skip "Notice" sections for Update category when COMPONENTS hive is loaded
                    // These are placeholder sections that don't provide useful functionality
                    if (key == "Update" && isComponentsHive && section.Title.Contains("Notice"))
                    {
                        continue;
                    }

                    string requiredHive;
                    bool isAvailable = IsSubcategoryAvailable(section.Title, out requiredHive);

                    // Use owner-draw for proper gray color rendering
                    var btn = new Button
                    {
                        Text = "",  // We'll draw text ourselves
                        FlatStyle = FlatStyle.Flat,
                        BackColor = isAvailable ? ModernTheme.Surface : ModernTheme.TreeViewBack,
                        Font = ModernTheme.RegularFont,
                        Height = DpiHelper.Scale(28),
                        AutoSize = false,
                        Margin = DpiHelper.ScalePadding(2),
                        Cursor = isAvailable ? Cursors.Hand : Cursors.Default,
                        Tag = section,
                        AccessibleName = StripEmojiPrefix(section.Title)
                    };
                    
                    // Measure text to set button width
                    using (var g = btn.CreateGraphics())
                    {
                        var textSize = g.MeasureString(section.Title, ModernTheme.RegularFont);
                        btn.Width = (int)textSize.Width + DpiHelper.Scale(20);
                    }
                    
                    btn.FlatAppearance.BorderColor = isAvailable ? ModernTheme.Border : grayedOutColor;
                    btn.FlatAppearance.BorderSize = 1;
                    btn.FlatAppearance.MouseOverBackColor = isAvailable ? ModernTheme.Selection : ModernTheme.TreeViewBack;

                    // Custom paint for proper gray text
                    var sectionTitle = section.Title;
                    var sectionAvailable = isAvailable;
                    btn.Paint += (s, ev) =>
                    {
                        ev.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                        var textColor = sectionAvailable ? ModernTheme.TextPrimary : grayedOutColor;
                        // Use TextRenderer for better Unicode/emoji support
                        TextRenderer.DrawText(
                            ev.Graphics,
                            sectionTitle,
                            ModernTheme.RegularFont,
                            btn.ClientRectangle,
                            textColor,
                            TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine
                        );
                    };

                    // Tooltip for unavailable subcategories
                    if (!isAvailable && !string.IsNullOrEmpty(requiredHive))
                    {
                        subCategoryToolTip.SetToolTip(btn, $"This feature requires {requiredHive} hive to be loaded");
                    }

                    var capturedRequiredHive = requiredHive;
                    btn.Click += (sender, args) =>
                    {
                        if (!sectionAvailable)
                        {
                            MessageBox.Show(
                                $"This feature requires {capturedRequiredHive} hive to be loaded.\n\n" +
                                $"Common location: C:\\Windows\\System32\\config\\{capturedRequiredHive}",
                                $"{capturedRequiredHive} Hive Required",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                        }
                        else
                        {
                            displaySection((AnalysisSection)btn.Tag);
                        }
                    };

                    subCategoryButtons.Add(btn);
                    
                    if (isAvailable)
                        availableButtons.Add(btn);
                    else
                        unavailableButtons.Add(btn);
                }

                // Add grayed-out SOFTWARE-based Update subcategory buttons when COMPONENTS hive is loaded
                if (key == "Update" && isComponentsHive)
                {
                    foreach (var subcatTitle in updateSoftwareSubcategories)
                    {
                        var softwareBtn = new Button
                        {
                            Text = "",  // We'll draw text ourselves
                            FlatStyle = FlatStyle.Flat,
                            BackColor = ModernTheme.TreeViewBack,
                            Font = ModernTheme.RegularFont,
                            Height = DpiHelper.Scale(28),
                            AutoSize = false,
                            Margin = DpiHelper.ScalePadding(2),
                            Cursor = Cursors.Default,
                            Tag = "SoftwareUpdateFeature",
                            AccessibleName = StripEmojiPrefix(subcatTitle)
                        };
                        
                        // Measure text to set button width
                        using (var g = softwareBtn.CreateGraphics())
                        {
                            var textSize = g.MeasureString(subcatTitle, ModernTheme.RegularFont);
                            softwareBtn.Width = (int)textSize.Width + DpiHelper.Scale(20);
                        }
                        
                        softwareBtn.FlatAppearance.BorderColor = grayedOutColor;
                        softwareBtn.FlatAppearance.BorderSize = 1;
                        softwareBtn.FlatAppearance.MouseOverBackColor = ModernTheme.TreeViewBack;

                        // Custom paint for gray text
                        var capturedTitle = subcatTitle;
                        softwareBtn.Paint += (s, ev) =>
                        {
                            ev.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                            TextRenderer.DrawText(
                                ev.Graphics,
                                capturedTitle,
                                ModernTheme.RegularFont,
                                softwareBtn.ClientRectangle,
                                grayedOutColor,
                                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine
                            );
                        };

                        // Tooltip explaining the requirement
                        subCategoryToolTip.SetToolTip(softwareBtn, "This feature requires SOFTWARE hive to be loaded");

                        // Click handler showing message
                        softwareBtn.Click += (sender, args) =>
                        {
                            MessageBox.Show(
                                "This feature requires SOFTWARE hive to be loaded.\n\n" +
                                "Common location: C:\\Windows\\System32\\config\\SOFTWARE",
                                "SOFTWARE Hive Required",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                        };

                        subCategoryButtons.Add(softwareBtn);
                        unavailableButtons.Add(softwareBtn);
                    }
                }

                // Add Activation button to System category (requires SOFTWARE hive)
                // Create it BEFORE adding to panel so it can be sorted properly
                Button? activationBtn = null;
                if (key == "System")
                {
                    // Use owner-draw button to ensure proper color rendering
                    activationBtn = new Button
                    {
                        Text = "",  // We'll draw the text ourselves
                        FlatStyle = FlatStyle.Flat,
                        BackColor = isSoftwareHive ? ModernTheme.Surface : ModernTheme.TreeViewBack,
                        Font = ModernTheme.RegularFont,
                        Size = DpiHelper.ScaleSize(110, 28),
                        Margin = DpiHelper.ScalePadding(2),
                        Cursor = isSoftwareHive ? Cursors.Hand : Cursors.Default,
                        Tag = isSoftwareHive ? _infoExtractor.GetActivationAnalysis() : null,
                        AccessibleName = "Activation"
                    };
                    activationBtn.FlatAppearance.BorderColor = isSoftwareHive ? ModernTheme.Border : grayedOutColor;
                    activationBtn.FlatAppearance.BorderSize = 1;
                    activationBtn.FlatAppearance.MouseOverBackColor = isSoftwareHive ? ModernTheme.Selection : ModernTheme.TreeViewBack;

                    // Custom paint to draw icon and text with proper colors
                    var isSoftware = isSoftwareHive; // Capture for closure
                    activationBtn.Paint += (s, ev) =>
                    {
                        ev.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                        var textColor = isSoftware ? ModernTheme.TextPrimary : grayedOutColor;
                        
                        // Draw icon using Segoe MDL2 Assets (key icon)
                        var iconFont = _iconFont10;
                        TextRenderer.DrawText(ev.Graphics, "\uE8D7", iconFont, DpiHelper.ScalePoint(8, 6), textColor);
                        
                        // Draw text
                        TextRenderer.DrawText(ev.Graphics, "Activation", ModernTheme.RegularFont, DpiHelper.ScalePoint(28, 5), textColor);
                    };

                    // Use shared ToolTip for activation button
                    if (!isSoftwareHive)
                    {
                        subCategoryToolTip.SetToolTip(activationBtn, "This feature requires SOFTWARE hive to be loaded");
                    }
                    else
                    {
                        subCategoryToolTip.SetToolTip(activationBtn, "View Windows activation and KMS settings");
                    }

                    activationBtn.Click += (sender, args) =>
                    {
                        if (!isSoftware)
                        {
                            MessageBox.Show(
                                "This feature requires SOFTWARE hive to be loaded.\n\n" +
                                "Windows activation and KMS settings are stored in the SOFTWARE hive.\n\n" +
                                "Common location: C:\\Windows\\System32\\config\\SOFTWARE",
                                "SOFTWARE Hive Required",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                        }
                        else if (activationBtn.Tag is AnalysisSection activationSection)
                        {
                            displaySection(activationSection);
                        }
                    };

                    subCategoryButtons.Add(activationBtn);
                    
                    // Add to correct list for sorting
                    if (isSoftwareHive)
                        availableButtons.Add(activationBtn);
                    else
                        unavailableButtons.Add(activationBtn);
                }

                // Add Components button to Update category (requires COMPONENTS hive)
                // Single button that shows an overview with CBS Packages and CBS Identities options
                Button? componentsBtn = null;
                
                if (key == "Update")
                {
                    var isComponents = isComponentsHive; // Capture for closure
                    
                    // Create Components button
                    componentsBtn = new Button
                    {
                        Text = "",
                        FlatStyle = FlatStyle.Flat,
                        BackColor = isComponentsHive ? ModernTheme.Surface : ModernTheme.TreeViewBack,
                        Font = ModernTheme.RegularFont,
                        Size = DpiHelper.ScaleSize(120, 28),
                        Margin = DpiHelper.ScalePadding(2),
                        Cursor = isComponentsHive ? Cursors.Hand : Cursors.Default,
                        Tag = "Components",
                        AccessibleName = "Components"
                    };
                    componentsBtn.FlatAppearance.BorderColor = isComponentsHive ? ModernTheme.Border : grayedOutColor;
                    componentsBtn.FlatAppearance.BorderSize = 1;
                    componentsBtn.FlatAppearance.MouseOverBackColor = isComponentsHive ? ModernTheme.Selection : ModernTheme.TreeViewBack;
                    
                    componentsBtn.Paint += (s, ev) =>
                    {
                        ev.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                        var textColor = isComponents ? ModernTheme.TextPrimary : grayedOutColor;
                        TextRenderer.DrawText(ev.Graphics, "Components", ModernTheme.RegularFont, 
                            componentsBtn.ClientRectangle, textColor,
                            TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                    };
                    
                    componentsBtn.Click += (sender, args) =>
                    {
                        if (!isComponents)
                        {
                            MessageBox.Show(
                                "This feature requires COMPONENTS hive to be loaded.\n\n" +
                                "The COMPONENTS hive contains the Windows Component Store data.\n\n" +
                                "Common location: C:\\Windows\\System32\\config\\COMPONENTS",
                                "COMPONENTS Hive Required",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                        }
                        else
                        {
                            // Display Components overview (landing page with options)
                            displayComponentsOverview();
                            
                            // Update button state
                            componentsBtn.BackColor = ModernTheme.Accent;
                            componentsBtn.Invalidate();
                        }
                    };
                    
                    if (!isComponentsHive)
                        subCategoryToolTip.SetToolTip(componentsBtn, "This feature requires COMPONENTS hive to be loaded");
                    else
                        subCategoryToolTip.SetToolTip(componentsBtn, "Analyze Component Store (CBS Packages and Identities)");
                    
                    subCategoryButtons.Add(componentsBtn);
                    if (isComponentsHive)
                        availableButtons.Add(componentsBtn);
                    else
                        unavailableButtons.Add(componentsBtn);
                }

                // Add available buttons first, then unavailable (greyed out) buttons
                foreach (var btn in availableButtons)
                    subCategoryPanel.Controls.Add(btn);
                foreach (var btn in unavailableButtons)
                    subCategoryPanel.Controls.Add(btn);

                // Select first available section by default
                AnalysisSection? firstAvailable = null;
                
                // For COMPONENTS hive + Update category, show Components overview (no auto-load)
                if (key == "Update" && isComponentsHive)
                {
                    // Show Components overview when clicking Update category with COMPONENTS hive
                    displayComponentsOverview();
                    
                    // Highlight the Components button as selected
                    foreach (var btn in subCategoryButtons)
                    {
                        if (btn.Tag is string tag && tag == "Components")
                        {
                            btn.BackColor = ModernTheme.Accent;
                            btn.Invalidate();
                        }
                    }
                }
                else
                {
                    foreach (var section in currentSections)
                    {
                        bool isAvail = true;
                        if (key == "System")
                        {
                            if (bothHiveSubcategories.Contains(section.Title))
                                isAvail = true;
                            else if (softwareHiveSubcategories.Contains(section.Title))
                                isAvail = isSoftwareHive;
                            else if (systemHiveSubcategories.Contains(section.Title))
                                isAvail = isSystemHive;
                        }
                        if (isAvail)
                        {
                            firstAvailable = section;
                            break;
                        }
                    }
                    
                    if (firstAvailable != null)
                    {
                        displaySection(firstAvailable);
                    }
                    else if (key == "System" && isSoftwareHive)
                    {
                        // For SOFTWARE hive, show Activation by default
                        var activationSection = _infoExtractor.GetActivationAnalysis();
                        displaySection(activationSection);
                    }
                }
            };

            splitContainer.Panel1.Controls.Add(leftPanel);
            splitContainer.Panel2.Controls.Add(rightPanel);
            form.Controls.Add(splitContainer);

            // Select first enabled item by default (if any)
            int firstEnabledIndex = -1;
            for (int i = 0; i < categoryList.Items.Count; i++)
            {
                var item = ((string text, string key))categoryList.Items[i];
                if (enabledCategories.Contains(item.key))
                {
                    firstEnabledIndex = i;
                    break;
                }
            }

            form.FormClosed += (s, ev) => { 
                themeData.Dispose();  // Dispose fonts and tooltips tracked by themeData
                _analyzeForm = null; // Clear the reference
                form.Dispose(); 
            };
            
            // Show form immediately, then load data after it's visible
            form.Show();
            form.BringToFront();
            form.Activate();
            
            // Defer initial category selection to after form is shown
            // This allows the form to appear instantly
            form.BeginInvoke((Action)(() => {
                categoryList.SelectedIndex = firstEnabledIndex;
            }));
        }


        public void NavigateToKey(string keyPath)
        {
            NavigateToKey(keyPath, null);
        }

        public void NavigateToKey(string keyPath, string? valueName)
        {
            if (_treeView.Nodes.Count == 0) return;

            var parts = keyPath.Split('\\');
            TreeNode? currentNode = _treeView.Nodes[0];
            
            foreach (var part in parts.Skip(1))
            {
                if (currentNode.Nodes.Count == 1 && currentNode.Nodes[0].Tag?.ToString() == "placeholder")
                {
                    currentNode.Expand();
                }

                TreeNode? found = null;
                foreach (TreeNode child in currentNode.Nodes)
                {
                    if (child.Text.Equals(part, StringComparison.OrdinalIgnoreCase))
                    {
                        found = child;
                        break;
                    }
                }

                if (found == null) break;
                currentNode = found;
            }

            _treeView.SelectedNode = currentNode;
            currentNode?.EnsureVisible();
            _treeView.Focus();

            // If a value name was specified, select it in the ListView
            if (!string.IsNullOrEmpty(valueName))
            {
                SelectValueInListView(valueName);
            }
        }

        private void SelectValueInListView(string valueName)
        {
            // Handle "(Default)" which represents empty value name
            var searchName = valueName == "(Default)" ? "" : valueName;
            
            foreach (ListViewItem item in _listView.Items)
            {
                var itemName = item.Text == "(Default)" ? "" : item.Text;
                if (itemName.Equals(searchName, StringComparison.OrdinalIgnoreCase))
                {
                    item.Selected = true;
                    item.Focused = true;
                    item.EnsureVisible();
                    _listView.Focus();
                    break;
                }
            }
        }

        private void SwitchTheme(ThemeType theme)
        {
            ModernTheme.SetTheme(theme);
            ApplyThemeToAll();
            
            // Recreate drop panel with new theme colors
            RefreshDropPanel();
            
            // Refresh analyze window theme without reloading data
            if (_analyzeForm != null && !_analyzeForm.IsDisposed)
            {
                RefreshAnalyzeFormTheme(_analyzeForm);
            }
            
            // Refresh search window theme without reloading data
            if (_searchForm != null && !_searchForm.IsDisposed)
            {
                _searchForm.RefreshTheme();
            }
            
            // Refresh statistics window theme
            if (_statisticsForm != null && !_statisticsForm.IsDisposed)
            {
                RefreshStatisticsFormTheme(_statisticsForm);
            }
            
            // Refresh timeline window theme
            if (_timelineForm != null && !_timelineForm.IsDisposed && _timelineForm is TimelineForm tf)
            {
                tf.RefreshTheme();
            }
            
            // Update menu checkmarks
            foreach (ToolStripMenuItem item in _menuStrip.Items)
            {
                if (item.Text == "&View")
                {
                    foreach (var dropItem in item.DropDownItems)
                    {
                        if (dropItem is ToolStripMenuItem menuItem)
                        {
                            if (menuItem.Text == "&Dark Theme")
                                menuItem.Checked = theme == ThemeType.Dark;
                            else if (menuItem.Text == "&Light Theme")
                                menuItem.Checked = theme == ThemeType.Light;
                        }
                    }
                }
            }
        }

        private void RefreshAnalyzeFormTheme(Form form)
        {
            // Apply base form theme
            ModernTheme.ApplyTo(form);
            
            // Get theme data from form tag
            if (form.Tag is not AnalyzeFormThemeData themeData) return;

            // Update split containers
            if (themeData.MainSplit != null)
            {
                themeData.MainSplit.BackColor = ModernTheme.Border;
                themeData.MainSplit.Panel1.BackColor = ModernTheme.Background;
                themeData.MainSplit.Panel2.BackColor = ModernTheme.Background;
            }
            if (themeData.NetworkSplit != null)
            {
                themeData.NetworkSplit.BackColor = ModernTheme.Border;
                themeData.NetworkSplit.Panel1.BackColor = ModernTheme.Background;
                themeData.NetworkSplit.Panel2.BackColor = ModernTheme.Background;
            }

            // Update left panel
            if (themeData.LeftPanel != null) themeData.LeftPanel.BackColor = ModernTheme.Background;
            if (themeData.CategoryHeader != null) themeData.CategoryHeader.BackColor = ModernTheme.Surface;
            if (themeData.CategoryTitle != null) themeData.CategoryTitle.ForeColor = ModernTheme.TextPrimary;
            if (themeData.CategoryList != null)
            {
                themeData.CategoryList.BackColor = ModernTheme.TreeViewBack;
                themeData.CategoryList.ForeColor = ModernTheme.TextPrimary;
                themeData.CategoryList.Invalidate(); // Force redraw for custom drawing
            }

            // Update right panel
            if (themeData.RightPanel != null) themeData.RightPanel.BackColor = ModernTheme.Background;
            if (themeData.ContentHeader != null) themeData.ContentHeader.BackColor = ModernTheme.Surface;
            if (themeData.ContentTitle != null) themeData.ContentTitle.ForeColor = ModernTheme.TextPrimary;
            if (themeData.SubCategoryPanel != null) themeData.SubCategoryPanel.BackColor = ModernTheme.Surface;

            // Update content grid
            if (themeData.ContentGrid != null)
            {
                var grid = themeData.ContentGrid;
                grid.BackgroundColor = ModernTheme.Surface;
                grid.ForeColor = ModernTheme.TextPrimary;
                grid.GridColor = ModernTheme.Border;
                grid.DefaultCellStyle.BackColor = ModernTheme.Surface;
                grid.DefaultCellStyle.ForeColor = ModernTheme.TextPrimary;
                grid.DefaultCellStyle.SelectionBackColor = ModernTheme.Selection;
                grid.DefaultCellStyle.SelectionForeColor = ModernTheme.TextPrimary;
                grid.AlternatingRowsDefaultCellStyle.BackColor = ModernTheme.ListViewAltRow;
                grid.ColumnHeadersDefaultCellStyle.BackColor = ModernTheme.Surface;
                grid.ColumnHeadersDefaultCellStyle.ForeColor = ModernTheme.TextSecondary;
                grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = ModernTheme.Surface;
                grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = ModernTheme.TextSecondary;
            }

            // Update registry info panel (now using ContentDetailSplit and DetailPanel)
            if (themeData.ContentDetailSplit != null)
            {
                themeData.ContentDetailSplit.BackColor = ModernTheme.Border;
                themeData.ContentDetailSplit.Panel1.BackColor = ModernTheme.Background;
                themeData.ContentDetailSplit.Panel2.BackColor = ModernTheme.TreeViewBack;
            }
            if (themeData.DetailPanel != null) themeData.DetailPanel.BackColor = ModernTheme.TreeViewBack;
            if (themeData.RegistryPathLabel != null)
            {
                themeData.RegistryPathLabel.ForeColor = ModernTheme.TextSecondary;
                themeData.RegistryPathLabel.BackColor = ModernTheme.TreeViewBack;
            }
            if (themeData.RegistryValueBox != null)
            {
                themeData.RegistryValueBox.ForeColor = ModernTheme.TextPrimary;
                themeData.RegistryValueBox.BackColor = ModernTheme.TreeViewBack;
            }

            // Update network adapter controls
            if (themeData.NetworkAdaptersList != null)
            {
                themeData.NetworkAdaptersList.BackColor = ModernTheme.TreeViewBack;
                themeData.NetworkAdaptersList.ForeColor = ModernTheme.TextPrimary;
                themeData.NetworkAdaptersList.Invalidate();
            }
            if (themeData.NetworkAdaptersHeader != null) themeData.NetworkAdaptersHeader.BackColor = ModernTheme.Surface;
            if (themeData.NetworkAdaptersLabel != null) themeData.NetworkAdaptersLabel.ForeColor = ModernTheme.TextSecondary;
            if (themeData.NetworkDetailsHeader != null) themeData.NetworkDetailsHeader.BackColor = ModernTheme.Surface;
            if (themeData.NetworkDetailsLabel != null) themeData.NetworkDetailsLabel.ForeColor = ModernTheme.TextSecondary;

            // Update network details grid
            if (themeData.NetworkDetailsGrid != null)
            {
                var grid = themeData.NetworkDetailsGrid;
                grid.BackgroundColor = ModernTheme.Surface;
                grid.ForeColor = ModernTheme.TextPrimary;
                grid.GridColor = ModernTheme.Border;
                grid.DefaultCellStyle.BackColor = ModernTheme.Surface;
                grid.DefaultCellStyle.ForeColor = ModernTheme.TextPrimary;
                grid.DefaultCellStyle.SelectionBackColor = ModernTheme.Selection;
                grid.DefaultCellStyle.SelectionForeColor = ModernTheme.TextPrimary;
                grid.AlternatingRowsDefaultCellStyle.BackColor = ModernTheme.ListViewAltRow;
                grid.ColumnHeadersDefaultCellStyle.BackColor = ModernTheme.TreeViewBack;
                grid.ColumnHeadersDefaultCellStyle.ForeColor = ModernTheme.TextSecondary;
                grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = ModernTheme.TreeViewBack;
                grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = ModernTheme.TextSecondary;
            }

            // Update subcategory buttons
            foreach (var btn in themeData.SubCategoryButtons)
            {
                bool isActive = btn.BackColor == ModernTheme.Accent || btn.ForeColor == Color.White;
                if (!isActive)
                {
                    btn.BackColor = ModernTheme.Surface;
                    btn.ForeColor = ModernTheme.TextPrimary;
                }
                btn.FlatAppearance.BorderColor = ModernTheme.Border;
                btn.FlatAppearance.MouseOverBackColor = ModernTheme.Selection;
            }

            // Update service filter buttons
            foreach (var btn in themeData.ServiceFilterButtons)
            {
                bool isActive = btn.BackColor == ModernTheme.Accent || btn.ForeColor == Color.White;
                if (!isActive)
                {
                    btn.BackColor = ModernTheme.Surface;
                    btn.ForeColor = ModernTheme.TextPrimary;
                }
                btn.FlatAppearance.BorderColor = ModernTheme.Border;
                btn.FlatAppearance.MouseOverBackColor = ModernTheme.Selection;
            }

            // Update Appx filter panel
            if (themeData.AppxFilterPanel != null) themeData.AppxFilterPanel.BackColor = ModernTheme.TreeViewBack;

            // Update Appx filter buttons
            foreach (var btn in themeData.AppxFilterButtons)
            {
                bool isActive = btn.BackColor == ModernTheme.Accent || btn.ForeColor == Color.White;
                if (!isActive)
                {
                    btn.BackColor = ModernTheme.Surface;
                    btn.ForeColor = ModernTheme.TextPrimary;
                }
                btn.FlatAppearance.BorderColor = ModernTheme.Border;
                btn.FlatAppearance.MouseOverBackColor = ModernTheme.Selection;
            }

            // Update Storage filter buttons
            foreach (var btn in themeData.StorageFilterButtons)
            {
                bool isActive = btn.BackColor == ModernTheme.Accent || btn.ForeColor == Color.White;
                if (!isActive)
                {
                    btn.BackColor = ModernTheme.Surface;
                    btn.ForeColor = ModernTheme.TextPrimary;
                }
                btn.FlatAppearance.BorderColor = ModernTheme.Border;
                btn.FlatAppearance.MouseOverBackColor = ModernTheme.Selection;
            }

            // Update firewall panel controls
            if (themeData.FirewallPanel != null) themeData.FirewallPanel.BackColor = ModernTheme.Background;
            if (themeData.FirewallProfileButtonsPanel != null) themeData.FirewallProfileButtonsPanel.BackColor = ModernTheme.Surface;
            if (themeData.FirewallProfileLabel != null) themeData.FirewallProfileLabel.ForeColor = ModernTheme.TextSecondary;
            if (themeData.FirewallRulesPanel != null) themeData.FirewallRulesPanel.BackColor = ModernTheme.Background;
            if (themeData.FirewallRulesHeader != null) themeData.FirewallRulesHeader.BackColor = ModernTheme.Surface;
            if (themeData.FirewallRulesLabel != null) themeData.FirewallRulesLabel.ForeColor = ModernTheme.TextSecondary;

            // Update firewall rules grid
            if (themeData.FirewallRulesGrid != null)
            {
                var grid = themeData.FirewallRulesGrid;
                grid.BackgroundColor = ModernTheme.Surface;
                grid.ForeColor = ModernTheme.TextPrimary;
                grid.GridColor = ModernTheme.Border;
                grid.DefaultCellStyle.BackColor = ModernTheme.Surface;
                grid.DefaultCellStyle.ForeColor = ModernTheme.TextPrimary;
                grid.DefaultCellStyle.SelectionBackColor = ModernTheme.Selection;
                grid.DefaultCellStyle.SelectionForeColor = ModernTheme.TextPrimary;
                grid.AlternatingRowsDefaultCellStyle.BackColor = ModernTheme.ListViewAltRow;
                grid.ColumnHeadersDefaultCellStyle.BackColor = ModernTheme.TreeViewBack;
                grid.ColumnHeadersDefaultCellStyle.ForeColor = ModernTheme.TextSecondary;
                grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = ModernTheme.TreeViewBack;
                grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = ModernTheme.TextSecondary;
            }

            // Update firewall profile buttons
            foreach (var btn in themeData.FirewallProfileButtons)
            {
                bool isActive = btn.BackColor == ModernTheme.Accent || btn.ForeColor == Color.White;
                if (!isActive)
                {
                    btn.BackColor = ModernTheme.Surface;
                    btn.ForeColor = ModernTheme.TextPrimary;
                }
                btn.FlatAppearance.BorderColor = ModernTheme.Border;
                btn.FlatAppearance.MouseOverBackColor = ModernTheme.Selection;
            }

            // Refresh firewall display to re-apply row colors
            themeData.RefreshFirewallDisplay?.Invoke();

            form.Invalidate(true);
        }

        private void RefreshStatisticsFormTheme(Form form)
        {
            // Apply base form theme
            ModernTheme.ApplyTo(form);
            
            // Recursively update all controls
            RefreshStatisticsControls(form);
            
            form.Invalidate(true);
        }

        private void RefreshStatisticsControls(Control parent)
        {
            foreach (Control ctrl in parent.Controls)
            {
                switch (ctrl)
                {
                    case TabControl tabControl:
                        ModernTheme.ApplyTo(tabControl);
                        foreach (TabPage page in tabControl.TabPages)
                        {
                            page.BackColor = ModernTheme.Background;
                        }
                        break;
                        
                    case TabPage tabPage:
                        tabPage.BackColor = ModernTheme.Background;
                        break;
                        
                    case TreeView tree:
                        tree.BackColor = ModernTheme.Background;
                        tree.ForeColor = ModernTheme.TextPrimary;
                        tree.Invalidate();
                        break;
                        
                    case Label label:
                        if (label.ForeColor != ModernTheme.Accent && 
                            label.ForeColor != ModernTheme.Success &&
                            label.ForeColor != ModernTheme.Warning &&
                            label.ForeColor != ModernTheme.Info)
                        {
                            label.ForeColor = label.ForeColor == ModernTheme.TextSecondary 
                                ? ModernTheme.TextSecondary 
                                : ModernTheme.TextPrimary;
                        }
                        break;
                        
                    case FlowLayoutPanel flow:
                        if (flow.BackColor != Color.Transparent)
                            flow.BackColor = ModernTheme.Background;
                        break;
                        
                    case Panel panel:
                        if (panel.BackColor != Color.Transparent)
                            panel.BackColor = ModernTheme.Background;
                        // Update header panel (Surface color, with border paint)
                        if (panel.Height == 100)
                        {
                            panel.BackColor = ModernTheme.Surface;
                        }
                        break;
                }
                
                // Recurse into children
                if (ctrl.HasChildren)
                {
                    RefreshStatisticsControls(ctrl);
                }
            }
        }

        private void ApplyThemeToAll()
        {
            // Apply to main form
            this.BackColor = ModernTheme.Background;
            this.ForeColor = ModernTheme.TextPrimary;

            // Apply to menu
            _menuStrip.BackColor = ModernTheme.Surface;
            _menuStrip.ForeColor = ModernTheme.TextPrimary;
            _menuStrip.Renderer = new ModernMenuRenderer();

            // Apply to toolbar panel and buttons
            _toolbarPanel.BackColor = ModernTheme.Surface;
            foreach (Control ctrl in _toolbarPanel.Controls)
            {
                if (ctrl is FlowLayoutPanel flow)
                {
                    foreach (Control c in flow.Controls)
                    {
                        if (c is Button btn)
                        {
                            btn.ForeColor = ModernTheme.TextPrimary;
                            btn.FlatAppearance.MouseOverBackColor = ModernTheme.Selection;
                        }
                        else if (c is Panel separator)
                        {
                            separator.BackColor = ModernTheme.Border;
                        }
                    }
                }
            }

            // Apply to tree view
            _treeView.BackColor = ModernTheme.TreeViewBack;
            _treeView.ForeColor = ModernTheme.TextPrimary;
            _treeView.LineColor = ModernTheme.Border;
            _treeView.Invalidate();

            // Apply to list view - need to re-apply theme for owner-draw
            ModernTheme.ApplyTo(_listView);
            _listView.Refresh();

            // Apply to details box
            _detailsBox.BackColor = ModernTheme.Surface;
            _detailsBox.ForeColor = ModernTheme.TextPrimary;

            // Apply to status panel
            _statusPanel.BackColor = ModernTheme.Surface;
            _statusForeColor = ModernTheme.TextSecondary;
            _statusPanel.Invalidate();
            _hiveTypeLabel.BackColor = ModernTheme.Surface;

            // Apply to splitters
            _mainSplitContainer.BackColor = ModernTheme.Border;
            _mainSplitContainer.Panel1.BackColor = ModernTheme.Background;
            _mainSplitContainer.Panel2.BackColor = ModernTheme.Background;
            _rightSplitContainer.BackColor = ModernTheme.Border;
            _rightSplitContainer.Panel1.BackColor = ModernTheme.Background;
            _rightSplitContainer.Panel2.BackColor = ModernTheme.Background;

            // Apply to bookmark panels
            if (_bookmarkBar != null)
            {
                _bookmarkBar.BackColor = ModernTheme.Surface;
                _bookmarkBar.Invalidate();
            }
            if (_bookmarkPanel != null)
            {
                _bookmarkPanel.BackColor = ModernTheme.Surface;
                foreach (Control c in _bookmarkPanel.Controls)
                {
                    if (c is Panel collapseBar && collapseBar.Tag as string == "bookmarkCollapseBar")
                    {
                        collapseBar.BackColor = ModernTheme.Surface;
                        collapseBar.Invalidate(); // Repaints arrow, text, and icon
                    }
                    else if (c is FlowLayoutPanel itemsPanel && itemsPanel.Tag as string == "bookmarkItems")
                    {
                        itemsPanel.BackColor = ModernTheme.Surface;
                        foreach (Control item in itemsPanel.Controls)
                        {
                            if (item is Label lbl)
                            {
                                lbl.ForeColor = ModernTheme.TextPrimary;
                                lbl.BackColor = ModernTheme.Surface;
                            }
                        }
                    }
                }
            }

            // Apply to all panels recursively (section headers, etc.)
            ApplyThemeToPanels(this);

            // Refresh all controls
            this.Refresh();
        }

        private void ApplyThemeToPanels(Control parent)
        {
            foreach (Control ctrl in parent.Controls)
            {
                if (ctrl is Panel panel)
                {
                    // Check if it's a section header (height 36 with label)
                    if (panel.Height == 36 && panel.Controls.Count >= 1)
                    {
                        foreach (Control c in panel.Controls)
                        {
                            if (c is Label lbl && lbl.Dock == DockStyle.Fill)
                            {
                                panel.BackColor = ModernTheme.Surface;
                                lbl.ForeColor = ModernTheme.TextPrimary;
                            }
                            else if (c is Panel border && border.Dock == DockStyle.Bottom && border.Height == 1)
                            {
                                border.BackColor = ModernTheme.Border;
                            }
                        }
                    }
                    else if (panel.BackColor != Color.Transparent)
                    {
                        // Regular panel - apply background
                        if (panel != _toolbarPanel && panel != _statusPanel)
                        {
                            panel.BackColor = ModernTheme.Background;
                        }
                    }
                }
                
                // Recurse into child controls
                if (ctrl.HasChildren)
                {
                    ApplyThemeToPanels(ctrl);
                }
            }
        }

        #region Update Checking

        /// <summary>
        /// Silently checks for updates on startup. Only shows dialog if update is available.
        /// </summary>
        private async Task CheckForUpdatesOnStartupAsync()
        {
            // Delay to let UI fully load
            await Task.Delay(2000).ConfigureAwait(true);
            
            var updateInfo = await UpdateChecker.CheckForUpdatesAsync().ConfigureAwait(true);
            
            // Only show dialog if update is available (silent on error or up-to-date)
            if (updateInfo?.UpdateAvailable == true)
            {
                ShowUpdateDialog(updateInfo, isManualCheck: false);
            }
        }

        /// <summary>
        /// Manual check for updates triggered from Help menu.
        /// </summary>
        private async void CheckForUpdates_Click(object? sender, EventArgs e)
        {
            // Show "checking" dialog with DPI-aware sizing
            using var checkingForm = new Form
            {
                Text = "Check for Updates",
                Size = DpiHelper.ScaleSize(300, 120),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                ShowInTaskbar = false,
                ControlBox = false
            };
            ModernTheme.ApplyTo(checkingForm);
            
            var checkingLabel = new Label
            {
                Text = "Checking for updates...",
                Font = ModernTheme.RegularFont,
                ForeColor = ModernTheme.TextPrimary,
                AutoSize = false,
                Size = DpiHelper.ScaleSize(280, 40),
                Location = DpiHelper.ScalePoint(10, 35),
                TextAlign = ContentAlignment.MiddleCenter
            };
            checkingForm.Controls.Add(checkingLabel);
            
            checkingForm.Show(this);
            checkingForm.Refresh();
            
            var updateInfo = await UpdateChecker.CheckForUpdatesAsync().ConfigureAwait(true);
            
            checkingForm.Close();
            
            if (updateInfo == null)
            {
                MessageBox.Show(this, "Unable to check for updates. Please check your internet connection.",
                    "Update Check Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                ShowUpdateDialog(updateInfo, isManualCheck: true);
            }
        }

        /// <summary>
        /// Shows the update dialog. If no update available and manual check, shows "up to date" message.
        /// </summary>
        private void ShowUpdateDialog(UpdateInfo info, bool isManualCheck)
        {
            // If no update and this is manual check, show "up to date" message
            if (!info.UpdateAvailable)
            {
                MessageBox.Show(this, 
                    $"You're up to date!\n\nRegistry Expert {info.CurrentVersion} is the latest version.",
                    "No Updates Available", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            
            // Show update available dialog
            using var form = new Form
            {
                Text = "Update Available",
                Size = DpiHelper.ScaleSize(500, 450),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                ShowInTaskbar = false
            };
            ModernTheme.ApplyTo(form);
            
            // Main panel with padding
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = DpiHelper.ScalePadding(25),
                BackColor = ModernTheme.Background
            };
            
            // Title: "A new version is available!"
            var titleLabel = new Label
            {
                Text = "A new version is available!",
                Font = new Font("Segoe UI Semibold", 14F),
                ForeColor = ModernTheme.Accent,
                AutoSize = true,
                Location = DpiHelper.ScalePoint(25, 20)
            };
            
            // Version info panel
            var versionPanel = new Panel
            {
                Location = DpiHelper.ScalePoint(25, 55),
                Size = DpiHelper.ScaleSize(440, 60),
                BackColor = ModernTheme.Surface
            };
            
            var currentLabel = new Label
            {
                Text = $"Current version: {info.CurrentVersion}",
                Font = ModernTheme.RegularFont,
                ForeColor = ModernTheme.TextSecondary,
                AutoSize = true,
                Location = DpiHelper.ScalePoint(15, 10)
            };
            
            var latestLabel = new Label
            {
                Text = $"Latest version: {info.LatestVersion}",
                Font = ModernTheme.BoldFont,
                ForeColor = ModernTheme.TextPrimary,
                AutoSize = true,
                Location = DpiHelper.ScalePoint(15, 32)
            };
            
            versionPanel.Controls.AddRange(new Control[] { currentLabel, latestLabel });
            
            // "What's New" section header
            var whatsNewHeader = ModernTheme.CreateSectionHeader("What's New");
            whatsNewHeader.Location = DpiHelper.ScalePoint(25, 130);
            whatsNewHeader.Size = DpiHelper.ScaleSize(440, 25);
            
            // Release notes RichTextBox (scrollable)
            var notesBox = new RichTextBox
            {
                Location = DpiHelper.ScalePoint(25, 160),
                Size = DpiHelper.ScaleSize(440, 180),
                Text = info.ReleaseNotes,
                Font = ModernTheme.DataFont,
                ForeColor = ModernTheme.TextPrimary,
                BackColor = ModernTheme.Surface,
                BorderStyle = BorderStyle.None,
                ReadOnly = true,
                ScrollBars = RichTextBoxScrollBars.Vertical
            };
            
            // Button panel at bottom
            var buttonPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Location = DpiHelper.ScalePoint(25, 355),
                Size = DpiHelper.ScaleSize(440, 45),
                BackColor = Color.Transparent
            };
            
            var laterButton = ModernTheme.CreateSecondaryButton("Later");
            laterButton.Size = DpiHelper.ScaleSize(110, 35);
            laterButton.Click += (s, ev) => form.Close();
            
            var downloadButton = ModernTheme.CreateButton("Download");
            downloadButton.Size = DpiHelper.ScaleSize(110, 35);
            downloadButton.Margin = DpiHelper.ScalePadding(0, 0, 10, 0);
            downloadButton.Click += (s, ev) =>
            {
                try
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = info.ReleaseUrl,
                        UseShellExecute = true
                    });
                    form.Close();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error opening URL: {ex.Message}");
                    MessageBox.Show(form, "Could not open the download page. Please visit GitHub manually.",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            
            buttonPanel.Controls.Add(downloadButton);
            buttonPanel.Controls.Add(laterButton);
            
            panel.Controls.AddRange(new Control[] 
            { 
                titleLabel, versionPanel, whatsNewHeader, notesBox, buttonPanel 
            });
            form.Controls.Add(panel);
            
            form.ShowDialog(this);
        }

        #endregion

        private void About_Click(object? sender, EventArgs e)
        {
            using var form = new Form
            {
                Text = "About Registry Expert",
                Size = DpiHelper.ScaleSize(480, 450),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                ShowInTaskbar = false
            };
            ModernTheme.ApplyTo(form);

            var panel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(30) };
            panel.BackColor = ModernTheme.Background;
            
            // App icon - use the application's icon
            var iconBox = new PictureBox
            {
                Size = DpiHelper.ScaleSize(64, 64),
                Location = DpiHelper.ScalePoint(30, 25),
                BackColor = Color.Transparent,
                SizeMode = PictureBoxSizeMode.Zoom
            };
            
            // Get icon from the form's icon (which is set from registry_fixed.ico)
            if (this.Icon != null)
            {
                iconBox.Image = this.Icon.ToBitmap();
            }
            
            var titleLabel = new Label
            {
                Text = "Registry Expert",
                Font = new Font("Segoe UI Semibold", 22F),
                ForeColor = ModernTheme.Accent,
                AutoSize = true,
                Location = DpiHelper.ScalePoint(105, 30)
            };
            
            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            var versionLabel = new Label
            {
                Text = $"Version {version?.Major ?? 1}.{version?.Minor ?? 0}.{version?.Build ?? 0}",
                Font = ModernTheme.RegularFont,
                ForeColor = ModernTheme.TextSecondary,
                AutoSize = true,
                Location = DpiHelper.ScalePoint(105, 68)
            };
            
            var descLabel = new Label
            {
                Text = "A lightweight tool for viewing and analyzing\noffline Windows registry hive files.",
                Font = ModernTheme.RegularFont,
                ForeColor = ModernTheme.TextPrimary,
                AutoSize = true,
                Location = DpiHelper.ScalePoint(30, 110)
            };
            
            // Author section
            var authorTitleLabel = new Label
            {
                Text = "Author",
                Font = new Font("Segoe UI Semibold", 10F),
                ForeColor = ModernTheme.TextSecondary,
                AutoSize = true,
                Location = DpiHelper.ScalePoint(30, 165)
            };
            
            var authorLabel = new Label
            {
                Text = "Bowen Zhang",
                Font = ModernTheme.RegularFont,
                ForeColor = ModernTheme.TextPrimary,
                AutoSize = true,
                Location = DpiHelper.ScalePoint(30, 185)
            };
            
            // Feedback section
            var feedbackTitleLabel = new Label
            {
                Text = "Feedback",
                Font = new Font("Segoe UI Semibold", 10F),
                ForeColor = ModernTheme.TextSecondary,
                AutoSize = true,
                Location = DpiHelper.ScalePoint(30, 220)
            };
            
            var emailLink = new LinkLabel
            {
                Text = "bowenzhang@microsoft.com",
                Font = ModernTheme.RegularFont,
                AutoSize = true,
                Location = DpiHelper.ScalePoint(30, 240),
                LinkColor = ModernTheme.Accent,
                ActiveLinkColor = ModernTheme.AccentDark,
                VisitedLinkColor = ModernTheme.Accent
            };
            emailLink.Click += (s, ev) =>
            {
                try
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = "mailto:bowenzhang@microsoft.com?subject=Registry Expert Feedback",
                        UseShellExecute = true
                    });
                }
                catch { }
            };
            
            // Data Protection Notice link
            var privacyLink = new LinkLabel
            {
                Text = "Data Protection Notice",
                Font = ModernTheme.RegularFont,
                AutoSize = true,
                Location = DpiHelper.ScalePoint(30, 270),
                LinkColor = ModernTheme.Accent,
                ActiveLinkColor = ModernTheme.AccentDark,
                VisitedLinkColor = ModernTheme.Accent
            };
            privacyLink.Click += (s, ev) =>
            {
                try
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = "https://www.microsoft.com/en-us/privacy/data-privacy-notice",
                        UseShellExecute = true
                    });
                }
                catch { }
            };
            
            // Supported hives
            var hivesLabel = new Label
            {
                Text = "Supported Hives: NTUSER.DAT, SAM, SECURITY, SOFTWARE,\nSYSTEM, USRCLASS.DAT, Amcache.hve",
                Font = new Font("Segoe UI", 8F),
                ForeColor = ModernTheme.TextDisabled,
                AutoSize = true,
                Location = DpiHelper.ScalePoint(30, 300)
            };
            
            var closeBtn = ModernTheme.CreateButton("Close", (s, e) => form.Close());
            closeBtn.Location = DpiHelper.ScalePoint(185, 350);
            closeBtn.Width = DpiHelper.Scale(100);
            
            panel.Controls.AddRange(new Control[] { iconBox, titleLabel, versionLabel, descLabel, 
                authorTitleLabel, authorLabel, feedbackTitleLabel, emailLink, privacyLink, hivesLabel, closeBtn });
            form.Controls.Add(panel);
            form.ShowDialog(this);
        }

        private void ShowError(string message)
        {
            MessageBox.Show(message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void ShowInfo(string message, string title = "Information")
        {
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // Hide form immediately so user sees instant close
            this.Visible = false;
            
            // CRITICAL: Dispose MenuStrip FIRST to unsubscribe from SystemEvents.UserPreferenceChanged
            // and SystemEvents.SessionSwitch static events. These events hold strong references that
            // prevent the process from terminating even after Environment.Exit() is called.
            try
            {
                if (_menuStrip != null)
                {
                    this.Controls.Remove(_menuStrip);
                    _menuStrip.Dispose();
                    _menuStrip = null!;
                }
            }
            catch { }
            
            // Disable drag-drop to revoke shell registration (prevents explorer.exe dependency)
            try
            {
                this.AllowDrop = false;
            }
            catch { }
            
            // Close and dispose child forms first (they have their own cleanup handlers)
            try
            {
                _searchForm?.Close();
                _searchForm = null;
                
                _analyzeForm?.Close();
                _analyzeForm = null;
                
                _statisticsForm?.Close();
                _statisticsForm = null;
                
                _compareForm?.Close();  // CompareForm has its own cleanup in FormClosing
                _compareForm = null;
                
                _timelineForm?.Close();
                _timelineForm = null;
            }
            catch { }
            
            // Clear TreeView/ListView to release RegistryKey references held in Tag properties
            try
            {
                _treeView?.Nodes.Clear();
                _listView?.Items.Clear();
                _detailsBox?.Clear();
            }
            catch { }
            
            // Dispose RichTextBox to release OLE/COM objects that can keep process alive
            try
            {
                if (_detailsBox != null)
                {
                    _detailsBox.Parent?.Controls.Remove(_detailsBox);
                    _detailsBox.Dispose();
                    _detailsBox = null!;
                }
            }
            catch { }
            
            // Dispose the parser to release memory
            try
            {
                _parser?.Dispose();
                _parser = null;
            }
            catch { }
            
            _infoExtractor = null;
            _previousSelectedNode = null;
            
            base.OnFormClosing(e);
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e);
            
            // Explicitly exit the application to ensure all threads and SystemEvents are cleaned up.
            // SystemEvents.UserPreferenceChanged (subscribed by WinForms controls internally) can keep
            // the process alive if not explicitly terminated.
            Application.Exit();
            
            // Force terminate the process. SystemEvents maintains static event subscriptions
            // with references to WindowsFormsSynchronizationContext that can keep the process alive
            // even after Application.Exit(). Environment.Exit ensures clean termination.
            Environment.Exit(0);
        }

        private const int WM_CLOSE = 0x0010;

        
        protected override void WndProc(ref Message m)
        {
            // IMMEDIATELY exit on WM_CLOSE - don't let anything else keep the process alive
            if (m.Msg == WM_CLOSE)
            {
                // Hide form first for instant visual feedback
                this.Visible = false;
                
                // Use multiple termination methods - one of them WILL work
                try { Application.Exit(); } catch { }
                try { Environment.Exit(0); } catch { }
                try { Process.GetCurrentProcess().Kill(); } catch { }
                
                return; // Never reached, but for clarity
            }
            
            base.WndProc(ref m);
        }







    }

    // Data models for structured analysis
    public class AnalysisSection
    {
        public string Title { get; set; } = "";
        public List<AnalysisItem> Items { get; set; } = new();
    }

    public class AnalysisItem
    {
        public string Name { get; set; } = "";
        public string Value { get; set; } = "";
        public string RegistryPath { get; set; } = "";
        public string RegistryValue { get; set; } = "";
        public bool IsSubSection { get; set; } = false;
        public List<AnalysisItem>? SubItems { get; set; }
    }

    // Network adapter helper classes for master-detail view
    public class NetworkAdapterItem
    {
        public string DisplayName { get; set; } = "";
        public string RegistryPath { get; set; } = "";
        public string FullGuid { get; set; } = "";
        public List<NetworkPropertyItem> Properties { get; set; } = new();
        
        public override string ToString() => DisplayName;
    }

    public class NetworkPropertyItem
    {
        public string Name { get; set; } = "";
        public string Value { get; set; } = "";
        public string RegistryValueName { get; set; } = "";
        public string RegistryPath { get; set; } = "";
    }
}

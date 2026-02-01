using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace RegistryExpert
{
    /// <summary>
    /// Theme type enumeration
    /// </summary>
    public enum ThemeType
    {
        Dark,
        Light
    }

    /// <summary>
    /// Modern theme colors and styling utilities with Dark/Light support
    /// </summary>
    public static class ModernTheme
    {
        // Current theme
        private static ThemeType _currentTheme = ThemeType.Dark;
        public static ThemeType CurrentTheme => _currentTheme;

        // Event for theme changes
        public static event EventHandler? ThemeChanged;

        // Track themed ListViews to prevent duplicate event handler registration
        private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<ListView, object> _themedListViews = new();
        
        // Track themed TabControls to prevent duplicate event handler registration
        private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<TabControl, object> _themedTabControls = new();
        
        // Lock object for thread-safe theme changes
        private static readonly object _themeLock = new object();

        // Color Palette - Dynamic based on theme (Modern VS Code / Fluent inspired)
        public static Color Background => _currentTheme == ThemeType.Dark 
            ? Color.FromArgb(24, 24, 28)      // Darker, richer background
            : Color.FromArgb(249, 249, 251);
        
        public static Color Surface => _currentTheme == ThemeType.Dark 
            ? Color.FromArgb(32, 32, 36)      // Subtle elevation
            : Color.FromArgb(255, 255, 255);
        
        public static Color SurfaceLight => _currentTheme == ThemeType.Dark 
            ? Color.FromArgb(44, 44, 50)      // Hover states
            : Color.FromArgb(245, 245, 248);
        
        public static Color SurfaceHover => _currentTheme == ThemeType.Dark 
            ? Color.FromArgb(55, 55, 62)      // Card hover
            : Color.FromArgb(235, 235, 240);
        
        public static Color Border => _currentTheme == ThemeType.Dark 
            ? Color.FromArgb(58, 58, 65)      // Subtle borders
            : Color.FromArgb(225, 225, 230);
        
        public static Color BorderLight => _currentTheme == ThemeType.Dark 
            ? Color.FromArgb(75, 75, 82)
            : Color.FromArgb(200, 200, 208);
        
        public static Color TextPrimary => _currentTheme == ThemeType.Dark 
            ? Color.FromArgb(230, 230, 235)   // Softer white
            : Color.FromArgb(28, 28, 35);
        
        public static Color TextSecondary => _currentTheme == ThemeType.Dark 
            ? Color.FromArgb(160, 160, 170)   // Muted text
            : Color.FromArgb(90, 90, 100);
        
        public static Color TextDisabled => _currentTheme == ThemeType.Dark 
            ? Color.FromArgb(70, 70, 80)
            : Color.FromArgb(170, 170, 180);
        
        // Modern accent colors (teal/cyan inspired)
        public static Color Accent => Color.FromArgb(45, 156, 219);       // Modern blue
        public static Color AccentHover => Color.FromArgb(60, 175, 235);  // Lighter hover
        public static Color AccentDark => Color.FromArgb(30, 130, 190);   // Pressed state
        public static Color AccentSubtle => _currentTheme == ThemeType.Dark 
            ? Color.FromArgb(45, 156, 219, 30)  // 30% opacity accent
            : Color.FromArgb(45, 156, 219, 20);
        
        // Status colors (Fluent-inspired)
        public static Color Success => Color.FromArgb(16, 185, 129);      // Modern green
        public static Color Warning => Color.FromArgb(245, 158, 11);      // Warm amber
        public static Color Error => Color.FromArgb(239, 68, 68);         // Modern red
        public static Color Info => Color.FromArgb(99, 102, 241);         // Indigo
        
        // Block/Error row highlighting (for firewall rules, etc.)
        public static Color BlockRowBackground => _currentTheme == ThemeType.Dark 
            ? Color.FromArgb(60, 40, 40)          // Dark reddish
            : Color.FromArgb(255, 235, 235);      // Light pink
        public static Color BlockRowForeground => _currentTheme == ThemeType.Dark 
            ? Color.FromArgb(255, 180, 180)       // Light red text
            : Color.FromArgb(180, 40, 40);        // Dark red text
        
        // Diff colors for Compare feature
        public static Color DiffAdded => _currentTheme == ThemeType.Dark 
            ? Color.FromArgb(46, 160, 67)         // Green for added
            : Color.FromArgb(35, 134, 54);
        public static Color DiffRemoved => _currentTheme == ThemeType.Dark 
            ? Color.FromArgb(218, 54, 51)         // Red for removed
            : Color.FromArgb(207, 34, 46);
        public static Color DiffModified => _currentTheme == ThemeType.Dark 
            ? Color.FromArgb(210, 153, 34)        // Yellow/orange for modified
            : Color.FromArgb(191, 135, 0);
        public static Color DiffAddedBackground => _currentTheme == ThemeType.Dark 
            ? Color.FromArgb(35, 55, 40)          // Subtle green bg
            : Color.FromArgb(230, 255, 235);
        public static Color DiffRemovedBackground => _currentTheme == ThemeType.Dark 
            ? Color.FromArgb(55, 35, 35)          // Subtle red bg
            : Color.FromArgb(255, 235, 235);
        public static Color DiffModifiedBackground => _currentTheme == ThemeType.Dark 
            ? Color.FromArgb(55, 50, 30)          // Subtle yellow bg
            : Color.FromArgb(255, 250, 230);
        
        public static Color TreeViewBack => _currentTheme == ThemeType.Dark
            ? Color.FromArgb(28, 28, 32)      // Slightly different from background
            : Color.FromArgb(252, 252, 254);
        
        public static Color ListViewBack => _currentTheme == ThemeType.Dark 
            ? Color.FromArgb(28, 28, 32)
            : Color.FromArgb(252, 252, 254);
        
        public static Color ListViewAltRow => _currentTheme == ThemeType.Dark 
            ? Color.FromArgb(34, 34, 40)      // Subtle striping
            : Color.FromArgb(247, 247, 250);
        
        public static Color Selection => _currentTheme == ThemeType.Dark 
            ? Color.FromArgb(45, 50, 60)      // Subtle selection
            : Color.FromArgb(230, 242, 255);
        
        public static Color SelectionActive => Color.FromArgb(45, 156, 219);
        
        // Gradient colors for headers
        public static Color GradientStart => _currentTheme == ThemeType.Dark 
            ? Color.FromArgb(35, 35, 42)
            : Color.FromArgb(250, 250, 252);
        
        public static Color GradientEnd => _currentTheme == ThemeType.Dark 
            ? Color.FromArgb(28, 28, 32)
            : Color.FromArgb(245, 245, 248);

        // Fonts (Modern, consistent sizing)
        public static readonly Font RegularFont = new Font("Segoe UI", 9F, FontStyle.Regular);
        public static readonly Font BoldFont = new Font("Segoe UI Semibold", 9F, FontStyle.Regular);
        public static readonly Font MonoFont = new Font("Consolas", 9F, FontStyle.Regular);
        public static readonly Font HeaderFont = new Font("Segoe UI Semibold", 12F, FontStyle.Regular);
        public static readonly Font SmallFont = new Font("Segoe UI", 8F, FontStyle.Regular);
        public static readonly Font TitleFont = new Font("Segoe UI Light", 14F, FontStyle.Regular);
        public static readonly Font TreeFont = new Font("Segoe UI", 9F, FontStyle.Regular);
        
        // Larger fonts for data display areas (Values, Statistics, Compare)
        public static readonly Font DataFont = new Font("Segoe UI", 10F, FontStyle.Regular);
        public static readonly Font DataBoldFont = new Font("Segoe UI Semibold", 10F, FontStyle.Regular);
        public static readonly Font DataMonoFont = new Font("Consolas", 10F, FontStyle.Regular);

        /// <summary>
        /// Switch to a different theme
        /// </summary>
        public static void SetTheme(ThemeType theme)
        {
            lock (_themeLock)
            {
                if (_currentTheme != theme)
                {
                    _currentTheme = theme;
                    ThemeChanged?.Invoke(null, EventArgs.Empty);
                }
            }
        }

        /// <summary>
        /// Toggle between dark and light themes
        /// </summary>
        public static void ToggleTheme()
        {
            SetTheme(_currentTheme == ThemeType.Dark ? ThemeType.Light : ThemeType.Dark);
        }

        /// <summary>
        /// Apply modern theme to a Form
        /// </summary>
        public static void ApplyTo(Form form)
        {
            form.BackColor = Background;
            form.ForeColor = TextPrimary;
            form.Font = RegularFont;
        }

        /// <summary>
        /// Apply modern theme to a TreeView
        /// </summary>
        public static void ApplyTo(TreeView tree)
        {
            tree.BackColor = TreeViewBack;
            tree.ForeColor = TextPrimary;
            tree.Font = TreeFont;
            tree.BorderStyle = BorderStyle.None;
            tree.LineColor = Border;
            tree.Indent = 24;  // Don't scale - let AutoScaleMode handle it
            tree.ItemHeight = 26;  // Don't scale - allows +/- button hit-detection to work properly
            tree.ShowLines = false;  // Cleaner modern look
            tree.ShowPlusMinus = true;  // Ensure expand/collapse buttons work on single click
            tree.FullRowSelect = true;
        }

        /// <summary>
        /// Apply modern theme to a ListView
        /// </summary>
        public static void ApplyTo(ListView list)
        {
            list.BackColor = ListViewBack;
            list.ForeColor = TextPrimary;
            list.Font = RegularFont;
            list.BorderStyle = BorderStyle.None;
            list.OwnerDraw = true;
            list.GridLines = false;
            
            // Enable double buffering to reduce flicker and improve painting
            list.GetType().GetProperty("DoubleBuffered", 
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)?
                .SetValue(list, true);
            
            // Only add handlers if this ListView hasn't been themed yet
            if (!_themedListViews.TryGetValue(list, out _))
            {
                _themedListViews.Add(list, new object());
                list.DrawColumnHeader += ListView_DrawColumnHeader;
                list.DrawItem += ListView_DrawItem;
                list.DrawSubItem += ListView_DrawSubItem;
                
                // Paint the background to cover blank areas (right side, below items)
                list.Paint += (sender, e) =>
                {
                    if (sender is ListView lv)
                    {
                        // Fill the entire client area with background color
                        // This covers the scrollbar area and any blank space
                        using var brush = new SolidBrush(ListViewBack);
                        e.Graphics.FillRectangle(brush, lv.ClientRectangle);
                    }
                };
            }
        }

        private static void ListView_DrawColumnHeader(object? sender, DrawListViewColumnHeaderEventArgs e)
        {
            // Modern gradient header
            using var brush = new LinearGradientBrush(e.Bounds, GradientStart, GradientEnd, LinearGradientMode.Vertical);
            e.Graphics.FillRectangle(brush, e.Bounds);
            
            using var borderPen = new Pen(Border);
            // Bottom border
            e.Graphics.DrawLine(borderPen, e.Bounds.Left, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1);
            // Right border for column separator
            e.Graphics.DrawLine(borderPen, e.Bounds.Right - 1, e.Bounds.Top + 4, e.Bounds.Right - 1, e.Bounds.Bottom - 4);
            
            using var textBrush = new SolidBrush(TextSecondary);
            var textRect = new Rectangle(e.Bounds.X + 12, e.Bounds.Y, e.Bounds.Width - 12, e.Bounds.Height);
            using var sf = new StringFormat { LineAlignment = StringAlignment.Center, Trimming = StringTrimming.EllipsisCharacter };
            e.Graphics.DrawString(e.Header?.Text ?? "", DataBoldFont, textBrush, textRect, sf);
        }

        private static void ListView_DrawItem(object? sender, DrawListViewItemEventArgs e)
        {
            e.DrawDefault = false;
        }

        private static void ListView_DrawSubItem(object? sender, DrawListViewSubItemEventArgs e)
        {
            if (e.Item == null) return;
            
            Color backColor = e.ItemIndex % 2 == 0 ? ListViewBack : ListViewAltRow;
            if (e.Item.Selected)
            {
                backColor = e.Item.ListView?.Focused == true ? SelectionActive : Selection;
            }
            
            using var brush = new SolidBrush(backColor);
            e.Graphics.FillRectangle(brush, e.Bounds);
            
            // Column separator (right border)
            using var borderPen = new Pen(Color.FromArgb(60, Border));
            e.Graphics.DrawLine(borderPen, e.Bounds.Right - 1, e.Bounds.Top + 2, e.Bounds.Right - 1, e.Bounds.Bottom - 2);
            
            Color textColor = e.Item.Selected && e.Item.ListView?.Focused == true ? Color.White : TextPrimary;
            using var textBrush = new SolidBrush(textColor);
            
            var textRect = new Rectangle(e.Bounds.X + 12, e.Bounds.Y, e.Bounds.Width - 12, e.Bounds.Height);
            using var sf = new StringFormat { LineAlignment = StringAlignment.Center, Trimming = StringTrimming.EllipsisCharacter };
            e.Graphics.DrawString(e.SubItem?.Text ?? "", DataFont, textBrush, textRect, sf);
        }

        /// <summary>
        /// Apply modern theme to a RichTextBox
        /// </summary>
        public static void ApplyTo(RichTextBox rtb)
        {
            rtb.BackColor = Surface;
            rtb.ForeColor = TextPrimary;
            rtb.Font = MonoFont;
            rtb.BorderStyle = BorderStyle.None;
            rtb.ReadOnly = true;
            rtb.WordWrap = false;
            rtb.ScrollBars = RichTextBoxScrollBars.Both;
            
            // Enable double buffering to reduce flicker
            rtb.GetType().GetProperty("DoubleBuffered", 
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)?
                .SetValue(rtb, true);
        }

        /// <summary>
        /// Apply modern theme to a DataGridView
        /// </summary>
        public static void ApplyTo(DataGridView grid)
        {
            grid.BackgroundColor = Surface;
            grid.ForeColor = TextPrimary;
            grid.GridColor = Border;
            grid.BorderStyle = BorderStyle.None;
            grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            grid.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            grid.EnableHeadersVisualStyles = false;
            grid.RowHeadersVisible = false;
            grid.AllowUserToAddRows = false;
            grid.AllowUserToDeleteRows = false;
            grid.AllowUserToResizeRows = false;
            grid.ReadOnly = true;
            grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            grid.RowTemplate.Height = DpiHelper.Scale(28);
            grid.Font = DataFont;
            
            // Ensure scrollbars are visible
            grid.ScrollBars = ScrollBars.Both;

            // Cell styles
            grid.DefaultCellStyle.BackColor = Surface;
            grid.DefaultCellStyle.ForeColor = TextPrimary;
            grid.DefaultCellStyle.SelectionBackColor = Selection;
            grid.DefaultCellStyle.SelectionForeColor = TextPrimary;
            grid.DefaultCellStyle.Padding = new Padding(5, 2, 5, 2);

            grid.AlternatingRowsDefaultCellStyle.BackColor = ListViewAltRow;

            // Header styles
            grid.ColumnHeadersDefaultCellStyle.BackColor = TreeViewBack;
            grid.ColumnHeadersDefaultCellStyle.ForeColor = TextSecondary;
            grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = TreeViewBack;
            grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = TextSecondary;
            grid.ColumnHeadersDefaultCellStyle.Font = DataBoldFont;
            grid.ColumnHeadersHeight = DpiHelper.Scale(32);
            grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
        }

        /// <summary>
        /// Apply modern theme to a TextBox
        /// </summary>
        public static void ApplyTo(TextBox tb)
        {
            tb.BackColor = Surface;
            tb.ForeColor = TextPrimary;
            tb.Font = RegularFont;
            tb.BorderStyle = BorderStyle.FixedSingle;
        }

        /// <summary>
        /// Apply modern theme to a MenuStrip
        /// </summary>
        public static void ApplyTo(MenuStrip menu)
        {
            menu.BackColor = Surface;
            menu.ForeColor = TextPrimary;
            menu.Font = RegularFont;
            menu.Renderer = new ModernMenuRenderer();
        }

        /// <summary>
        /// Apply modern theme to a ToolStrip
        /// </summary>
        public static void ApplyTo(ToolStrip strip)
        {
            strip.BackColor = Surface;
            strip.ForeColor = TextPrimary;
            strip.Font = RegularFont;
            strip.GripStyle = ToolStripGripStyle.Hidden;
            strip.Renderer = new ModernToolStripRenderer();
        }

        /// <summary>
        /// Apply modern theme to a StatusStrip
        /// </summary>
        public static void ApplyTo(StatusStrip status)
        {
            status.BackColor = Surface;
            status.ForeColor = TextSecondary;
            status.Font = RegularFont;
            status.Renderer = new ModernToolStripRenderer();
            status.SizingGrip = false;
        }

        /// <summary>
        /// Apply modern theme to a TabControl
        /// </summary>
        public static void ApplyTo(TabControl tabs)
        {
            tabs.DrawMode = TabDrawMode.OwnerDrawFixed;
            
            // Only add handler if this TabControl hasn't been themed yet
            if (!_themedTabControls.TryGetValue(tabs, out _))
            {
                _themedTabControls.Add(tabs, new object());
                tabs.DrawItem += TabControl_DrawItem;
            }
            
            tabs.Padding = new Point(12, 6);
            
            foreach (TabPage page in tabs.TabPages)
            {
                page.BackColor = Background;
                page.ForeColor = TextPrimary;
            }
        }

        private static void TabControl_DrawItem(object? sender, DrawItemEventArgs e)
        {
            var tabs = sender as TabControl;
            if (tabs == null) return;
            
            var page = tabs.TabPages[e.Index];
            var bounds = tabs.GetTabRect(e.Index);
            
            bool isSelected = tabs.SelectedIndex == e.Index;
            Color backColor = isSelected ? Accent : Surface;
            Color textColor = isSelected ? Color.White : TextSecondary;
            
            using var brush = new SolidBrush(backColor);
            e.Graphics.FillRectangle(brush, bounds);
            
            using var textBrush = new SolidBrush(textColor);
            using var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            e.Graphics.DrawString(page.Text, isSelected ? BoldFont : RegularFont, textBrush, bounds, sf);
        }

        /// <summary>
        /// Apply theme to a Panel used as a section header
        /// </summary>
        public static Panel CreateSectionHeader(string text)
        {
            var panel = new Panel
            {
                Height = DpiHelper.Scale(32),
                Dock = DockStyle.Top,
                BackColor = Surface,
                Padding = new Padding(10, 0, 0, 0)
            };
            
            var label = new Label
            {
                Text = text,
                ForeColor = TextSecondary,
                Font = BoldFont,
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };
            
            panel.Controls.Add(label);
            return panel;
        }

        /// <summary>
        /// Create a modern styled button
        /// </summary>
        public static Button CreateButton(string text, EventHandler? onClick = null)
        {
            var btn = new Button
            {
                Text = text,
                FlatStyle = FlatStyle.Flat,
                BackColor = Accent,
                ForeColor = Color.White,
                Font = BoldFont,
                Cursor = Cursors.Hand,
                Height = DpiHelper.Scale(36),
                Padding = new Padding(16, 0, 16, 0)
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.FlatAppearance.MouseOverBackColor = AccentHover;
            btn.FlatAppearance.MouseDownBackColor = AccentDark;
            
            if (onClick != null)
                btn.Click += onClick;
            
            return btn;
        }

        /// <summary>
        /// Create a secondary styled button (outlined style)
        /// </summary>
        public static Button CreateSecondaryButton(string text, EventHandler? onClick = null)
        {
            var btn = new Button
            {
                Text = text,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.Transparent,
                ForeColor = Accent,
                Font = BoldFont,
                Cursor = Cursors.Hand,
                Height = DpiHelper.Scale(36),
                Padding = new Padding(16, 0, 16, 0)
            };
            btn.FlatAppearance.BorderColor = Accent;
            btn.FlatAppearance.BorderSize = 1;
            btn.FlatAppearance.MouseOverBackColor = Selection;
            btn.FlatAppearance.MouseDownBackColor = SurfaceLight;
            
            if (onClick != null)
                btn.Click += onClick;
            
            return btn;
        }

        /// <summary>
        /// Create folder icon for tree view
        /// </summary>
        public static Bitmap CreateFolderIcon(bool isOpen = false)
        {
            var bmp = new Bitmap(18, 18);
            using var g = Graphics.FromImage(bmp);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            
            var color = isOpen ? Color.FromArgb(255, 213, 79) : Color.FromArgb(255, 193, 7);
            using var brush = new SolidBrush(color);
            
            // Folder shape
            var points = new Point[]
            {
                new Point(1, 5),
                new Point(7, 5),
                new Point(9, 3),
                new Point(16, 3),
                new Point(16, 15),
                new Point(1, 15)
            };
            g.FillPolygon(brush, points);
            
            // Slight 3D effect
            using var darkBrush = new SolidBrush(Color.FromArgb(50, 0, 0, 0));
            g.FillRectangle(darkBrush, 1, 13, 15, 2);
            
            return bmp;
        }

        /// <summary>
        /// Create value icon for tree view
        /// </summary>
        public static Bitmap CreateValueIcon()
        {
            var bmp = new Bitmap(18, 18);
            using var g = Graphics.FromImage(bmp);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            
            using var brush = new SolidBrush(Accent);
            g.FillRectangle(brush, 3, 3, 12, 12);
            
            using var innerBrush = new SolidBrush(Color.White);
            g.FillRectangle(innerBrush, 5, 7, 8, 2);
            g.FillRectangle(innerBrush, 8, 5, 2, 6);
            
            return bmp;
        }
    }

    /// <summary>
    /// Custom renderer for modern menu appearance
    /// </summary>
    public class ModernMenuRenderer : ToolStripProfessionalRenderer
    {
        public ModernMenuRenderer() : base(new ModernColorTable()) { }

        protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e)
        {
            if (e.Item.Selected || e.Item.Pressed)
            {
                using var brush = new SolidBrush(ModernTheme.Selection);
                e.Graphics.FillRectangle(brush, new Rectangle(Point.Empty, e.Item.Size));
            }
        }

        protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e)
        {
            e.TextColor = e.Item.Enabled ? ModernTheme.TextPrimary : ModernTheme.TextDisabled;
            base.OnRenderItemText(e);
        }

        protected override void OnRenderSeparator(ToolStripSeparatorRenderEventArgs e)
        {
            using var pen = new Pen(ModernTheme.Border);
            int y = e.Item.Height / 2;
            e.Graphics.DrawLine(pen, 0, y, e.Item.Width, y);
        }

        protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e)
        {
            using var pen = new Pen(ModernTheme.Border);
            e.Graphics.DrawRectangle(pen, 0, 0, e.ToolStrip.Width - 1, e.ToolStrip.Height - 1);
        }
    }

    /// <summary>
    /// Custom renderer for modern toolbar appearance
    /// </summary>
    public class ModernToolStripRenderer : ToolStripProfessionalRenderer
    {
        public ModernToolStripRenderer() : base(new ModernColorTable()) { }

        protected override void OnRenderButtonBackground(ToolStripItemRenderEventArgs e)
        {
            if (e.Item.Selected || e.Item.Pressed)
            {
                using var brush = new SolidBrush(e.Item.Pressed ? ModernTheme.AccentDark : ModernTheme.Selection);
                var rect = new Rectangle(2, 2, e.Item.Width - 4, e.Item.Height - 4);
                e.Graphics.FillRectangle(brush, rect);
            }
        }

        protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e)
        {
            e.TextColor = e.Item.Enabled ? ModernTheme.TextPrimary : ModernTheme.TextDisabled;
            base.OnRenderItemText(e);
        }

        protected override void OnRenderToolStripBackground(ToolStripRenderEventArgs e)
        {
            using var brush = new SolidBrush(ModernTheme.Surface);
            e.Graphics.FillRectangle(brush, e.AffectedBounds);
        }

        protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e)
        {
            // No border for cleaner look
        }

        protected override void OnRenderSeparator(ToolStripSeparatorRenderEventArgs e)
        {
            using var pen = new Pen(ModernTheme.Border);
            int x = e.Item.Width / 2;
            e.Graphics.DrawLine(pen, x, 4, x, e.Item.Height - 4);
        }
    }

    /// <summary>
    /// Color table for modern theme
    /// </summary>
    public class ModernColorTable : ProfessionalColorTable
    {
        public override Color MenuBorder => ModernTheme.Border;
        public override Color MenuItemBorder => Color.Transparent;
        public override Color MenuItemSelected => ModernTheme.Selection;
        public override Color MenuItemSelectedGradientBegin => ModernTheme.Selection;
        public override Color MenuItemSelectedGradientEnd => ModernTheme.Selection;
        public override Color MenuItemPressedGradientBegin => ModernTheme.AccentDark;
        public override Color MenuItemPressedGradientEnd => ModernTheme.AccentDark;
        public override Color MenuStripGradientBegin => ModernTheme.Surface;
        public override Color MenuStripGradientEnd => ModernTheme.Surface;
        public override Color ToolStripDropDownBackground => ModernTheme.Surface;
        public override Color ImageMarginGradientBegin => ModernTheme.Surface;
        public override Color ImageMarginGradientMiddle => ModernTheme.Surface;
        public override Color ImageMarginGradientEnd => ModernTheme.Surface;
        public override Color SeparatorDark => ModernTheme.Border;
        public override Color SeparatorLight => ModernTheme.Border;
    }
}

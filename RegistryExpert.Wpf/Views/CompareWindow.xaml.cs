using System.Windows;
using System.Windows.Controls;
using RegistryExpert.Wpf.Helpers;
using RegistryExpert.Wpf.ViewModels;

namespace RegistryExpert.Wpf.Views
{
    public partial class CompareWindow : Window
    {
        private CompareViewModel? _vm;
        private bool _isSyncingValues;

        public CompareWindow(IReadOnlyList<LoadedHiveInfo> loadedHives)
        {
            InitializeComponent();
            _vm = new CompareViewModel(loadedHives);
            DataContext = _vm;

            // Subscribe to tree expand/collapse routed events
            LeftTree.AddHandler(TreeViewItem.ExpandedEvent,
                new RoutedEventHandler(OnLeftTreeItemExpanded));
            LeftTree.AddHandler(TreeViewItem.CollapsedEvent,
                new RoutedEventHandler(OnLeftTreeItemCollapsed));
            RightTree.AddHandler(TreeViewItem.ExpandedEvent,
                new RoutedEventHandler(OnRightTreeItemExpanded));
            RightTree.AddHandler(TreeViewItem.CollapsedEvent,
                new RoutedEventHandler(OnRightTreeItemCollapsed));

            // Set up available hives dropdown on load buttons
            if (_vm.HasAvailableHives)
                SetupAvailableHivesDropdown();

            ThemeManager.ThemeChanged += OnThemeChanged;
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            ThemeManager.ApplyWindowChrome(this);
        }

        protected override void OnClosed(EventArgs e)
        {
            ThemeManager.ThemeChanged -= OnThemeChanged;
            _vm?.Dispose();
            base.OnClosed(e);
        }

        private void OnThemeChanged(object? sender, EventArgs e)
        {
            ThemeManager.ApplyWindowChrome(this);
        }

        // ── Available Hives Dropdown ──

        private void SetupAvailableHivesDropdown()
        {
            if (_vm == null) return;

            // LEFT load button: dropdown with available hives + browse option
            var leftMenu = new ContextMenu { PlacementTarget = LeftLoadButton, Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom };
            foreach (var hive in _vm.AvailableHives)
            {
                var item = new MenuItem
                {
                    Header = hive.DisplayName,
                    Command = _vm.LoadLeftFromAvailableCommand,
                    CommandParameter = hive
                };
                leftMenu.Items.Add(item);
            }
            leftMenu.Items.Add(new Separator());
            leftMenu.Items.Add(new MenuItem
            {
                Header = "Browse File...",
                Command = _vm.LoadLeftHiveCommand
            });

            LeftLoadButton.ContextMenu = leftMenu;
            LeftLoadButton.Click += (s, e) =>
            {
                e.Handled = true;
                LeftLoadButton.ContextMenu.IsOpen = true;
            };
            LeftLoadButton.Command = null;
            LeftLoadButton.Content = "Load Hive File \u25BE";

            // RIGHT load button: always opens file browser (no dropdown)
            // since the second hive must match the first hive's type,
            // available hives from main window aren't useful here.
        }

        // ── Tree Selection ──

        private void LeftTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (e.NewValue is CompareTreeNode node)
                _vm!.SelectedLeftNode = node;
        }

        private void RightTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (e.NewValue is CompareTreeNode node)
                _vm!.SelectedRightNode = node;
        }

        // ── Tree Lazy Loading + Sync ──

        private void OnLeftTreeItemExpanded(object sender, RoutedEventArgs e)
        {
            if (e.OriginalSource is TreeViewItem tvi && tvi.DataContext is CompareTreeNode node)
            {
                // Lazy load: if first child is dummy, populate real children
                if (node.Children.Count == 1 && node.Children[0].IsDummy)
                    _vm?.PopulateChildren(node, isLeft: true);

                // Sync expand to right tree
                _vm?.SyncExpand(node, sourceIsLeft: true);
            }
        }

        private void OnLeftTreeItemCollapsed(object sender, RoutedEventArgs e)
        {
            if (e.OriginalSource is TreeViewItem tvi && tvi.DataContext is CompareTreeNode node)
                _vm?.SyncCollapse(node, sourceIsLeft: true);
        }

        private void OnRightTreeItemExpanded(object sender, RoutedEventArgs e)
        {
            if (e.OriginalSource is TreeViewItem tvi && tvi.DataContext is CompareTreeNode node)
            {
                if (node.Children.Count == 1 && node.Children[0].IsDummy)
                    _vm?.PopulateChildren(node, isLeft: false);

                _vm?.SyncExpand(node, sourceIsLeft: false);
            }
        }

        private void OnRightTreeItemCollapsed(object sender, RoutedEventArgs e)
        {
            if (e.OriginalSource is TreeViewItem tvi && tvi.DataContext is CompareTreeNode node)
                _vm?.SyncCollapse(node, sourceIsLeft: false);
        }

        // ── Value Grid Selection Sync ──

        private void LeftValuesGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isSyncingValues || _vm == null) return;
            if (sender is DataGrid grid && grid.SelectedItem is CompareValueItem item)
            {
                _isSyncingValues = true;
                try
                {
                    int idx = _vm.FindMatchingValueIndex(item.Name, sourceIsLeft: true);
                    if (idx >= 0 && idx < RightValuesGrid.Items.Count)
                    {
                        RightValuesGrid.SelectedIndex = idx;
                        RightValuesGrid.ScrollIntoView(RightValuesGrid.Items[idx]);
                    }
                    else
                    {
                        RightValuesGrid.SelectedIndex = -1;
                    }
                }
                finally
                {
                    _isSyncingValues = false;
                }
            }
        }

        private void RightValuesGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isSyncingValues || _vm == null) return;
            if (sender is DataGrid grid && grid.SelectedItem is CompareValueItem item)
            {
                _isSyncingValues = true;
                try
                {
                    int idx = _vm.FindMatchingValueIndex(item.Name, sourceIsLeft: false);
                    if (idx >= 0 && idx < LeftValuesGrid.Items.Count)
                    {
                        LeftValuesGrid.SelectedIndex = idx;
                        LeftValuesGrid.ScrollIntoView(LeftValuesGrid.Items[idx]);
                    }
                    else
                    {
                        LeftValuesGrid.SelectedIndex = -1;
                    }
                }
                finally
                {
                    _isSyncingValues = false;
                }
            }
        }

        private void HelpButton_Click(object sender, RoutedEventArgs e)
        {
            HelpPopup.IsOpen = !HelpPopup.IsOpen;
        }
    }
}

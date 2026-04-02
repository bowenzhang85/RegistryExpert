using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using RegistryExpert.Wpf.Helpers;
using RegistryExpert.Wpf.ViewModels;

namespace RegistryExpert.Wpf.Views
{
    public partial class StatisticsWindow : Window
    {
        private StatisticsViewModel? _vm;

        public StatisticsWindow(IReadOnlyList<LoadedHiveInfo> loadedHives)
        {
            InitializeComponent();
            _vm = new StatisticsViewModel(loadedHives);
            DataContext = _vm;

            // Subscribe to TreeViewItem.Expanded routed event on all trees
            KeyCountTree.AddHandler(TreeViewItem.ExpandedEvent,
                new RoutedEventHandler(OnTreeItemExpanded));
            ValueCountTree.AddHandler(TreeViewItem.ExpandedEvent,
                new RoutedEventHandler(OnTreeItemExpanded));
            DataSizeTree.AddHandler(TreeViewItem.ExpandedEvent,
                new RoutedEventHandler(OnTreeItemExpanded));

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
            base.OnClosed(e);
        }

        private void OnThemeChanged(object? sender, EventArgs e)
        {
            ThemeManager.ApplyWindowChrome(this);
        }

        /// <summary>
        /// Handle TreeViewItem expand events to process "... and N more" pagination nodes.
        /// When a "more" node is expanded, we replace it with the next batch of sibling items
        /// at the same level (not as children).
        /// </summary>
        private void OnTreeItemExpanded(object sender, RoutedEventArgs e)
        {
            if (_vm == null) return;
            if (e.OriginalSource is not TreeViewItem treeViewItem) return;
            if (treeViewItem.DataContext is not StatisticsViewModel.KeyStatNode node) return;
            if (!node.IsMoreNode) return;

            // Find the parent collection this "more" node belongs to
            var parentCollection = FindParentCollection(node, (TreeView)sender);
            if (parentCollection == null) return;

            // Collapse the node first so the TreeView doesn't try to show children
            node.IsExpanded = false;

            // Dispatch to the correct expand method based on node type
            if (node.RemainingValueItems != null)
                _vm.ExpandMoreValueNode(node, parentCollection);
            else if (node.RemainingItems != null)
                _vm.ExpandMoreNode(node, parentCollection);
        }

        /// <summary>
        /// Walk up the visual/data tree to find the ObservableCollection that contains this node.
        /// </summary>
        private ObservableCollection<StatisticsViewModel.KeyStatNode>? FindParentCollection(
            StatisticsViewModel.KeyStatNode moreNode, TreeView tree)
        {
            if (_vm == null) return null;

            // Check if it's a root-level node
            ObservableCollection<StatisticsViewModel.KeyStatNode> rootCollection;
            if (ReferenceEquals(tree, KeyCountTree))
                rootCollection = _vm.KeyCountNodes;
            else if (ReferenceEquals(tree, ValueCountTree))
                rootCollection = _vm.ValueCountNodes;
            else
                rootCollection = _vm.DataSizeNodes;

            if (rootCollection.Contains(moreNode))
                return rootCollection;

            // Search recursively through the tree to find the parent node
            return FindCollectionContaining(rootCollection, moreNode);
        }

        private static ObservableCollection<StatisticsViewModel.KeyStatNode>? FindCollectionContaining(
            ObservableCollection<StatisticsViewModel.KeyStatNode> collection,
            StatisticsViewModel.KeyStatNode target)
        {
            foreach (var node in collection)
            {
                if (node.Children.Contains(target))
                    return node.Children;

                var found = FindCollectionContaining(node.Children, target);
                if (found != null) return found;
            }
            return null;
        }
    }
}

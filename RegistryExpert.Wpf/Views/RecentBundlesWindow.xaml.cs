using System.Windows;
using System.Windows.Input;
using RegistryExpert.Core.Models;
using RegistryExpert.Wpf.Helpers;

namespace RegistryExpert.Wpf.Views
{
    public partial class RecentBundlesWindow : Window
    {
        private readonly List<BundleInfo> _bundles;

        public RecentBundlesWindow(List<BundleInfo> bundles)
        {
            InitializeComponent();
            _bundles = bundles;
            BundleList.ItemsSource = _bundles;

            // Pre-select the first (most recent) bundle
            if (_bundles.Count > 0)
                BundleList.SelectedIndex = 0;
        }

        /// <summary>
        /// The bundle the user selected, or null if none.
        /// </summary>
        public BundleInfo? SelectedBundle =>
            BundleList.SelectedItem as BundleInfo;

        /// <summary>
        /// True if the user clicked "Browse..." to open the standard folder picker.
        /// </summary>
        public bool BrowseRequested { get; private set; }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            ThemeManager.ApplyWindowChrome(this);
            ThemeManager.ThemeChanged += OnThemeChanged;
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

        private void LoadButton_Click(object sender, RoutedEventArgs e)
        {
            if (BundleList.SelectedItem == null)
                return;
            DialogResult = true;
        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            BrowseRequested = true;
            DialogResult = false;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private void BundleList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (BundleList.SelectedItem != null)
                DialogResult = true;
        }
    }
}

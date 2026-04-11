using System.Windows;
using RegistryExpert.Wpf.Helpers;
using RegistryExpert.Wpf.ViewModels;

namespace RegistryExpert.Wpf.Views
{
    public partial class HivePickerWindow : Window
    {
        private readonly List<HivePickerItem> _items;

        public HivePickerWindow(List<HivePickerItem> items)
        {
            InitializeComponent();
            _items = items;
            HiveList.ItemsSource = _items;
            HeaderText.Text = $"Found {_items.Count} registry hive(s) in folder:";
        }

        /// <summary>
        /// Returns only the items the user checked.
        /// </summary>
        public List<HivePickerItem> SelectedItems =>
            _items.Where(i => i.IsSelected).ToList();

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

        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in _items)
                item.IsSelected = true;
        }

        private void SelectNone_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in _items)
                item.IsSelected = false;
        }

        private void LoadButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}

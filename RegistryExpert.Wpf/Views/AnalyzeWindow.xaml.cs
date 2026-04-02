using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using RegistryExpert.Wpf.Helpers;
using RegistryExpert.Wpf.ViewModels;

namespace RegistryExpert.Wpf.Views
{
    public partial class AnalyzeWindow : Window
    {
        private AnalyzeViewModel? _vm;

        public AnalyzeWindow(IReadOnlyList<LoadedHiveInfo> loadedHives)
        {
            InitializeComponent();
            _vm = new AnalyzeViewModel(loadedHives);
            DataContext = _vm;

            // Sync initial column widths from ViewModel
            UpdateDefaultGridColumnWidths();

            // Listen for ViewModel property changes to update column widths
            _vm.PropertyChanged += OnViewModelPropertyChanged;

            ThemeManager.ThemeChanged += OnThemeChanged;
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            ThemeManager.ApplyWindowChrome(this);
        }

        protected override void OnClosed(EventArgs e)
        {
            if (_vm != null)
                _vm.PropertyChanged -= OnViewModelPropertyChanged;
            ThemeManager.ThemeChanged -= OnThemeChanged;
            base.OnClosed(e);
        }

        private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName is nameof(AnalyzeViewModel.GridColumn1Star)
                               or nameof(AnalyzeViewModel.GridColumn2Star)
                               or nameof(AnalyzeViewModel.GridColumn3Star)
                               or nameof(AnalyzeViewModel.GridColumn4Star))
            {
                UpdateDefaultGridColumnWidths();
            }
        }

        private void UpdateDefaultGridColumnWidths()
        {
            if (_vm == null || DefaultGrid.Columns.Count < 2) return;
            DefaultGrid.Columns[0].Width = new DataGridLength(_vm.GridColumn1Star, DataGridLengthUnitType.Star);
            DefaultGrid.Columns[1].Width = new DataGridLength(_vm.GridColumn2Star, DataGridLengthUnitType.Star);
            if (DefaultGrid.Columns.Count > 2)
                DefaultGrid.Columns[2].Width = new DataGridLength(_vm.GridColumn3Star, DataGridLengthUnitType.Star);
            if (DefaultGrid.Columns.Count > 3)
                DefaultGrid.Columns[3].Width = new DataGridLength(_vm.GridColumn4Star, DataGridLengthUnitType.Star);
        }

        private void OnThemeChanged(object? sender, EventArgs e)
        {
            ThemeManager.ApplyWindowChrome(this);
        }

        private void DeviceManagerTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (DataContext is AnalyzeViewModel vm && e.NewValue is AnalyzeViewModel.DeviceTreeNode node)
            {
                vm.SelectedDeviceNode = node;
            }
        }

        private void ScheduledTasksTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (DataContext is AnalyzeViewModel vm && e.NewValue is AnalyzeViewModel.ScheduledTaskTreeNode node)
            {
                vm.SelectedScheduledTaskNode = node;
            }
        }

        private void CertStoresTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (DataContext is AnalyzeViewModel vm && e.NewValue is AnalyzeViewModel.ScheduledTaskTreeNode node)
            {
                vm.SelectedCertStoreNode = node;
            }
        }

        private void RolesTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (DataContext is AnalyzeViewModel vm && e.NewValue is AnalyzeViewModel.DeviceTreeNode node)
            {
                vm.SelectedRolesFeatureNode = node;
            }
        }
    }
}

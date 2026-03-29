using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;
using RegistryExpert.Wpf.Helpers;
using RegistryExpert.Wpf.ViewModels;

namespace RegistryExpert.Wpf.Views
{
    public partial class TimelineWindow : Window
    {
        private TimelineViewModel? _vm;

        public TimelineWindow(IReadOnlyList<LoadedHiveInfo> loadedHives, Action<string>? navigateToKey = null)
        {
            InitializeComponent();
            _vm = new TimelineViewModel(loadedHives, navigateToKey);
            DataContext = _vm;
            ThemeManager.ThemeChanged += OnThemeChanged;
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            ThemeManager.ApplyWindowChrome(this);
        }

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            if (_vm != null)
                await _vm.AutoScanAsync();
        }

        private void Window_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            ThemeManager.ThemeChanged -= OnThemeChanged;
            _vm?.Cleanup();
        }

        private void OnThemeChanged(object? sender, EventArgs e)
        {
            ThemeManager.ApplyWindowChrome(this);
        }

        private void TimelineGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            _vm?.NavigateCommand.Execute(null);
        }
    }
}

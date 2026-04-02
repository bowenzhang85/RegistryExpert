using System.Diagnostics;
using System.Reflection;
using System.Windows;
using System.Windows.Documents;
using RegistryExpert.Wpf.Helpers;

namespace RegistryExpert.Wpf.Views
{
    public partial class AboutWindow : Window
    {
        public AboutWindow()
        {
            InitializeComponent();

            var version = Assembly.GetEntryAssembly()?.GetName().Version;
            VersionLabel.Text = version != null
                ? $"Version {version.Major}.{version.Minor}.{version.Build}"
                : "Version unknown";
        }

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

        private void EmailLink_Click(object sender, RoutedEventArgs e)
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "mailto:bowenzhang@microsoft.com?subject=Registry Expert Feedback",
                UseShellExecute = true
            });
        }

        private void SubmitFeedback_Click(object sender, RoutedEventArgs e)
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "https://forms.office.com/r/9JAtABC2Ki",
                UseShellExecute = true
            });
        }

        private void PrivacyLink_Click(object sender, RoutedEventArgs e)
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "https://www.microsoft.com/en-us/privacy/data-privacy-notice",
                UseShellExecute = true
            });
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}

using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Documents;
using RegistryExpert.Wpf.Helpers;

namespace RegistryExpert.Wpf.Views
{
    /// <summary>
    /// Polished in-app release notes viewer. Renders markdown (a subset) into a themed
    /// FlowDocument. Reused by Help &gt; What's New, the update dialogs, and the
    /// post-update banner.
    /// </summary>
    public partial class ReleaseNotesWindow : Window
    {
        private readonly string _version;
        private readonly DateTimeOffset? _publishedAt;
        private readonly string _markdownBody;
        private readonly string? _githubUrl;

        public ReleaseNotesWindow(string version, DateTimeOffset? publishedAt, string markdownBody, string? githubUrl)
        {
            InitializeComponent();

            _version = version ?? "";
            _publishedAt = publishedAt;
            _markdownBody = markdownBody ?? "";
            _githubUrl = githubUrl;

            UpdateHeaderSubtitle();
            RenderBody();

            // Disable GitHub buttons when we have no URL to navigate to
            if (string.IsNullOrWhiteSpace(_githubUrl))
            {
                GitHubButton.Visibility = Visibility.Collapsed;
                EmptyStateGitHubLink.IsEnabled = false;
            }
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
            // Re-render so the FlowDocument picks up new theme brushes
            RenderBody();
        }

        private void UpdateHeaderSubtitle()
        {
            var versionLabel = string.IsNullOrWhiteSpace(_version)
                ? "Registry Expert"
                : $"Registry Expert {_version}";

            if (_publishedAt.HasValue)
            {
                var local = _publishedAt.Value.ToLocalTime();
                HeaderSubtitle.Text = $"{versionLabel}  ·  Released {local:MMMM d, yyyy}";
            }
            else
            {
                HeaderSubtitle.Text = versionLabel;
            }
        }

        private void RenderBody()
        {
            if (string.IsNullOrWhiteSpace(_markdownBody))
            {
                NotesViewer.Visibility = Visibility.Collapsed;
                EmptyState.Visibility = Visibility.Visible;
                return;
            }

            NotesViewer.Visibility = Visibility.Visible;
            EmptyState.Visibility = Visibility.Collapsed;
            NotesViewer.Document = MarkdownRenderer.Render(_markdownBody, this);
        }

        private void ViewOnGitHubButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_githubUrl)) return;
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = _githubUrl,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to open release URL: {ex.Message}");
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e) => Close();
    }
}

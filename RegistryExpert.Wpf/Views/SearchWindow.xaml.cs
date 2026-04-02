using System.ComponentModel;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using RegistryExpert.Wpf.Helpers;
using RegistryExpert.Wpf.ViewModels;

namespace RegistryExpert.Wpf.Views
{
    public partial class SearchWindow : Window
    {
        private SearchViewModel ViewModel => (SearchViewModel)DataContext;

        public SearchWindow(MainViewModel mainViewModel)
        {
            InitializeComponent();
            DataContext = new SearchViewModel(mainViewModel);

            // Focus the search box on load
            Loaded += (s, e) => SearchBox.Focus();

            // Subscribe to property changes for preview highlighting
            ViewModel.PropertyChanged += ViewModel_PropertyChanged;
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
            ViewModel.PropertyChanged -= ViewModel_PropertyChanged;
            ViewModel.Dispose();
            base.OnClosed(e);
        }

        private void OnThemeChanged(object? sender, EventArgs e)
        {
            ThemeManager.ApplyWindowChrome(this);
            // Re-render preview highlight after theme change
            UpdatePreviewHighlight();
            // Refresh status bar brush from new theme
            ViewModel.RefreshStatusBrush();
        }

        private void ViewModel_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(SearchViewModel.PreviewText))
            {
                UpdatePreviewHighlight();
            }
        }

        private void UpdatePreviewHighlight()
        {
            var text = ViewModel.PreviewText;
            var searchTerm = ViewModel.CurrentSearchTerm;
            var foreground = (Brush)FindResource("TextPrimaryBrush");

            var doc = new FlowDocument();
            var paragraph = new Paragraph();

            if (string.IsNullOrEmpty(text))
            {
                PreviewRichTextBox.Document = doc;
                return;
            }

            if (string.IsNullOrEmpty(searchTerm))
            {
                paragraph.Inlines.Add(new Run(text) { Foreground = foreground });
                doc.Blocks.Add(paragraph);
                PreviewRichTextBox.Document = doc;
                return;
            }

            // Find all occurrences (case-insensitive) and build Runs with highlighting
            int pos = 0;
            while (pos < text.Length)
            {
                int matchIndex = text.IndexOf(searchTerm, pos, StringComparison.OrdinalIgnoreCase);
                if (matchIndex < 0)
                {
                    // No more matches — add remaining text
                    paragraph.Inlines.Add(new Run(text[pos..]) { Foreground = foreground });
                    break;
                }

                // Add text before the match
                if (matchIndex > pos)
                {
                    paragraph.Inlines.Add(new Run(text[pos..matchIndex]) { Foreground = foreground });
                }

                // Add the highlighted match (use the original casing from the text)
                paragraph.Inlines.Add(new Run(text[matchIndex..(matchIndex + searchTerm.Length)])
                {
                    Background = (Brush)FindResource("SearchHighlightBackgroundBrush"),
                    Foreground = (Brush)FindResource("SearchHighlightForegroundBrush")
                });

                pos = matchIndex + searchTerm.Length;
            }

            doc.Blocks.Add(paragraph);
            PreviewRichTextBox.Document = doc;
        }

        private void ResultsGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (ViewModel.SelectedResult == null) return;

            ViewModel.NavigateCommand.Execute(null);

            // Bring main window to front, then bring search window back
            Application.Current.MainWindow?.Activate();
            Activate();
        }
    }
}

using System.Windows;

namespace RegistryExpert.Wpf.Helpers
{
    /// <summary>
    /// Manages runtime theme switching by swapping ResourceDictionaries.
    /// </summary>
    public static class ThemeManager
    {
        public enum Theme { Dark, Light }

        private static Theme _currentTheme = Theme.Dark;
        public static Theme CurrentTheme => _currentTheme;

        private static readonly Uri DarkThemeUri = new("Themes/DarkTheme.xaml", UriKind.Relative);
        private static readonly Uri LightThemeUri = new("Themes/LightTheme.xaml", UriKind.Relative);

        public static event EventHandler? ThemeChanged;

        /// <summary>
        /// Switch the application theme at runtime.
        /// </summary>
        public static void SetTheme(Theme theme)
        {
            if (_currentTheme == theme) return;
            _currentTheme = theme;

            var mergedDicts = Application.Current.Resources.MergedDictionaries;

            // Remove existing theme dictionary (it's always at index 0)
            if (mergedDicts.Count > 0)
            {
                // Check if first dictionary is a theme dictionary
                var existing = mergedDicts[0];
                if (existing.Source == DarkThemeUri || existing.Source == LightThemeUri)
                {
                    mergedDicts.RemoveAt(0);
                }
            }

            // Insert new theme at position 0
            var newTheme = new ResourceDictionary
            {
                Source = theme == Theme.Dark ? DarkThemeUri : LightThemeUri
            };
            mergedDicts.Insert(0, newTheme);

            // Apply dark title bar via DWM
            foreach (Window window in Application.Current.Windows)
            {
                ApplyWindowChrome(window);
            }

            ThemeChanged?.Invoke(null, EventArgs.Empty);
        }

        /// <summary>
        /// Apply dark title bar and rounded corners on Windows 11.
        /// </summary>
        public static void ApplyWindowChrome(Window window)
        {
            if (Environment.OSVersion.Version.Build >= 22000)
            {
                var hwnd = new System.Windows.Interop.WindowInteropHelper(window).Handle;
                if (hwnd == IntPtr.Zero) return;

                int darkMode = _currentTheme == Theme.Dark ? 1 : 0;
                DwmSetWindowAttribute(hwnd, 20, ref darkMode, sizeof(int)); // DWMWA_USE_IMMERSIVE_DARK_MODE
                int cornerPref = 2; // DWMWCP_ROUND
                DwmSetWindowAttribute(hwnd, 33, ref cornerPref, sizeof(int)); // DWMWA_WINDOW_CORNER_PREFERENCE
            }
        }

        [System.Runtime.InteropServices.DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);
    }
}

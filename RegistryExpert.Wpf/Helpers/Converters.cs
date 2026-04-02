using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using RegistryExpert.Wpf.ViewModels;

namespace RegistryExpert.Wpf.Helpers
{
    public class BoolToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            => value is true ? Visibility.Visible : Visibility.Collapsed;

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => value is Visibility.Visible;
    }

    public class InverseBoolToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            => value is true ? Visibility.Collapsed : Visibility.Visible;

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => value is Visibility.Collapsed;
    }

    public class ValueImageKeyConverter : IValueConverter
    {
        private static readonly Dictionary<string, BitmapImage> _cache = new();

        public object? Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is not string key || string.IsNullOrEmpty(key))
                return null;

            // "folder" key uses the native shell folder icon
            if (key == "folder")
                return NativeIconHelper.FolderIcon;

            if (!_cache.TryGetValue(key, out var image))
            {
                var uri = new Uri($"pack://application:,,,/Assets/{key}.png", UriKind.Absolute);
                image = new BitmapImage(uri);
                image.Freeze();
                _cache[key] = image;
            }
            return image;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }

    public class BoolToFontWeightConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            => value is true ? FontWeights.Bold : FontWeights.Normal;

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }

    /// <summary>
    /// Converts a ContentMode enum to Visibility. Returns Visible when the current mode
    /// matches the ConverterParameter (a comma-separated list of mode names).
    /// </summary>
    public class ContentModeToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is not ContentMode mode || parameter is not string paramStr)
                return Visibility.Collapsed;

            // Support multiple modes: "DefaultGrid,CbsPackages"
            var modes = paramStr.Split(',');
            foreach (var m in modes)
            {
                if (Enum.TryParse<ContentMode>(m.Trim(), out var target) && mode == target)
                    return Visibility.Visible;
            }
            return Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }

    /// <summary>
    /// Converts GridColumnCount (int) to column visibility.
    /// ConverterParameter is the minimum column count needed for this column to be visible.
    /// E.g., parameter "3" means this column is visible when GridColumnCount >= 3.
    /// </summary>
    public class GridColumnCountToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is int count && parameter is string paramStr && int.TryParse(paramStr, out int minCount))
                return count >= minCount ? Visibility.Visible : Visibility.Collapsed;
            return Visibility.Visible;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }

    /// <summary>
    /// Converts a bool to a double value (for Opacity binding on category icons).
    /// </summary>
    public class BoolToDoubleConverter : IValueConverter
    {
        public double TrueValue { get; set; } = 1.0;
        public double FalseValue { get; set; } = 0.35;

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            => value is true ? TrueValue : FalseValue;

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }

    /// <summary>
    /// Converts a bool (IsEnabled) to foreground brush: enabled = TextPrimary, disabled = TextDisabled.
    /// </summary>
    public class BoolToForegroundConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is true)
                return Application.Current.FindResource("TextPrimaryBrush") as Brush ?? Brushes.White;
            return Application.Current.FindResource("TextDisabledBrush") as Brush ?? Brushes.Gray;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }

    public class DoubleToGridLengthConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            => value is double d ? new GridLength(d) : new GridLength(280);

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => value is GridLength gl ? gl.Value : 280.0;
    }
}

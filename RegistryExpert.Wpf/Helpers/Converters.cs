using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using RegistryExpert.Core.Models;
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

    /// <summary>
    /// Converts a string to Visibility: non-empty -> Visible, null/empty/whitespace -> Collapsed.
    /// Used in CompareWindow to hide the file-path subtitle row when no path is set.
    /// </summary>
    public class StringToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            => string.IsNullOrWhiteSpace(value as string) ? Visibility.Collapsed : Visibility.Visible;

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotImplementedException();
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
    /// Inverse of <see cref="ContentModeToVisibilityConverter"/>: returns Visible
    /// when the current ContentMode is NOT in the parameter list. Used to hide
    /// chrome (e.g. the global Detail pane) for views that have their own details
    /// surface (like Disk Layout with its inline Expander).
    /// </summary>
    public class InverseContentModeToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is not ContentMode mode || parameter is not string paramStr)
                return Visibility.Visible;

            var modes = paramStr.Split(',');
            foreach (var m in modes)
            {
                if (Enum.TryParse<ContentMode>(m.Trim(), out var target) && mode == target)
                    return Visibility.Collapsed;
            }
            return Visibility.Visible;
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

    /// <summary>
    /// Selects between section header and policy row templates for the GPResult-style document view.
    /// </summary>
    public class GpDocumentTemplateSelector : System.Windows.Controls.DataTemplateSelector
    {
        public DataTemplate? SectionHeaderTemplate { get; set; }
        public DataTemplate? PolicyRowTemplate { get; set; }
        public DataTemplate? ListChildRowTemplate { get; set; }

        public override DataTemplate? SelectTemplate(object item, DependencyObject container)
        {
            if (item is AnalyzeViewModel.GpDocumentRow row)
            {
                if (row.IsSectionHeader) return SectionHeaderTemplate;
                if (row.IsListChild) return ListChildRowTemplate;
                return PolicyRowTemplate;
            }
            return base.SelectTemplate(item, container);
        }
    }

    /// <summary>
    /// Renders "?" when bound to true (filesystem type is inferred), empty string when false.
    /// Used as a small visual marker beside inferred filesystem labels.
    /// </summary>
    public class BoolToInferredMarkerConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            => value is true ? "?" : "";

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }

    /// <summary>
    /// Formats a DiskLayoutPartition's capacity for the volume table.
    /// - "—" when length is null
    /// - "~123.45 GB" when LengthIsEstimated
    /// - "123.45 GB" when CapacityFromExternalSource (authoritative)
    /// </summary>
    public class PartitionCapacityFormatterConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is not DiskLayoutPartition p) return "";
            if (!p.EstimatedLengthBytes.HasValue) return "—";
            var size = DiskMgmtFormat.FormatBytes(p.EstimatedLengthBytes.Value);
            return p.LengthIsEstimated ? "~" + size : size;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }

    /// <summary>
    /// Formats a DiskLayoutPartition's offset as an LBA value (offset / 512).
    /// </summary>
    public class PartitionOffsetFormatterConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is not DiskLayoutPartition p) return "";
            return (p.PartitionOffsetBytes / 512).ToString("N0", culture);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }

    /// <summary>
    /// Formats an orphan partition's signature/identifier for compact display:
    /// "0x{HEX}" for MBR, the GPT partition GUID otherwise.
    /// </summary>
    public class OrphanSignatureFormatterConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is not DiskLayoutPartition p) return "";
            if (p.MbrDiskSignature.HasValue) return $"0x{p.MbrDiskSignature.Value:X8}";
            if (!string.IsNullOrEmpty(p.GptPartitionTypeGuid)) return p.GptPartitionTypeGuid;
            return "";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }

    /// <summary>
    /// Renders <see cref="DiskLayoutPartition.FreeSpaceBytes"/> as a human-readable
    /// size, or "—" when free-space data is not available (no diskinfo.txt enrichment).
    /// </summary>
    public class FreeSpaceConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is long bytes) return DiskMgmtFormat.FormatBytes(bytes);
            return "—";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }

    /// <summary>
    /// Computes the percentage of free space for a partition.
    /// Returns "N %" when both FreeSpaceBytes and EstimatedLengthBytes are known.
    /// Returns "—" otherwise.
    /// </summary>
    public class PercentFreeConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is not DiskLayoutPartition p) return "—";
            if (!p.FreeSpaceBytes.HasValue || !p.EstimatedLengthBytes.HasValue || p.EstimatedLengthBytes.Value <= 0)
                return "—";
            double pct = 100.0 * p.FreeSpaceBytes.Value / p.EstimatedLengthBytes.Value;
            return $"{Math.Round(pct):F0} %";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }

    /// <summary>
    /// Renders the Status column in diskmgmt.msc style:
    ///   "Healthy (Boot, System, Page File, Crash Dump, Hibernation File)"
    /// Maps our PartitionStatus + PartitionRoleFlags to diskmgmt's friendly strings.
    /// MultiBinding inputs: partition (whole), DiskLayoutDisks collection (for parent disk lookup
    /// — currently unused but reserved for future role-context lookups).
    /// </summary>
    public class DiskMgmtStatusConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values == null || values.Length == 0 || values[0] is not DiskLayoutPartition p)
                return "";

            string statusWord = p.Status switch
            {
                PartitionStatus.Online => "Healthy",
                PartitionStatus.Offline => "Offline",
                PartitionStatus.Stale => "Stale",
                _ => "Unknown",
            };

            var tags = DiskMgmtFormat.RoleTags(p);
            return tags.Count == 0 ? statusWord : $"{statusWord} ({string.Join(", ", tags)})";
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }

    /// <summary>
    /// Formats the Volume column matching diskmgmt.msc conventions:
    ///   "Windows (C:)" when label + letter known
    ///   "E:"           when only letter
    ///   "(Disk 0 partition 3)"  when neither, using the partition's 1-based position
    ///                            within its parent disk
    /// MultiBinding inputs: [0] partition, [1] DiskLayoutDisks observable collection.
    /// </summary>
    public class DiskMgmtVolumeNameConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values == null || values.Length < 1 || values[0] is not DiskLayoutPartition p)
                return "";

            // Letter + label
            if (!string.IsNullOrEmpty(p.DriveLetter))
            {
                return !string.IsNullOrEmpty(p.VolumeLabel)
                    ? $"{p.VolumeLabel} ({p.DriveLetter})"
                    : p.DriveLetter;
            }

            // Resolve partition index within parent disk
            int partitionIndex = -1;
            if (values.Length >= 2 && values[1] is System.Collections.IEnumerable disks && p.ParentDiskNumber.HasValue)
            {
                foreach (var item in disks)
                {
                    if (item is DiskLayoutDisk d && d.DiskNumber == p.ParentDiskNumber.Value)
                    {
                        for (int i = 0; i < d.Partitions.Count; i++)
                        {
                            if (ReferenceEquals(d.Partitions[i], p))
                            {
                                partitionIndex = i + 1;
                                break;
                            }
                        }
                        break;
                    }
                }
            }

            if (p.ParentDiskNumber.HasValue && partitionIndex > 0)
                return $"(Disk {p.ParentDiskNumber.Value} partition {partitionIndex})";

            // Orphan or no parent
            var guid = p.VolumeGuid;
            if (!string.IsNullOrEmpty(guid) && guid.Length > 14)
                return guid.Substring(0, 13) + "…)";
            return guid;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }

    /// <summary>
    /// Shared formatting helpers for diskmgmt-style column rendering.
    /// Public so DiskBarView and the Volume Details Expander can reuse the same mappings.
    /// </summary>
    public static class DiskMgmtFormat
    {
        public static string FormatBytes(long bytes)
        {
            if (bytes < 1024) return $"{bytes} B";
            double kb = bytes / 1024.0;
            if (kb < 1024) return $"{kb:F1} KB";
            double mb = kb / 1024;
            if (mb < 1024) return $"{mb:F0} MB";
            double gb = mb / 1024;
            if (gb < 1024) return $"{gb:F2} GB";
            double tb = gb / 1024;
            return $"{tb:F2} TB";
        }

        /// <summary>
        /// Returns the diskmgmt-friendly role tag list for a partition, in
        /// diskmgmt's display order: Boot, System, Active, Page File, Crash Dump,
        /// Hibernation File, special-type partitions, then "Basic Data Partition" /
        /// "Unallocated" defaults.
        /// </summary>
        public static System.Collections.Generic.List<string> RoleTags(DiskLayoutPartition p)
        {
            var tags = new System.Collections.Generic.List<string>();
            var r = p.Roles;

            if ((r & PartitionRoleFlags.Boot) != 0) tags.Add("Boot");
            if ((r & PartitionRoleFlags.System) != 0) tags.Add("System");
            if ((r & PartitionRoleFlags.Active) != 0) tags.Add("Active");
            if ((r & PartitionRoleFlags.Pagefile) != 0) tags.Add("Page File");
            if ((r & PartitionRoleFlags.CrashDump) != 0) tags.Add("Crash Dump");
            if ((r & PartitionRoleFlags.Hibernation) != 0) tags.Add("Hibernation File");
            if ((r & PartitionRoleFlags.ESP) != 0) tags.Add("EFI System Partition");
            if ((r & PartitionRoleFlags.MSR) != 0) tags.Add("Microsoft Reserved Partition");
            if ((r & PartitionRoleFlags.Recovery) != 0) tags.Add("Recovery Partition");
            if ((r & PartitionRoleFlags.Temp) != 0) tags.Add("Temporary Storage");

            // Default tag when no special roles applied
            if (tags.Count == 0)
            {
                if (!string.IsNullOrEmpty(p.DriveLetter))
                    tags.Add("Basic Data Partition");
                else if ((r & PartitionRoleFlags.Unmounted) != 0)
                    tags.Add("Unallocated");
            }
            return tags;
        }
    }
}

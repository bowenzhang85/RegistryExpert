using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using RegistryExpert.Core.Models;

namespace RegistryExpert.Wpf.Views
{
    /// <summary>
    /// Renders a single physical disk as a horizontal bar in diskmgmt.msc style.
    /// Left side = disk header (number, type, capacity, status).
    /// Right side = proportional partition strip with one rectangle per partition.
    /// </summary>
    public partial class DiskBarView : UserControl
    {
        public static readonly DependencyProperty DiskProperty = DependencyProperty.Register(
            nameof(Disk),
            typeof(DiskLayoutDisk),
            typeof(DiskBarView),
            new PropertyMetadata(null, OnDiskChanged));

        public DiskLayoutDisk? Disk
        {
            get => (DiskLayoutDisk?)GetValue(DiskProperty);
            set => SetValue(DiskProperty, value);
        }

        public DiskBarView()
        {
            InitializeComponent();
        }

        private static void OnDiskChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is DiskBarView view)
                view.Render(e.NewValue as DiskLayoutDisk);
        }

        private void Render(DiskLayoutDisk? disk)
        {
            PartitionStrip.ColumnDefinitions.Clear();
            PartitionStrip.Children.Clear();

            if (disk == null)
            {
                DiskNumberText.Text = "—";
                DiskTypeText.Text = "";
                DiskCapacityText.Text = "";
                DiskStatusText.Text = "";
                return;
            }

            // Header
            DiskNumberText.Text = $"Disk {disk.DiskNumber}";
            DiskTypeText.Text = disk.PartitionStyle switch
            {
                "MBR" => "Basic (MBR)",
                "GPT" => "Basic (GPT)",
                _ => "Basic",
            };

            // Capacity rollup
            long? knownTotal = ComputeKnownTotal(disk);
            DiskCapacityText.Text = knownTotal.HasValue
                ? $"≥ {FormatBytes(knownTotal.Value)} known"
                : "Unknown";

            DiskStatusText.Text = disk.Status.ToString();
            DiskStatusText.Foreground = disk.Status switch
            {
                DiskStatus.Online => (Brush)FindResource("AccentBrush"),
                DiskStatus.Offline => (Brush)FindResource("WarningTextBrush"),
                DiskStatus.Stale => (Brush)FindResource("WarningTextBrush"),
                _ => (Brush)FindResource("TextSecondaryBrush"),
            };

            // Empty-disk case
            if (disk.Partitions.Count == 0)
            {
                PartitionStrip.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                var empty = new Border
                {
                    Background = (Brush)FindResource("SurfaceBrush"),
                    BorderBrush = (Brush)FindResource("BorderBrush"),
                    BorderThickness = new Thickness(0, 0, 1, 0),
                    Child = new TextBlock
                    {
                        Text = "(no partitions registered)",
                        HorizontalAlignment = HorizontalAlignment.Center,
                        VerticalAlignment = VerticalAlignment.Center,
                        Foreground = (Brush)FindResource("TextSecondaryBrush"),
                        FontStyle = FontStyles.Italic,
                        FontSize = 11,
                    },
                };
                Grid.SetColumn(empty, 0);
                PartitionStrip.Children.Add(empty);
                return;
            }

            // Build partition rectangles with proportional widths.
            // For each partition: width = EstimatedLengthBytes when known, otherwise
            // a "fill remaining" share (handled below).
            //
            // The last partition almost always has unknown length — we give it a
            // generous default star weight so it visually represents "to end of disk".
            var partitions = disk.Partitions
                .OrderBy(p => p.PartitionOffsetBytes)
                .ToList();

            // Find the largest known partition size so we can give the "unknown size"
            // last partition a comparable weight.
            long maxKnown = partitions
                .Where(p => p.EstimatedLengthBytes.HasValue)
                .Select(p => p.EstimatedLengthBytes!.Value)
                .DefaultIfEmpty(1)
                .Max();

            for (int i = 0; i < partitions.Count; i++)
            {
                var p = partitions[i];
                double weight = p.EstimatedLengthBytes.HasValue
                    ? p.EstimatedLengthBytes.Value
                    : Math.Max(maxKnown, 1_000_000_000d); // at least 1 GB equivalent

                PartitionStrip.ColumnDefinitions.Add(new ColumnDefinition
                {
                    Width = new GridLength(weight, GridUnitType.Star),
                    // Minimum 2px so tiny partitions (KB-scale next to TB-scale)
                    // never round to 0 pixels and become invisible. They render as
                    // colored slivers. Matches diskmgmt.msc: tiny partitions remain
                    // detectable even at extreme size ratios. Detailed per-partition
                    // info is shown in the volume table above the graph.
                    MinWidth = 2,
                });

                var rect = BuildPartitionRectangle(p, disk.DiskNumber, i + 1);
                Grid.SetColumn(rect, i);
                PartitionStrip.Children.Add(rect);
            }
        }

        private FrameworkElement BuildPartitionRectangle(DiskLayoutPartition p, int diskNumber, int partitionIndex)
        {
            // Pick background brush based on role / status
            Brush bg;
            if (p.Status == PartitionStatus.Stale)
            {
                bg = (Brush)FindResource("WarningTextBrush");
            }
            else if ((p.Roles & PartitionRoleFlags.ESP) != 0 ||
                     (p.Roles & PartitionRoleFlags.MSR) != 0 ||
                     (p.Roles & PartitionRoleFlags.Recovery) != 0)
            {
                // Hatched-style differentiation (use a slightly different shade for now;
                // proper diagonal-stripe pattern comes in A.6.6)
                bg = (Brush)FindResource("BorderBrush");
            }
            else if (string.IsNullOrEmpty(p.DriveLetter))
            {
                // Unmounted — use surface (lighter neutral)
                bg = (Brush)FindResource("SurfaceBrush");
            }
            else
            {
                // Primary partition with drive letter — accent
                bg = (Brush)FindResource("AccentBrush");
            }

            var border = new Border
            {
                Background = bg,
                BorderBrush = (Brush)FindResource("BorderBrush"),
                BorderThickness = new Thickness(0, 0, 1, 0),
                Padding = new Thickness(6, 4, 6, 4),
                ToolTip = BuildTooltip(p),
            };

            var stack = new StackPanel { Orientation = Orientation.Vertical };

            // Volume label / drive letter — matches the volume table's diskmgmt-style
            // labeling: drive letter + label, then friendly role name, then
            // "(Disk N partition M)" for partitions with no other identifier
            var titleText = !string.IsNullOrEmpty(p.DriveLetter)
                ? (!string.IsNullOrEmpty(p.VolumeLabel) ? $"{p.VolumeLabel} ({p.DriveLetter})" : p.DriveLetter)
                : (p.Roles & PartitionRoleFlags.ESP) != 0 ? "EFI System"
                : (p.Roles & PartitionRoleFlags.MSR) != 0 ? "Microsoft Reserved"
                : (p.Roles & PartitionRoleFlags.Recovery) != 0 ? "Recovery"
                : $"(Disk {diskNumber} partition {partitionIndex})";
            var title = new TextBlock
            {
                Text = titleText,
                FontWeight = FontWeights.SemiBold,
                FontSize = 11,
                TextTrimming = TextTrimming.CharacterEllipsis,
                Foreground = string.IsNullOrEmpty(p.DriveLetter)
                    ? (Brush)FindResource("TextPrimaryBrush")
                    : (Brush)FindResource("AccentForegroundBrush"),
            };
            stack.Children.Add(title);

            // Capacity line — FS suffix intentionally omitted: offline registry
            // cannot reliably determine filesystem type (Phase A.6 finding).
            var capText = p.EstimatedLengthBytes.HasValue
                ? (p.LengthIsEstimated ? "~" : "") + FormatBytes(p.EstimatedLengthBytes.Value)
                : "size unknown";
            var cap = new TextBlock
            {
                Text = capText,
                FontSize = 10,
                TextTrimming = TextTrimming.CharacterEllipsis,
                Foreground = string.IsNullOrEmpty(p.DriveLetter)
                    ? (Brush)FindResource("TextSecondaryBrush")
                    : (Brush)FindResource("AccentForegroundBrush"),
            };
            stack.Children.Add(cap);

            // Status / roles line — use shared DiskMgmtFormat for friendly names
            // consistent with the volume table above
            string roleLine;
            if (p.Status == PartitionStatus.Stale)
            {
                roleLine = "Stale";
            }
            else
            {
                var tags = Helpers.DiskMgmtFormat.RoleTags(p);
                roleLine = tags.Count == 0 ? "Healthy" : "Healthy (" + string.Join(", ", tags) + ")";
            }
            var roleBlock = new TextBlock
            {
                Text = roleLine,
                FontSize = 10,
                TextTrimming = TextTrimming.CharacterEllipsis,
                Foreground = string.IsNullOrEmpty(p.DriveLetter)
                    ? (Brush)FindResource("TextSecondaryBrush")
                    : (Brush)FindResource("AccentForegroundBrush"),
            };
            stack.Children.Add(roleBlock);

            border.Child = stack;
            return border;
        }

        private static string BuildTooltip(DiskLayoutPartition p)
        {
            var sb = new System.Text.StringBuilder();
            if (!string.IsNullOrEmpty(p.DriveLetter)) sb.AppendLine($"Drive: {p.DriveLetter}");
            sb.AppendLine($"Volume GUID: {p.VolumeGuid}");
            sb.AppendLine($"Style: {p.PartitionStyle}");
            sb.AppendLine($"Offset: {p.PartitionOffsetBytes:N0} bytes (LBA {p.PartitionOffsetBytes / 512:N0})");
            if (p.EstimatedLengthBytes.HasValue)
                sb.AppendLine($"Capacity: {(p.LengthIsEstimated ? "~" : "")}{FormatBytes(p.EstimatedLengthBytes.Value)}");
            if (p.FreeSpaceBytes.HasValue)
                sb.AppendLine($"Free: {FormatBytes(p.FreeSpaceBytes.Value)} (from diskinfo.txt)");
            // Filesystem intentionally omitted from tooltip — see capacity line above.
            if (p.Roles != PartitionRoleFlags.None)
                sb.AppendLine($"Roles: {p.Roles}");
            sb.Append($"Status: {p.Status}");
            return sb.ToString();
        }

        private static long? ComputeKnownTotal(DiskLayoutDisk disk)
        {
            if (disk.Partitions.Count == 0) return null;

            // If any partition has authoritative external capacity, prefer summing
            // those that do plus the offset of the last unknown partition.
            long total = 0;
            bool anyKnown = false;
            foreach (var p in disk.Partitions)
            {
                if (p.EstimatedLengthBytes.HasValue)
                {
                    total += p.EstimatedLengthBytes.Value;
                    anyKnown = true;
                }
            }
            // Add the unallocated head if first partition's offset > 0 — typically 1 MB,
            // not material; skip for simplicity.
            return anyKnown ? total : (long?)null;
        }

        private static string FormatBytes(long bytes)
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
    }
}

using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace RegistryExpert.Wpf.Helpers;

/// <summary>
/// Extracts native Windows shell icons (e.g., folder icon) and converts them
/// to WPF BitmapSource for use in TreeView and other controls.
/// </summary>
public static class NativeIconHelper
{
    // P/Invoke constants
    private const uint SHGFI_ICON = 0x000000100;
    private const uint SHGFI_SMALLICON = 0x000000001;
    private const uint SHGFI_USEFILEATTRIBUTES = 0x000000010;
    private const uint FILE_ATTRIBUTE_DIRECTORY = 0x00000010;

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct SHFILEINFO
    {
        public IntPtr hIcon;
        public int iIcon;
        public uint dwAttributes;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string szDisplayName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
        public string szTypeName;
    }

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes,
        ref SHFILEINFO psfi, uint cbFileInfo, uint uFlags);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DestroyIcon(IntPtr hIcon);

    /// <summary>
    /// Cached folder icon. Loaded once on first access, then reused.
    /// </summary>
    private static BitmapSource? _folderIcon;

    /// <summary>
    /// Gets the native Windows Explorer folder icon as a frozen BitmapSource.
    /// Falls back to a programmatically drawn amber folder if the shell call fails.
    /// The result is cached and reused for all subsequent calls.
    /// </summary>
    public static BitmapSource FolderIcon => _folderIcon ??= LoadFolderIcon();

    private static BitmapSource LoadFolderIcon()
    {
        BitmapSource? icon = null;

        try
        {
            icon = GetNativeFolderIcon();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Failed to load native folder icon: {ex.Message}");
        }

        icon ??= CreateFallbackFolderIcon();
        icon.Freeze();
        return icon;
    }

    /// <summary>
    /// Extracts the native shell folder icon via SHGetFileInfo and converts
    /// the HICON to a WPF BitmapSource.
    /// </summary>
    private static BitmapSource? GetNativeFolderIcon()
    {
        var shfi = new SHFILEINFO();
        uint flags = SHGFI_ICON | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES;

        var result = SHGetFileInfo("folder", FILE_ATTRIBUTE_DIRECTORY,
            ref shfi, (uint)Marshal.SizeOf(typeof(SHFILEINFO)), flags);

        if (result == IntPtr.Zero || shfi.hIcon == IntPtr.Zero)
            return null;

        try
        {
            var bitmapSource = Imaging.CreateBitmapSourceFromHIcon(
                shfi.hIcon,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());
            bitmapSource.Freeze();
            return bitmapSource;
        }
        finally
        {
            DestroyIcon(shfi.hIcon);
        }
    }

    /// <summary>
    /// Creates a fallback amber-colored folder icon using WPF drawing primitives.
    /// Matches the WinForms CreateFolderIcon() fallback appearance.
    /// </summary>
    private static BitmapSource CreateFallbackFolderIcon()
    {
        const int size = 18;
        var visual = new DrawingVisual();

        using (var dc = visual.RenderOpen())
        {
            // Amber folder color (#FFC107)
            var folderBrush = new SolidColorBrush(Color.FromRgb(255, 193, 7));
            folderBrush.Freeze();

            // Folder shape (same polygon as WinForms version)
            var folderGeometry = new StreamGeometry();
            using (var ctx = folderGeometry.Open())
            {
                ctx.BeginFigure(new Point(1, 5), true, true);
                ctx.LineTo(new Point(7, 5), false, false);
                ctx.LineTo(new Point(9, 3), false, false);
                ctx.LineTo(new Point(16, 3), false, false);
                ctx.LineTo(new Point(16, 15), false, false);
                ctx.LineTo(new Point(1, 15), false, false);
            }
            folderGeometry.Freeze();
            dc.DrawGeometry(folderBrush, null, folderGeometry);

            // Slight 3D shadow effect at bottom
            var shadowBrush = new SolidColorBrush(Color.FromArgb(50, 0, 0, 0));
            shadowBrush.Freeze();
            dc.DrawRectangle(shadowBrush, null, new Rect(1, 13, 15, 2));
        }

        var bitmap = new RenderTargetBitmap(size, size, 96, 96, PixelFormats.Pbgra32);
        bitmap.Render(visual);
        bitmap.Freeze();
        return bitmap;
    }
}

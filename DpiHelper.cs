using System;
using System.Drawing;
using System.Windows.Forms;

namespace RegistryExpert
{
    /// <summary>
    /// Helper class for DPI-aware scaling of UI elements.
    /// Provides methods to scale pixel values based on the current DPI setting.
    /// </summary>
    public static class DpiHelper
    {
        private static float? _scaleFactor;
        private static readonly object _lock = new object();

        /// <summary>
        /// Gets the DPI scale factor (1.0 = 100%, 1.25 = 125%, 1.5 = 150%, 2.0 = 200%)
        /// </summary>
        public static float ScaleFactor
        {
            get
            {
                if (_scaleFactor == null)
                {
                    lock (_lock)
                    {
                        if (_scaleFactor == null)
                        {
                            using var g = Graphics.FromHwnd(IntPtr.Zero);
                            _scaleFactor = g.DpiX / 96f;
                        }
                    }
                }
                return _scaleFactor.Value;
            }
        }

        /// <summary>
        /// Resets the cached scale factor. Call this when DPI changes at runtime.
        /// </summary>
        public static void ResetScaleFactor()
        {
            lock (_lock)
            {
                _scaleFactor = null;
            }
        }

        /// <summary>
        /// Scales an integer value by the current DPI factor
        /// </summary>
        public static int Scale(int value) => (int)Math.Round(value * ScaleFactor);

        /// <summary>
        /// Scales a float value by the current DPI factor
        /// </summary>
        public static float Scale(float value) => value * ScaleFactor;

        /// <summary>
        /// Creates a DPI-scaled Size from the given width and height (at 96 DPI / 100%)
        /// </summary>
        public static Size ScaleSize(int width, int height) => new Size(Scale(width), Scale(height));

        /// <summary>
        /// Scales an existing Size by the current DPI factor
        /// </summary>
        public static Size ScaleSize(Size size) => new Size(Scale(size.Width), Scale(size.Height));

        /// <summary>
        /// Creates a DPI-scaled Point from the given x and y coordinates (at 96 DPI / 100%)
        /// </summary>
        public static Point ScalePoint(int x, int y) => new Point(Scale(x), Scale(y));

        /// <summary>
        /// Scales an existing Point by the current DPI factor
        /// </summary>
        public static Point ScalePoint(Point point) => new Point(Scale(point.X), Scale(point.Y));

        /// <summary>
        /// Creates a DPI-scaled Padding from the given values (at 96 DPI / 100%)
        /// </summary>
        public static Padding ScalePadding(int all) => new Padding(Scale(all));

        /// <summary>
        /// Creates a DPI-scaled Padding from the given values (at 96 DPI / 100%)
        /// </summary>
        public static Padding ScalePadding(int horizontal, int vertical) 
            => new Padding(Scale(horizontal), Scale(vertical), Scale(horizontal), Scale(vertical));

        /// <summary>
        /// Creates a DPI-scaled Padding from the given values (at 96 DPI / 100%)
        /// </summary>
        public static Padding ScalePadding(int left, int top, int right, int bottom) 
            => new Padding(Scale(left), Scale(top), Scale(right), Scale(bottom));

        /// <summary>
        /// Scales an existing Padding by the current DPI factor
        /// </summary>
        public static Padding ScalePadding(Padding padding) 
            => new Padding(Scale(padding.Left), Scale(padding.Top), Scale(padding.Right), Scale(padding.Bottom));

        /// <summary>
        /// Creates a DPI-scaled Font from the given font parameters.
        /// Note: Font sizes in Windows Forms are already DPI-aware when using Point units,
        /// so this is primarily useful when you need pixel-specific font sizing.
        /// </summary>
        public static Font ScaleFont(string familyName, float emSize, FontStyle style = FontStyle.Regular)
            => new Font(familyName, Scale(emSize), style);

        /// <summary>
        /// Determines if the system is running at high DPI (above 100%)
        /// </summary>
        public static bool IsHighDpi => ScaleFactor > 1.0f;

        /// <summary>
        /// Gets a friendly description of the current DPI scaling
        /// </summary>
        public static string ScaleDescription => $"{(int)(ScaleFactor * 100)}%";
    }
}

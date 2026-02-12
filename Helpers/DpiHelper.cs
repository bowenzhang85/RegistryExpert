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
        /// Creates a DPI-scaled Point from the given x and y coordinates (at 96 DPI / 100%)
        /// </summary>
        public static Point ScalePoint(int x, int y) => new Point(Scale(x), Scale(y));

        /// <summary>
        /// Creates a DPI-scaled Padding from the given values (at 96 DPI / 100%)
        /// </summary>
        public static Padding ScalePadding(int all) => new Padding(Scale(all));

        /// <summary>
        /// Creates a DPI-scaled Padding from the given values (at 96 DPI / 100%)
        /// </summary>
        public static Padding ScalePadding(int left, int top, int right, int bottom) 
            => new Padding(Scale(left), Scale(top), Scale(right), Scale(bottom));

    }
}

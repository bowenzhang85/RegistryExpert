using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace RegistryExpert.Wpf.Helpers
{
    /// <summary>
    /// Attached property that parses a string for known status emoji codepoints
    /// and renders the result as a sequence of colored Runs inside a TextBlock.
    ///
    /// WPF's classic TextBlock pipeline does not support COLR/CPAL color emoji
    /// fonts. We work around this by substituting emoji with text-class glyphs
    /// (e.g. U+2705 ✅ -> U+2713 ✓) and applying a themed Foreground brush.
    ///
    /// Usage:
    ///   &lt;TextBlock helpers:RichText.FormattedText="{Binding SomeString}" /&gt;
    /// </summary>
    public static class RichText
    {
        public static readonly DependencyProperty FormattedTextProperty =
            DependencyProperty.RegisterAttached(
                "FormattedText",
                typeof(string),
                typeof(RichText),
                new PropertyMetadata(string.Empty, OnFormattedTextChanged));

        public static string GetFormattedText(DependencyObject d)
            => (string)d.GetValue(FormattedTextProperty);

        public static void SetFormattedText(DependencyObject d, string value)
            => d.SetValue(FormattedTextProperty, value);

        private static void OnFormattedTextChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is not TextBlock tb) return;
            tb.Inlines.Clear();
            var text = e.NewValue as string;
            if (string.IsNullOrEmpty(text)) return;

            foreach (var run in Parse(text))
                tb.Inlines.Add(run);
        }

        /// <summary>
        /// Walks the string and emits Runs. Emoji codepoints are replaced with
        /// text-class glyphs and given a themed colored Foreground; everything
        /// else passes through as a plain Run that inherits the cell foreground.
        /// </summary>
        /// <summary>
        /// Literal text markers that should be emitted as a bold, themed Run
        /// instead of plain text. Order matters only if markers can overlap.
        /// </summary>
        private static readonly (string Marker, string BrushKey, Brush Fallback)[] TextMarkers =
        {
            ("(EXPIRED)", "ErrorBrush", Brushes.Red),
        };

        private static IEnumerable<Run> Parse(string text)
        {
            var buffer = new System.Text.StringBuilder();

            for (int i = 0; i < text.Length; i++)
            {
                // Literal-text marker match (e.g. "(EXPIRED)") takes precedence
                // over codepoint scanning at this position.
                bool matchedMarker = false;
                foreach (var (marker, brushKey, fallback) in TextMarkers)
                {
                    if (i + marker.Length <= text.Length &&
                        string.CompareOrdinal(text, i, marker, 0, marker.Length) == 0)
                    {
                        if (buffer.Length > 0)
                        {
                            yield return new Run(buffer.ToString());
                            buffer.Clear();
                        }

                        yield return new Run(marker)
                        {
                            FontWeight = FontWeights.Bold,
                            Foreground = ResolveBrush(brushKey, fallback),
                        };

                        i += marker.Length - 1; // -1 because the for-loop will i++
                        matchedMarker = true;
                        break;
                    }
                }
                if (matchedMarker) continue;

                int codepoint;
                int consumed;

                if (char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                {
                    codepoint = char.ConvertToUtf32(text[i], text[i + 1]);
                    consumed = 2;
                }
                else
                {
                    codepoint = text[i];
                    consumed = 1;
                }

                if (TryGetGlyph(codepoint, out var glyph, out var brush))
                {
                    // Flush any buffered text first
                    if (buffer.Length > 0)
                    {
                        yield return new Run(buffer.ToString());
                        buffer.Clear();
                    }

                    // Skip optional variation selector U+FE0F that often follows emoji
                    int next = i + consumed;
                    if (next < text.Length && text[next] == '\ufe0f')
                        consumed++;

                    // Skip a single trailing space if present, and re-emit it as
                    // a normal-text space so the spacing reads correctly.
                    int afterEmoji = i + consumed;
                    bool hadTrailingSpace = afterEmoji < text.Length && text[afterEmoji] == ' ';

                    var coloredRun = new Run(glyph) { FontWeight = FontWeights.Bold };
                    if (brush != null) coloredRun.Foreground = brush;
                    yield return coloredRun;

                    if (hadTrailingSpace)
                    {
                        yield return new Run(" ");
                        consumed++;
                    }

                    i += consumed - 1; // -1 because the for-loop will i++
                }
                else
                {
                    // Append both surrogate halves (or the single BMP char)
                    buffer.Append(text, i, consumed);
                    i += consumed - 1;
                }
            }

            if (buffer.Length > 0)
                yield return new Run(buffer.ToString());
        }

        /// <summary>
        /// Maps a recognized emoji codepoint to its text-class substitute and
        /// the brush that should color it. Returns false for non-status chars.
        /// </summary>
        private static bool TryGetGlyph(int codepoint, out string glyph, out Brush? brush)
        {
            switch (codepoint)
            {
                case 0x2705: // ✅ white heavy check
                    glyph = "\u2713"; // ✓
                    brush = ResolveBrush("HealthyTextBrush", Brushes.LimeGreen);
                    return true;
                case 0x26A0: // ⚠ warning sign (text-class)
                    glyph = "\u26a0";
                    brush = ResolveBrush("WarningTextBrush", Brushes.OrangeRed);
                    return true;
                case 0x2139: // ℹ information source
                    glyph = "\u24d8"; // ⓘ
                    brush = ResolveBrush("AccentBrush", Brushes.DodgerBlue);
                    return true;
                case 0x274C: // ❌ cross mark
                    glyph = "\u2717"; // ✗
                    brush = ResolveBrush("WarningTextBrush", Brushes.OrangeRed);
                    return true;
                case 0x1F534: // 🔴 large red circle
                    glyph = "\u25cf"; // ●
                    brush = ResolveBrush("WarningTextBrush", Brushes.Red);
                    return true;
                case 0x1F7E0: // 🟠 large orange circle
                    glyph = "\u25cf";
                    brush = Brushes.DarkOrange;
                    return true;
                case 0x1F7E1: // 🟡 large yellow circle
                    glyph = "\u25cf";
                    brush = Brushes.Goldenrod;
                    return true;
                default:
                    glyph = string.Empty;
                    brush = null;
                    return false;
            }
        }

        private static Brush ResolveBrush(string resourceKey, Brush fallback)
        {
            try
            {
                if (Application.Current?.TryFindResource(resourceKey) is Brush b)
                    return b;
            }
            catch { /* design-time / no app */ }
            return fallback;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace RegistryExpert.Wpf.Helpers
{
    /// <summary>
    /// Tiny markdown renderer for release notes. Supports the subset our notes use:
    /// headings (# / ## / ###), unordered/ordered lists with nesting, bold/italic,
    /// inline `code`, fenced ```code``` blocks, [text](url) links + bare URLs,
    /// and --- horizontal rules. Falls back to plain text for anything else.
    ///
    /// Output is a WPF FlowDocument suitable for FlowDocumentScrollViewer.
    /// All brushes are resolved through the supplied themedHost (DynamicResource lookup),
    /// so re-rendering after a theme change picks up the new colors.
    /// </summary>
    public static class MarkdownRenderer
    {
        public static FlowDocument Render(string markdown, FrameworkElement themedHost)
        {
            var doc = new FlowDocument
            {
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 13,
                PagePadding = new Thickness(0),
                LineHeight = 20,
                Foreground = Brush(themedHost, "TextPrimaryBrush", Brushes.Black),
                Background = Brushes.Transparent,
                IsOptimalParagraphEnabled = true,
                IsHyphenationEnabled = false
            };

            if (string.IsNullOrWhiteSpace(markdown))
                return doc;

            // Normalize line endings; GitHub may send \r\n
            var text = markdown.Replace("\r\n", "\n").Replace('\r', '\n');
            var lines = text.Split('\n');

            int i = 0;
            while (i < lines.Length)
            {
                var raw = lines[i];
                var line = raw.TrimEnd();

                // Skip blank lines (handled by paragraph margins)
                if (string.IsNullOrWhiteSpace(line))
                {
                    i++;
                    continue;
                }

                // Fenced code block ```
                if (line.TrimStart().StartsWith("```", StringComparison.Ordinal))
                {
                    var code = new StringBuilder();
                    i++; // skip opening fence
                    while (i < lines.Length && !lines[i].TrimStart().StartsWith("```", StringComparison.Ordinal))
                    {
                        code.AppendLine(lines[i]);
                        i++;
                    }
                    if (i < lines.Length) i++; // skip closing fence
                    doc.Blocks.Add(BuildCodeBlock(code.ToString().TrimEnd('\n'), themedHost));
                    continue;
                }

                // Horizontal rule
                if (IsHorizontalRule(line))
                {
                    doc.Blocks.Add(BuildHorizontalRule(themedHost));
                    i++;
                    continue;
                }

                // Headings
                if (line.StartsWith("### ", StringComparison.Ordinal))
                {
                    doc.Blocks.Add(BuildHeading(line.Substring(4), 3, themedHost));
                    i++;
                    continue;
                }
                if (line.StartsWith("## ", StringComparison.Ordinal))
                {
                    doc.Blocks.Add(BuildHeading(line.Substring(3), 2, themedHost));
                    i++;
                    continue;
                }
                if (line.StartsWith("# ", StringComparison.Ordinal))
                {
                    doc.Blocks.Add(BuildHeading(line.Substring(2), 1, themedHost));
                    i++;
                    continue;
                }

                // Lists (unordered or ordered) — collect contiguous list lines, then build one List
                if (IsListLine(line, out _, out _))
                {
                    var listLines = new List<string>();
                    while (i < lines.Length && (IsListLine(lines[i], out _, out _)
                        || (listLines.Count > 0 && IsListContinuation(lines[i]))))
                    {
                        listLines.Add(lines[i]);
                        i++;
                    }
                    doc.Blocks.Add(BuildList(listLines, themedHost));
                    continue;
                }

                // Paragraph: collect contiguous non-blank, non-special lines
                var paraLines = new List<string> { line };
                i++;
                while (i < lines.Length
                    && !string.IsNullOrWhiteSpace(lines[i])
                    && !IsBlockStart(lines[i]))
                {
                    paraLines.Add(lines[i].TrimEnd());
                    i++;
                }
                doc.Blocks.Add(BuildParagraph(string.Join(" ", paraLines), themedHost));
            }

            return doc;
        }

        // ── Block builders ────────────────────────────────────────────────

        private static Paragraph BuildHeading(string text, int level, FrameworkElement host)
        {
            double size = level switch { 1 => 22, 2 => 16, 3 => 14, _ => 13 };
            var brush = level == 1
                ? Brush(host, "AccentBrush", Brushes.SteelBlue)
                : Brush(host, "TextPrimaryBrush", Brushes.Black);

            var p = new Paragraph
            {
                FontSize = size,
                FontWeight = FontWeights.SemiBold,
                Foreground = brush,
                Margin = level switch
                {
                    1 => new Thickness(0, 4, 0, 8),
                    2 => new Thickness(0, 14, 0, 6),
                    3 => new Thickness(0, 10, 0, 4),
                    _ => new Thickness(0, 6, 0, 4)
                }
            };
            AddInlines(p.Inlines, text, host);
            return p;
        }

        private static Paragraph BuildParagraph(string text, FrameworkElement host)
        {
            var p = new Paragraph
            {
                Margin = new Thickness(0, 0, 0, 8),
                Foreground = Brush(host, "TextPrimaryBrush", Brushes.Black)
            };
            AddInlines(p.Inlines, text, host);
            return p;
        }

        private static List BuildList(List<string> lines, FrameworkElement host)
        {
            // Determine marker type from the first line
            IsListLine(lines[0], out int firstIndent, out bool firstOrdered);
            var rootList = new List
            {
                MarkerStyle = firstOrdered ? TextMarkerStyle.Decimal : TextMarkerStyle.Disc,
                Padding = new Thickness(0),
                Margin = new Thickness(20, 0, 0, 8),
                Foreground = Brush(host, "TextPrimaryBrush", Brushes.Black)
            };

            // Single-level rendering with rudimentary nesting (handle one level of nesting by indent)
            ListItem? currentItem = null;
            List? nestedList = null;

            foreach (var raw in lines)
            {
                if (IsListLine(raw, out int indent, out bool ordered))
                {
                    var text = StripListMarker(raw);
                    var para = new Paragraph
                    {
                        Margin = new Thickness(0, 0, 0, 2),
                        Foreground = Brush(host, "TextPrimaryBrush", Brushes.Black)
                    };
                    AddInlines(para.Inlines, text, host);

                    if (indent <= firstIndent)
                    {
                        // Top-level item
                        currentItem = new ListItem(para) { Margin = new Thickness(0) };
                        rootList.ListItems.Add(currentItem);
                        nestedList = null;
                    }
                    else
                    {
                        // Nested item under currentItem
                        if (currentItem == null)
                        {
                            currentItem = new ListItem(new Paragraph()) { Margin = new Thickness(0) };
                            rootList.ListItems.Add(currentItem);
                        }
                        if (nestedList == null)
                        {
                            nestedList = new List
                            {
                                MarkerStyle = ordered ? TextMarkerStyle.Decimal : TextMarkerStyle.Circle,
                                Padding = new Thickness(0),
                                Margin = new Thickness(18, 2, 0, 2),
                                Foreground = Brush(host, "TextPrimaryBrush", Brushes.Black)
                            };
                            currentItem.Blocks.Add(nestedList);
                        }
                        nestedList.ListItems.Add(new ListItem(para) { Margin = new Thickness(0) });
                    }
                }
                else if (currentItem != null)
                {
                    // Continuation line — append to last paragraph of current item
                    var continuation = raw.Trim();
                    var lastPara = currentItem.Blocks.LastBlock as Paragraph;
                    if (lastPara != null)
                    {
                        lastPara.Inlines.Add(new Run(" "));
                        AddInlines(lastPara.Inlines, continuation, host);
                    }
                }
            }
            return rootList;
        }

        private static Paragraph BuildCodeBlock(string code, FrameworkElement host)
        {
            var p = new Paragraph
            {
                FontFamily = new FontFamily("Cascadia Mono, Consolas, Courier New"),
                FontSize = 12,
                Background = Brush(host, "SurfaceBrush", Brushes.WhiteSmoke),
                Foreground = Brush(host, "TextPrimaryBrush", Brushes.Black),
                Padding = new Thickness(12, 8, 12, 8),
                Margin = new Thickness(0, 4, 0, 10),
                TextIndent = 0
            };
            p.Inlines.Add(new Run(code));
            return p;
        }

        private static BlockUIContainer BuildHorizontalRule(FrameworkElement host)
        {
            var rule = new Rectangle
            {
                Height = 1,
                Fill = Brush(host, "BorderBrush", Brushes.LightGray),
                Margin = new Thickness(0, 8, 0, 12),
                HorizontalAlignment = HorizontalAlignment.Stretch
            };
            return new BlockUIContainer(rule) { Margin = new Thickness(0) };
        }

        // ── Inline parsing ────────────────────────────────────────────────

        // Matches: `code`, **bold**, *italic*, [text](url), bare http(s) URL
        private static readonly Regex InlineRegex = new Regex(
            @"(?<code>`[^`\n]+`)" +
            @"|(?<bold>\*\*[^*\n]+\*\*)" +
            @"|(?<italic>(?<!\*)\*(?!\s)[^*\n]+?(?<!\s)\*(?!\*))" +
            @"|(?<link>\[(?<linktext>[^\]]+)\]\((?<linkurl>https?://[^\s)]+)\))" +
            @"|(?<bareurl>https?://[^\s)<>\]]+)",
            RegexOptions.Compiled);

        private static void AddInlines(InlineCollection target, string text, FrameworkElement host)
        {
            if (string.IsNullOrEmpty(text)) return;

            int pos = 0;
            foreach (Match m in InlineRegex.Matches(text))
            {
                if (m.Index > pos)
                    target.Add(new Run(text.Substring(pos, m.Index - pos)));

                if (m.Groups["code"].Success)
                {
                    var inner = m.Value.Substring(1, m.Value.Length - 2);
                    target.Add(new Run(inner)
                    {
                        FontFamily = new FontFamily("Cascadia Mono, Consolas, Courier New"),
                        FontSize = 12,
                        Background = Brush(host, "SurfaceLightBrush", Brushes.WhiteSmoke),
                        Foreground = Brush(host, "TextPrimaryBrush", Brushes.Black)
                    });
                }
                else if (m.Groups["bold"].Success)
                {
                    var inner = m.Value.Substring(2, m.Value.Length - 4);
                    target.Add(new Bold(new Run(inner)));
                }
                else if (m.Groups["italic"].Success)
                {
                    var inner = m.Value.Substring(1, m.Value.Length - 2);
                    target.Add(new Italic(new Run(inner)));
                }
                else if (m.Groups["link"].Success)
                {
                    var label = m.Groups["linktext"].Value;
                    var url = m.Groups["linkurl"].Value;
                    target.Add(BuildHyperlink(label, url, host));
                }
                else if (m.Groups["bareurl"].Success)
                {
                    var url = m.Value;
                    target.Add(BuildHyperlink(url, url, host));
                }

                pos = m.Index + m.Length;
            }

            if (pos < text.Length)
                target.Add(new Run(text.Substring(pos)));
        }

        private static Hyperlink BuildHyperlink(string text, string url, FrameworkElement host)
        {
            var link = new Hyperlink(new Run(text))
            {
                Foreground = Brush(host, "AccentBrush", Brushes.SteelBlue),
                TextDecorations = TextDecorations.Underline
            };
            try
            {
                link.NavigateUri = new Uri(url, UriKind.Absolute);
            }
            catch
            {
                // Bad URL — render as plain inline but no click handler
                return link;
            }
            link.RequestNavigate += OnNavigate;
            return link;
        }

        private static void OnNavigate(object sender, RequestNavigateEventArgs e)
        {
            try
            {
                if (e.Uri != null)
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = e.Uri.AbsoluteUri,
                        UseShellExecute = true
                    });
                }
                e.Handled = true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to open link {e.Uri}: {ex.Message}");
            }
        }

        // ── Helpers ───────────────────────────────────────────────────────

        private static bool IsListLine(string raw, out int indent, out bool ordered)
        {
            indent = 0;
            ordered = false;
            if (string.IsNullOrEmpty(raw)) return false;

            int i = 0;
            while (i < raw.Length && raw[i] == ' ') { i++; indent++; }
            if (i >= raw.Length) return false;

            // Unordered: - * +
            if ((raw[i] == '-' || raw[i] == '*' || raw[i] == '+')
                && i + 1 < raw.Length && raw[i + 1] == ' ')
            {
                ordered = false;
                return true;
            }

            // Ordered: digits then "."
            int j = i;
            while (j < raw.Length && char.IsDigit(raw[j])) j++;
            if (j > i && j < raw.Length && raw[j] == '.' && j + 1 < raw.Length && raw[j + 1] == ' ')
            {
                ordered = true;
                return true;
            }

            return false;
        }

        private static bool IsListContinuation(string raw)
        {
            // A continuation line is indented (typically 2+ spaces) and non-empty
            if (string.IsNullOrWhiteSpace(raw)) return false;
            return raw.Length > 0 && raw[0] == ' ';
        }

        private static string StripListMarker(string raw)
        {
            int i = 0;
            while (i < raw.Length && raw[i] == ' ') i++;
            if (i >= raw.Length) return "";

            if (raw[i] == '-' || raw[i] == '*' || raw[i] == '+')
            {
                // Skip marker and the single trailing space
                return raw.Substring(Math.Min(raw.Length, i + 2)).Trim();
            }

            // Ordered: skip digits + '.' + space
            int j = i;
            while (j < raw.Length && char.IsDigit(raw[j])) j++;
            if (j < raw.Length && raw[j] == '.') j++;
            if (j < raw.Length && raw[j] == ' ') j++;
            return raw.Substring(Math.Min(raw.Length, j)).Trim();
        }

        private static bool IsHorizontalRule(string line)
        {
            var s = line.Trim();
            if (s.Length < 3) return false;
            char c = s[0];
            if (c != '-' && c != '*' && c != '_') return false;
            foreach (var ch in s) if (ch != c) return false;
            return true;
        }

        private static bool IsBlockStart(string raw)
        {
            var t = raw.TrimStart();
            if (t.StartsWith("# ", StringComparison.Ordinal)) return true;
            if (t.StartsWith("## ", StringComparison.Ordinal)) return true;
            if (t.StartsWith("### ", StringComparison.Ordinal)) return true;
            if (t.StartsWith("```", StringComparison.Ordinal)) return true;
            if (IsHorizontalRule(raw)) return true;
            if (IsListLine(raw, out _, out _)) return true;
            return false;
        }

        private static Brush Brush(FrameworkElement host, string key, Brush fallback)
        {
            try
            {
                var b = host.TryFindResource(key) as Brush;
                return b ?? fallback;
            }
            catch
            {
                return fallback;
            }
        }
    }
}

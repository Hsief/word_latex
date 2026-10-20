using System;
using System.Collections.Generic;

namespace WordLatexAddin
{
    /// <summary>Describes one LaTeX formula and its character offsets inside a Word range.</summary>
    public sealed class LatexSpan
    {
        public LatexSpan(int start, int length, string latex, bool display)
        {
            Start = start;
            Length = length;
            Latex = latex;
            IsDisplay = display;
        }

        public int Start { get; private set; }
        public int Length { get; private set; }
        public int End { get { return Start + Length; } }
        public string Latex { get; private set; }
        public bool IsDisplay { get; private set; }
    }

    /// <summary>
    /// Finds $...$, \(...\), $$...$$ and \[...\] formulas without treating escaped
    /// delimiters as equation boundaries.
    /// </summary>
    public sealed class LatexScanner
    {
        public IList<LatexSpan> FindAll(string text)
        {
            List<LatexSpan> result = new List<LatexSpan>();
            if (string.IsNullOrEmpty(text)) return result;

            int position = 0;
            while (position < text.Length)
            {
                if (StartsWith(text, position, "$$") && !IsEscaped(text, position))
                {
                    int closing = FindClosing(text, position + 2, "$$", false);
                    if (closing >= 0)
                    {
                        AddSpan(result, text, position, 2, closing, 2, true);
                        position = closing + 2;
                        continue;
                    }
                    // A double-dollar opener is one token. Do not retry its second
                    // dollar as an inline opener when the display formula is incomplete.
                    position += 2;
                    continue;
                }
                else if (StartsWith(text, position, "\\[") && !IsEscaped(text, position))
                {
                    int closing = FindClosing(text, position + 2, "\\]", false);
                    if (closing >= 0)
                    {
                        AddSpan(result, text, position, 2, closing, 2, true);
                        position = closing + 2;
                        continue;
                    }
                }
                else if (StartsWith(text, position, "\\(") && !IsEscaped(text, position))
                {
                    int closing = FindClosing(text, position + 2, "\\)", true);
                    if (closing >= 0)
                    {
                        AddSpan(result, text, position, 2, closing, 2, false);
                        position = closing + 2;
                        continue;
                    }
                }
                else if (text[position] == '$' && !IsEscaped(text, position))
                {
                    int closing = FindClosing(text, position + 1, "$", true);
                    if (closing >= 0)
                    {
                        AddSpan(result, text, position, 1, closing, 1, false);
                        position = closing + 1;
                        continue;
                    }
                }

                position++;
            }
            return result;
        }

        public LatexSpan FindNearest(string text, int caretOffset)
        {
            LatexSpan nearest = null;
            foreach (LatexSpan span in FindAll(text))
            {
                if (caretOffset >= span.Start && caretOffset <= span.End) return span;
                if (span.End <= caretOffset) nearest = span;
            }
            return nearest;
        }

        public string RemoveDelimiters(string text, out bool display)
        {
            display = false;
            if (text == null) return string.Empty;
            string value = text.Trim().TrimEnd('\r', '\a').Trim();
            IList<LatexSpan> spans = FindAll(value);
            if (spans.Count == 1 && spans[0].Start == 0 && spans[0].End == value.Length)
            {
                display = spans[0].IsDisplay;
                return spans[0].Latex;
            }

            // A malformed wrapper must never be passed to the parser as ordinary
            // math text: that would turn leftover dollar signs into an OMath object.
            for (int i = 0; i < value.Length; i++)
            {
                if (IsEscaped(value, i)) continue;
                if (value[i] == '$' || StartsWith(value, i, "\\(") || StartsWith(value, i, "\\)") ||
                    StartsWith(value, i, "\\[") || StartsWith(value, i, "\\]"))
                {
                    throw new FormatException("LaTeX 公式定界符未正确配对，请选择完整的公式。");
                }
            }
            return value;
        }

        private static void AddSpan(ICollection<LatexSpan> result, string text, int start, int openLength, int closing, int closeLength, bool display)
        {
            string latex = text.Substring(start + openLength, closing - start - openLength).Trim();
            if (latex.Length > 0) result.Add(new LatexSpan(start, closing + closeLength - start, latex, display));
        }

        private static int FindClosing(string text, int start, string delimiter, bool stopAtParagraph)
        {
            for (int i = start; i <= text.Length - delimiter.Length; i++)
            {
                // Word ends a table cell with \r\a. A formula must not join two
                // cells, even when its display delimiter permits paragraph breaks.
                if (text[i] == '\a' || (stopAtParagraph && (text[i] == '\r' || text[i] == '\n'))) return -1;
                if (IsEscaped(text, i)) continue;
                if (delimiter == "$" && StartsWith(text, i, "$$"))
                {
                    // Preserve the valid adjacent-inline form $x$$y$. Otherwise a
                    // display opener is not the closing half of an earlier inline.
                    return HasSingleDollarClose(text, i + 2) ? i : -1;
                }
                if (delimiter == "$$" && text[i] == '$' && !StartsWith(text, i, "$$")) return -1;
                if (StartsWith(text, i, delimiter)) return i;
            }
            return -1;
        }

        private static bool HasSingleDollarClose(string text, int start)
        {
            for (int i = start; i < text.Length; i++)
            {
                if (text[i] == '\r' || text[i] == '\n' || text[i] == '\a') return false;
                if (text[i] != '$' || IsEscaped(text, i)) continue;
                if (!StartsWith(text, i, "$$") || StartsWith(text, i, "$$$")) return true;
                // Another close/open pair can continue an adjacent-inline chain:
                // $x$$y$$z$ still has a final single-dollar closer.
                i++;
            }
            return false;
        }

        private static bool StartsWith(string text, int position, string value)
        {
            return position >= 0 && position + value.Length <= text.Length &&
                string.CompareOrdinal(text, position, value, 0, value.Length) == 0;
        }

        private static bool IsEscaped(string text, int position)
        {
            int slashCount = 0;
            for (int i = position - 1; i >= 0 && text[i] == '\\'; i--) slashCount++;
            return slashCount % 2 != 0;
        }
    }
}

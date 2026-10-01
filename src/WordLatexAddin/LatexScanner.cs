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
            string value = text.Trim().TrimEnd('\r', '\a');
            if (value.Length >= 4 && value.StartsWith("$$", StringComparison.Ordinal) && value.EndsWith("$$", StringComparison.Ordinal))
            {
                display = true;
                return value.Substring(2, value.Length - 4).Trim();
            }
            if (value.Length >= 4 && value.StartsWith("\\[", StringComparison.Ordinal) && value.EndsWith("\\]", StringComparison.Ordinal))
            {
                display = true;
                return value.Substring(2, value.Length - 4).Trim();
            }
            if (value.Length >= 4 && value.StartsWith("\\(", StringComparison.Ordinal) && value.EndsWith("\\)", StringComparison.Ordinal))
            {
                return value.Substring(2, value.Length - 4).Trim();
            }
            if (value.Length >= 2 && value.StartsWith("$", StringComparison.Ordinal) && value.EndsWith("$", StringComparison.Ordinal))
            {
                return value.Substring(1, value.Length - 2).Trim();
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
                if (stopAtParagraph && (text[i] == '\r' || text[i] == '\n')) return -1;
                if (StartsWith(text, i, delimiter) && !IsEscaped(text, i)) return i;
            }
            return -1;
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

using System;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace WordLatexAddin
{
    /// <summary>
    /// Resolves offsets in a text snapshot to actual Word ranges. Word character
    /// positions also count content that Range.Text may omit (for example field
    /// codes and hidden text), so a string offset is never used as a document CP.
    /// </summary>
    internal static class WordRangeLocator
    {
        public static string ReadText(Word.Range range)
        {
            if (range == null) throw new ArgumentNullException("range");
            Word.Range readable = range.Duplicate;
            SetRetrievalMode(readable, false, false);
            return readable.Text ?? string.Empty;
        }

        public static int GetTextOffset(Word.Range scope, int position)
        {
            if (scope == null) throw new ArgumentNullException("scope");
            if (position < scope.Start || position > scope.End)
                throw new ArgumentOutOfRangeException("position");
            Word.Range prefix = scope.Duplicate;
            prefix.SetRange(scope.Start, position);
            return ReadText(prefix).Length;
        }

        public static Word.Range Resolve(Word.Range scope, string snapshot, LatexSpan span, int searchStart)
        {
            if (scope == null) throw new ArgumentNullException("scope");
            if (snapshot == null) throw new ArgumentNullException("snapshot");
            if (span == null) throw new ArgumentNullException("span");
            if (span.Start < 0 || span.Length <= 0 || span.Start > snapshot.Length - span.Length)
                throw new InvalidOperationException("公式扫描范围已失效，已保留原文。");
            if (searchStart < scope.Start || searchStart > scope.End)
                throw new InvalidOperationException("公式搜索范围已发生变化，已保留原文。");

            string raw = snapshot.Substring(span.Start, span.Length);
            string expectedPrefix = snapshot.Substring(0, span.Start);
            string opening;
            string closing;
            GetDelimiters(raw, out opening, out closing);
            int scopeEnd = scope.End;
            int cursor = searchStart;

            while (cursor < scopeEnd)
            {
                Word.Range opener = FindDelimiter(scope, cursor, scopeEnd, opening);
                if (opener == null) break;
                cursor = opener.End;

                // This identifies the exact occurrence even when identical formulas
                // occur several times or Find also finds an escaped delimiter.
                Word.Range prefix = scope.Duplicate;
                prefix.SetRange(scope.Start, opener.Start);
                if (!string.Equals(ReadText(prefix), expectedPrefix, StringComparison.Ordinal)) continue;

                int closingCursor = opener.End;
                while (closingCursor < scopeEnd)
                {
                    Word.Range closer = FindDelimiter(scope, closingCursor, scopeEnd, closing);
                    if (closer == null) break;
                    closingCursor = closer.End;
                    Word.Range target = scope.Duplicate;
                    target.SetRange(opener.Start, closer.End);
                    string candidateText = ReadText(target);
                    if (!string.Equals(candidateText, raw, StringComparison.Ordinal))
                    {
                        // An earlier escaped closer is still a prefix of the source.
                        // Once the candidate differs, scanning the rest of the story
                        // cannot safely identify this occurrence, so leave it intact.
                        if (!raw.StartsWith(candidateText, StringComparison.Ordinal))
                            throw new InvalidOperationException("公式范围与扫描原文不一致，已保留原文。");
                        continue;
                    }

                    EnsurePlainSource(target, raw);
                    return target;
                }

                // An opener with the right complete prefix has only one valid
                // source occurrence. Do not substitute a later similar formula.
                throw new InvalidOperationException("无法核对公式的完整结束位置，已保留原文。");
            }

            throw new InvalidOperationException("无法在 Word 中准确定位公式原文，已保留原文。");
        }

        private static Word.Range FindDelimiter(Word.Range scope, int start, int end, string delimiter)
        {
            int cursor = start;
            while (cursor < end)
            {
                Word.Range search = scope.Duplicate;
                search.SetRange(cursor, end);
                SetRetrievalMode(search, false, false);
                Word.Find find = search.Find;
                find.ClearFormatting();
                TrySetLanguageOption(delegate { find.ClearAllFuzzyOptions(); });
                find.Text = delimiter;
                find.Forward = true;
                find.Wrap = Word.WdFindWrap.wdFindStop;
                find.Format = false;
                find.MatchWildcards = false;
                find.MatchCase = true;
                find.MatchWholeWord = false;
                find.MatchSoundsLike = false;
                find.MatchAllWordForms = false;
                find.MatchPrefix = false;
                find.MatchSuffix = false;
                find.MatchPhrase = false;
                find.IgnoreSpace = false;
                find.IgnorePunct = false;
                // Some language-specific members are unavailable in installations
                // without the corresponding language support. Reset those that are
                // available; the exact delimiter/source checks below remain mandatory.
                TrySetLanguageOption(delegate { find.MatchFuzzy = false; });
                TrySetLanguageOption(delegate { find.MatchByte = true; });
                TrySetLanguageOption(delegate { find.MatchControl = true; });
                TrySetLanguageOption(delegate { find.MatchDiacritics = true; });
                TrySetLanguageOption(delegate { find.MatchKashida = true; });
                TrySetLanguageOption(delegate { find.MatchAlefHamza = true; });
                TrySetLanguageOption(delegate { find.HanjaPhoneticHangul = false; });

                // Only the one- or two-character delimiter enters Word's Find
                // syntax. Long formulas and LaTeX ^ commands are never searched.
                if (!find.Execute(Replace: Word.WdReplace.wdReplaceNone)) return null;
                if (search.StoryType != scope.StoryType || search.Start < cursor ||
                    search.End > end || search.End <= search.Start)
                    throw new InvalidOperationException("Word 返回了范围外的公式匹配，已保留原文。");

                cursor = search.End;
                if (string.Equals(ReadText(search), delimiter, StringComparison.Ordinal)) return search;
            }
            return null;
        }

        private static void TrySetLanguageOption(Action action)
        {
            try { action(); }
            catch (COMException) { }
        }

        internal static void EnsurePlainSource(Word.Range target, string raw)
        {
            if (target.Fields.Count != 0 || IsInField(target))
                throw new InvalidOperationException("公式源码包含 Word 域，已保留原文。");
            if (target.OMaths.Count != 0)
                throw new InvalidOperationException("公式源码位于已有 Word 公式内，已保留原文。");
            if (raw.IndexOf('\a') >= 0 || target.Cells.Count > 1)
                throw new InvalidOperationException("公式源码跨越表格单元格，已保留原文。");
            if (target.Font.Hidden != 0)
                throw new InvalidOperationException("公式源码包含隐藏文字，已保留原文。");

            Word.Range complete = target.Duplicate;
            SetRetrievalMode(complete, true, true);
            if (!string.Equals(complete.Text ?? string.Empty, raw, StringComparison.Ordinal))
                throw new InvalidOperationException("公式源码包含不可见内容，已保留原文。");
        }

        private static bool IsInField(Word.Range range)
        {
            Word.Range first = range.Duplicate;
            first.SetRange(range.Start, range.Start + 1);
            Word.Range last = range.Duplicate;
            last.SetRange(range.End - 1, range.End);
            return Convert.ToBoolean(first.Information[Word.WdInformation.wdInFieldCode]) ||
                Convert.ToBoolean(first.Information[Word.WdInformation.wdInFieldResult]) ||
                Convert.ToBoolean(last.Information[Word.WdInformation.wdInFieldCode]) ||
                Convert.ToBoolean(last.Information[Word.WdInformation.wdInFieldResult]);
        }

        private static void SetRetrievalMode(Word.Range range, bool hiddenText, bool fieldCodes)
        {
            Word.TextRetrievalMode mode = range.TextRetrievalMode;
            mode.ViewType = Word.WdViewType.wdPrintView;
            mode.IncludeHiddenText = hiddenText;
            mode.IncludeFieldCodes = fieldCodes;
        }

        private static void GetDelimiters(string raw, out string opening, out string closing)
        {
            if (raw.Length >= 4 && raw.StartsWith("$$", StringComparison.Ordinal) && raw.EndsWith("$$", StringComparison.Ordinal))
            {
                opening = "$$";
                closing = "$$";
            }
            else if (raw.Length >= 4 && raw.StartsWith("\\[", StringComparison.Ordinal) && raw.EndsWith("\\]", StringComparison.Ordinal))
            {
                opening = "\\[";
                closing = "\\]";
            }
            else if (raw.Length >= 4 && raw.StartsWith("\\(", StringComparison.Ordinal) && raw.EndsWith("\\)", StringComparison.Ordinal))
            {
                opening = "\\(";
                closing = "\\)";
            }
            else if (raw.Length >= 2 && raw[0] == '$' && raw[raw.Length - 1] == '$')
            {
                opening = "$";
                closing = "$";
            }
            else
            {
                throw new InvalidOperationException("无法核对公式定界符，已保留原文。");
            }
        }
    }
}

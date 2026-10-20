using System;
using System.Collections.Generic;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace WordLatexAddin
{
    /// <summary>Summary returned by selection, paragraph and document conversion commands.</summary>
    public sealed class ConversionSummary
    {
        public int Found { get; set; }
        public int Converted { get; set; }
        public bool RequiresUndo { get; set; }
        public IList<string> Errors { get; private set; }

        public ConversionSummary()
        {
            Errors = new List<string>();
        }

        public string ToUserMessage()
        {
            StringBuilder message = new StringBuilder();
            message.AppendFormat("找到 {0} 个公式，成功转换 {1} 个。", Found, Converted);
            if (Errors.Count > 0)
            {
                message.AppendLine();
                if (RequiresUndo)
                    message.Append("转换已停止，未能完整恢复本次写入。请先按 Ctrl+Z 撤销，再检查文档。");
                else
                    message.AppendFormat("{0} 个公式保留了原文：", Errors.Count);
                int limit = Math.Min(Errors.Count, 5);
                for (int i = 0; i < limit; i++) message.AppendLine().Append("• ").Append(Errors[i]);
            }
            return message.ToString();
        }
    }

    /// <summary>Coordinates safe range scanning and reverse-order replacement in Word.</summary>
    public sealed class LatexConverter
    {
        private readonly Word.Application _application;
        private readonly LatexScanner _scanner;
        private readonly OmmlBuilder _builder;

        public LatexConverter(Word.Application application, LatexScanner scanner, OmmlBuilder builder)
        {
            _application = application;
            _scanner = scanner;
            _builder = builder;
        }

        public bool IsBusy { get; private set; }

        public ConversionSummary ConvertCurrent(Word.Selection selection)
        {
            if (selection == null) throw new InvalidOperationException("Word 中没有活动选区。");
            Word.Range selected = selection.Range.Duplicate;
            if (selected.Start != selected.End)
            {
                string source = WordRangeLocator.ReadText(selected);
                IList<LatexSpan> spans = _scanner.FindAll(source);
                if (spans.Count > 0) return ConvertSpans(selected, source, spans, "转换当前公式");
                // A second pass over a converted selection must not turn its equations
                // and surrounding prose into one giant raw-LaTeX equation.
                if (selected.OMaths.Count > 0 || string.IsNullOrWhiteSpace(source)) return new ConversionSummary();

                ConversionSummary rawSummary = new ConversionSummary { Found = 1 };
                ExecuteUndo("转换当前公式", delegate
                {
                    try
                    {
                        bool display;
                        string latex = _scanner.RemoveDelimiters(source, out display);
                        _builder.ReplaceWithEquation(selected, latex, display);
                        rawSummary.Converted = 1;
                    }
                    catch (Exception exception)
                    {
                        rawSummary.RequiresUndo = exception is EquationRecoveryException;
                        rawSummary.Errors.Add(ShortError(source, exception));
                    }
                });
                return rawSummary;
            }

            Word.Range paragraph = selection.Paragraphs[1].Range.Duplicate;
            string paragraphText = WordRangeLocator.ReadText(paragraph);
            int caret = WordRangeLocator.GetTextOffset(paragraph, selection.Start);
            LatexSpan nearest = _scanner.FindNearest(paragraphText, caret);
            if (nearest == null) return new ConversionSummary();
            return ConvertSpans(paragraph, paragraphText, new List<LatexSpan> { nearest }, "转换当前公式");
        }

        public ConversionSummary ConvertParagraph(Word.Selection selection)
        {
            Word.Range paragraph = selection.Paragraphs[1].Range.Duplicate;
            string source = WordRangeLocator.ReadText(paragraph);
            return ConvertSpans(paragraph, source, _scanner.FindAll(source), "转换当前段落");
        }

        public ConversionSummary ConvertDocument(Word.Document document)
        {
            if (document == null) throw new InvalidOperationException("Word 中没有活动文档。");
            Word.Range body = document.Content.Duplicate;
            string source = WordRangeLocator.ReadText(body);
            return ConvertSpans(body, source, _scanner.FindAll(source), "转换全文");
        }

        public bool TryAutoConvert(Word.Selection selection)
        {
            if (IsBusy || selection == null || selection.Start != selection.End) return false;
            Word.Range paragraph = selection.Paragraphs[1].Range.Duplicate;
            string paragraphText = WordRangeLocator.ReadText(paragraph);
            int caret = WordRangeLocator.GetTextOffset(paragraph, selection.Start);
            if (caret <= 0 || caret > paragraphText.Length) return false;

            char boundary = paragraphText[caret - 1];
            if (!char.IsWhiteSpace(boundary) && ",.;:!?，。；：！？)）]】".IndexOf(boundary) < 0) return false;

            LatexSpan candidate = null;
            foreach (LatexSpan span in _scanner.FindAll(paragraphText))
            {
                if (span.End <= caret) candidate = span;
            }
            if (candidate == null) return false;

            string trailing = paragraphText.Substring(candidate.End, caret - candidate.End);
            if (trailing.Length > 3 || trailing.Trim(' ', '\t', '\r', '\n', ',', '.', ';', ':', '!', '?', '，', '。', '；', '：', '！', '？', ')', '）', ']', '】').Length != 0)
                return false;

            ConversionSummary summary = ConvertSpans(paragraph, paragraphText, new List<LatexSpan> { candidate }, "自动转换 LaTeX");
            return summary.Converted == 1;
        }

        private ConversionSummary ConvertSpans(Word.Range baseRange, string source, IList<LatexSpan> spans, string undoName)
        {
            ConversionSummary summary = new ConversionSummary { Found = spans.Count };
            if (spans.Count == 0) return summary;

            ExecuteUndo(undoName, delegate
            {
                // Resolve every span while the snapshot is still current. Word character
                // positions include structures that Range.Text may omit or represent with
                // multiple UTF-16 units, so string offsets are never document positions.
                List<EquationTarget> targets = new List<EquationTarget>();
                int searchStart = baseRange.Start;
                foreach (LatexSpan span in spans)
                {
                    try
                    {
                        Word.Range target = WordRangeLocator.Resolve(baseRange, source, span, searchStart);
                        targets.Add(new EquationTarget
                        {
                            Range = target.Duplicate,
                            Span = span,
                            Source = source.Substring(span.Start, span.Length)
                        });
                        searchStart = target.End;
                    }
                    catch (Exception exception)
                    {
                        summary.Errors.Add(ShortError(span.Latex, exception));
                    }
                }

                for (int i = targets.Count - 1; i >= 0; i--)
                {
                    EquationTarget target = targets[i];
                    try
                    {
                        if (!string.Equals(WordRangeLocator.ReadText(target.Range), target.Source, StringComparison.Ordinal))
                            throw new InvalidOperationException("公式范围与原文不一致，已保留原文。");
                        _builder.ReplaceWithEquation(target.Range, target.Span.Latex, target.Span.IsDisplay);
                        summary.Converted++;
                    }
                    catch (Exception exception)
                    {
                        summary.Errors.Add(ShortError(target.Span.Latex, exception));
                        if (exception is EquationRecoveryException)
                        {
                            summary.RequiresUndo = true;
                            break;
                        }
                    }
                }
            });
            return summary;
        }

        private sealed class EquationTarget
        {
            public Word.Range Range { get; set; }
            public LatexSpan Span { get; set; }
            public string Source { get; set; }
        }

        private void ExecuteUndo(string name, Action action)
        {
            if (IsBusy) return;
            IsBusy = true;
            bool undoStarted = false;
            try
            {
                try
                {
                    _application.UndoRecord.StartCustomRecord(name);
                    undoStarted = true;
                }
                catch { }
                action();
            }
            finally
            {
                if (undoStarted)
                {
                    try { _application.UndoRecord.EndCustomRecord(); }
                    catch { }
                }
                IsBusy = false;
            }
        }

        private static string ShortError(string latex, Exception exception)
        {
            string source = latex.Length > 36 ? latex.Substring(0, 36) : latex;
            if (source.Length > 0 && char.IsHighSurrogate(source[source.Length - 1])) source = source.Substring(0, source.Length - 1);
            if (source.Length < latex.Length) source += "…";
            return source + " — " + exception.Message;
        }
    }
}

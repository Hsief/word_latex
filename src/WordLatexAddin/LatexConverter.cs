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
                IList<LatexSpan> spans = _scanner.FindAll(selected.Text);
                if (spans.Count > 0) return ConvertSpans(selected, spans, "转换当前公式");

                bool display;
                string latex = _scanner.RemoveDelimiters(selected.Text, out display);
                ConversionSummary rawSummary = new ConversionSummary { Found = 1 };
                ExecuteUndo("转换当前公式", delegate
                {
                    try
                    {
                        _builder.ReplaceWithEquation(selected, latex, display);
                        rawSummary.Converted = 1;
                    }
                    catch (Exception exception)
                    {
                        rawSummary.Errors.Add(ShortError(latex, exception));
                    }
                });
                return rawSummary;
            }

            Word.Range paragraph = selection.Paragraphs[1].Range.Duplicate;
            int caret = selection.Start - paragraph.Start;
            LatexSpan nearest = _scanner.FindNearest(paragraph.Text, caret);
            if (nearest == null) return new ConversionSummary();
            return ConvertSpans(paragraph, new List<LatexSpan> { nearest }, "转换当前公式");
        }

        public ConversionSummary ConvertParagraph(Word.Selection selection)
        {
            Word.Range paragraph = selection.Paragraphs[1].Range.Duplicate;
            return ConvertSpans(paragraph, _scanner.FindAll(paragraph.Text), "转换当前段落");
        }

        public ConversionSummary ConvertDocument(Word.Document document)
        {
            if (document == null) throw new InvalidOperationException("Word 中没有活动文档。");
            Word.Range body = document.Content.Duplicate;
            return ConvertSpans(body, _scanner.FindAll(body.Text), "转换全文");
        }

        public bool TryAutoConvert(Word.Selection selection)
        {
            if (IsBusy || selection == null || selection.Start != selection.End) return false;
            Word.Range paragraph = selection.Paragraphs[1].Range.Duplicate;
            int caret = selection.Start - paragraph.Start;
            if (caret <= 0 || caret > paragraph.Text.Length) return false;

            char boundary = paragraph.Text[Math.Min(caret - 1, paragraph.Text.Length - 1)];
            if (!char.IsWhiteSpace(boundary) && ",.;:!?，。；：！？)）]】".IndexOf(boundary) < 0) return false;

            LatexSpan candidate = null;
            foreach (LatexSpan span in _scanner.FindAll(paragraph.Text))
            {
                if (span.End <= caret) candidate = span;
            }
            if (candidate == null) return false;

            string trailing = paragraph.Text.Substring(candidate.End, caret - candidate.End);
            if (trailing.Length > 3 || trailing.Trim(' ', '\t', '\r', '\n', ',', '.', ';', ':', '!', '?', '，', '。', '；', '：', '！', '？', ')', '）', ']', '】').Length != 0)
                return false;

            ConversionSummary summary = ConvertSpans(paragraph, new List<LatexSpan> { candidate }, "自动转换 LaTeX");
            return summary.Converted == 1;
        }

        private ConversionSummary ConvertSpans(Word.Range baseRange, IList<LatexSpan> spans, string undoName)
        {
            ConversionSummary summary = new ConversionSummary { Found = spans.Count };
            if (spans.Count == 0) return summary;
            int baseStart = baseRange.Start;
            Word.Document document = baseRange.Document;

            ExecuteUndo(undoName, delegate
            {
                for (int i = spans.Count - 1; i >= 0; i--)
                {
                    LatexSpan span = spans[i];
                    try
                    {
                        Word.Range target = document.Range(baseStart + span.Start, baseStart + span.End);
                        _builder.ReplaceWithEquation(target, span.Latex, span.IsDisplay);
                        summary.Converted++;
                    }
                    catch (Exception exception)
                    {
                        summary.Errors.Add(ShortError(span.Latex, exception));
                    }
                }
            });
            return summary;
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

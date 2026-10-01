using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordLatexAddin
{
    /// <summary>Creates IEEE-style centered equations with right-aligned SEQ fields and REF links.</summary>
    public sealed class EquationNumbering
    {
        private const string BookmarkPrefix = "WordLatexEq_";
        private readonly Settings _settings;

        public EquationNumbering(Settings settings)
        {
            _settings = settings;
        }

        public void AddNumber(Word.Selection selection)
        {
            if (selection == null) throw new InvalidOperationException("Word 中没有活动选区。");
            Word.Document document = selection.Document;
            Word.Paragraph paragraph = selection.Paragraphs[1];
            Word.Range paragraphRange = paragraph.Range.Duplicate;
            int contentEnd = Math.Max(paragraphRange.Start, paragraphRange.End - 1);

            if (document.Range(paragraphRange.Start, contentEnd).Text.Trim().Length == 0)
                throw new InvalidOperationException("当前段落没有可编号的公式。");

            paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            paragraph.TabStops.ClearAll();
            float usableWidth = document.PageSetup.PageWidth - document.PageSetup.LeftMargin - document.PageSetup.RightMargin;
            paragraph.TabStops.Add(usableWidth / 2F, Word.WdTabAlignment.wdAlignTabCenter, Word.WdTabLeader.wdTabLeaderSpaces);
            paragraph.TabStops.Add(usableWidth, Word.WdTabAlignment.wdAlignTabRight, Word.WdTabLeader.wdTabLeaderSpaces);

            Word.Range first = document.Range(paragraphRange.Start, paragraphRange.Start);
            first.InsertBefore("\t");
            paragraphRange = paragraph.Range.Duplicate;
            contentEnd = Math.Max(paragraphRange.Start, paragraphRange.End - 1);
            Word.Range number = document.Range(contentEnd, contentEnd);
            number.InsertAfter("\t(");
            number.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            Word.Field field = document.Fields.Add(number, Word.WdFieldType.wdFieldSequence,
                SafeSequenceName(_settings.EquationSequenceName) + " \\* ARABIC", true);
            field.Update();
            Word.Range closing = field.Result.Duplicate;
            closing.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            closing.InsertAfter(")");

            string result = field.Result.Text.Trim();
            string bookmarkName = BookmarkPrefix + (result.Length == 0 ? DateTime.Now.Ticks.ToString() : result) + "_" + Guid.NewGuid().ToString("N").Substring(0, 6);
            document.Bookmarks.Add(bookmarkName, field.Result);
            document.Fields.Update();
        }

        public void InsertReference(Word.Selection selection)
        {
            if (selection == null) throw new InvalidOperationException("Word 中没有活动选区。");
            List<EquationBookmark> equations = GetEquations(selection.Document);
            if (equations.Count == 0) throw new InvalidOperationException("文档中还没有由本插件创建的公式编号。");

            using (EquationReferenceDialog dialog = new EquationReferenceDialog(equations))
            {
                if (dialog.ShowDialog() != DialogResult.OK || dialog.Selected == null) return;
                Word.Range insertion = selection.Range;
                insertion.Text = "(";
                insertion.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                Word.Field reference = selection.Document.Fields.Add(insertion, Word.WdFieldType.wdFieldRef,
                    dialog.Selected.BookmarkName + " \\h", true);
                reference.Update();
                Word.Range close = reference.Result.Duplicate;
                close.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                close.InsertAfter(")");
            }
        }

        private static List<EquationBookmark> GetEquations(Word.Document document)
        {
            List<EquationBookmark> result = new List<EquationBookmark>();
            foreach (Word.Bookmark bookmark in document.Bookmarks)
            {
                if (!bookmark.Name.StartsWith(BookmarkPrefix, StringComparison.Ordinal)) continue;
                string number = bookmark.Range.Text.Trim();
                result.Add(new EquationBookmark(bookmark.Name, "公式 (" + number + ")"));
            }
            return result;
        }

        private static string SafeSequenceName(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return "Equation";
            foreach (char character in value)
            {
                if (!char.IsLetterOrDigit(character) && character != '_') return "Equation";
            }
            return value;
        }
    }

    internal sealed class EquationBookmark
    {
        public EquationBookmark(string bookmarkName, string label)
        {
            BookmarkName = bookmarkName;
            Label = label;
        }
        public string BookmarkName { get; private set; }
        public string Label { get; private set; }
        public override string ToString() { return Label; }
    }

    internal sealed class EquationReferenceDialog : Form
    {
        private readonly ListBox _list;

        public EquationReferenceDialog(IList<EquationBookmark> equations)
        {
            Text = "插入公式引用";
            Font = new Font("Microsoft YaHei UI", 9F);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(350, 300);
            _list = new ListBox { Left = 16, Top = 16, Width = 318, Height = 220 };
            foreach (EquationBookmark equation in equations) _list.Items.Add(equation);
            if (_list.Items.Count > 0) _list.SelectedIndex = 0;
            Button ok = new Button { Left = 172, Top = 252, Width = 75, Text = "插入", DialogResult = DialogResult.OK };
            Button cancel = new Button { Left = 259, Top = 252, Width = 75, Text = "取消", DialogResult = DialogResult.Cancel };
            Controls.AddRange(new Control[] { _list, ok, cancel });
            AcceptButton = ok;
            CancelButton = cancel;
        }

        public EquationBookmark Selected { get { return _list.SelectedItem as EquationBookmark; } }
    }
}

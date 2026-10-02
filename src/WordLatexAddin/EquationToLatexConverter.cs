using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using AfterMathCore;
using Word = Microsoft.Office.Interop.Word;

namespace WordLatexAddin
{
    /// <summary>Replaces selected Word OMath objects with editable LaTeX source text.</summary>
    public sealed class EquationToLatexConverter
    {
        private readonly Word.Application _application;
        private readonly OmmlToLatexConverter _converter;

        public EquationToLatexConverter(Word.Application application, OmmlToLatexConverter converter)
        {
            _application = application;
            _converter = converter;
        }

        public bool IsBusy { get; private set; }

        /// <summary>Converts the equation at the caret, or all equations intersecting the selection.</summary>
        public ConversionSummary ConvertCurrent(Word.Selection selection)
        {
            if (selection == null) throw new InvalidOperationException("Word 中没有活动选区。");
            Word.Document document = selection.Document;
            return ConvertTargets(document, FindTargets(document, selection.Range));
        }

        /// <summary>Converts every native equation in the main document body.</summary>
        public ConversionSummary ConvertDocument(Word.Document document)
        {
            if (document == null) throw new InvalidOperationException("Word 中没有活动文档。");
            return ConvertTargets(document, FindTargets(document, null));
        }

        private ConversionSummary ConvertTargets(Word.Document document, IList<EquationTarget> targets)
        {
            ConversionSummary summary = new ConversionSummary { Found = targets.Count };
            if (targets.Count == 0) return summary;

            ExecuteUndo("Word 公式转 LaTeX", delegate
            {
                for (int index = targets.Count - 1; index >= 0; index--)
                {
                    EquationTarget target = targets[index];
                    try
                    {
                        string latex = Regex.Replace(_converter.Convert(target.Xml), @"\s*[\r\n]+\s*", " ");
                        // Setting Range.Text keeps the OMath object around the new text; delete the object, then insert plain text.
                        document.Range(target.Start, target.End).Delete();
                        document.Range(target.Start, target.Start).InsertAfter(target.Display ? "$$" + latex + "$$" : "$" + latex + "$");
                        summary.Converted++;
                    }
                    catch (Exception exception)
                    {
                        summary.Errors.Add("位置 " + target.Start + " — " + exception.Message);
                    }
                }
            });
            return summary;
        }

        private static IList<EquationTarget> FindTargets(Word.Document document, Word.Range selection)
        {
            List<EquationTarget> targets = new List<EquationTarget>();
            bool wholeDocument = selection == null;
            bool collapsed = !wholeDocument && selection.Start == selection.End;
            int caret = wholeDocument ? 0 : selection.Start;
            for (int index = 1; index <= document.OMaths.Count; index++)
            {
                Word.OMath equation = document.OMaths[index];
                Word.Range range = equation.Range;
                bool selected = wholeDocument || (collapsed
                    ? range.Start <= caret && caret <= range.End
                    : range.Start < selection.End && range.End > selection.Start);
                if (!selected) continue;
                targets.Add(new EquationTarget
                {
                    Start = range.Start,
                    End = range.End,
                    Display = equation.Type == Word.WdOMathType.wdOMathDisplay,
                    Xml = range.WordOpenXML
                });
            }
            return targets;
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

        private sealed class EquationTarget
        {
            public int Start { get; set; }
            public int End { get; set; }
            public bool Display { get; set; }
            public string Xml { get; set; }
        }
    }
}

using System;
using AfterMathCore;
using Word = Microsoft.Office.Interop.Word;

namespace WordLatexAddin
{
    /// <summary>Bridges generated OMML XML into a live Word range as an editable OMath object.</summary>
    public sealed class OmmlBuilder
    {
        private readonly OmmlGenerator _generator;

        public OmmlBuilder(OmmlGenerator generator)
        {
            _generator = generator;
        }

        public Word.Range ReplaceWithEquation(Word.Range target, string latex, bool display)
        {
            if (target == null) throw new ArgumentNullException("target");
            string source = target.Text;
            int start = target.Start;
            Word.Document document = target.Document;
            string omml = _generator.Convert(latex);

            try
            {
                target.Text = string.Empty;
                object transformSource = Type.Missing;
                target.InsertXML(omml, ref transformSource);
                Word.Range inserted = document.Range(start, Math.Max(start, target.End));
                if (inserted.OMaths.Count == 0)
                {
                    int probeEnd = Math.Min(document.Content.End, Math.Max(start + 1, target.End + 1));
                    inserted = document.Range(start, probeEnd);
                }
                if (inserted.OMaths.Count == 0) throw new InvalidOperationException("Word 未创建 OMath 对象。");

                Word.OMath equation = inserted.OMaths[1];
                equation.BuildUp();
                equation.Type = display ? Word.WdOMathType.wdOMathDisplay : Word.WdOMathType.wdOMathInline;
                return equation.Range;
            }
            catch
            {
                Word.Range recovery = document.Range(start, start);
                recovery.Text = source;
                throw;
            }
        }
    }
}

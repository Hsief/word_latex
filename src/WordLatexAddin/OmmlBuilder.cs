using System;
using System.IO;
using System.IO.Compression;
using System.Text;
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
            string packagePath = Path.Combine(Path.GetTempPath(), "WordLatexVSTO-" + Guid.NewGuid().ToString("N") + ".docx");

            try
            {
                CreateEquationPackage(packagePath, omml);
                target.Text = string.Empty;
                int countBefore = document.OMaths.Count;
                object bookmark = "Equation";
                object confirmConversions = false;
                object link = false;
                object attachment = false;
                target.InsertFile(packagePath, ref bookmark, ref confirmConversions, ref link, ref attachment);
                if (document.OMaths.Count <= countBefore) throw new InvalidOperationException("Word 未创建 OMath 对象。");

                Word.OMath equation = FindInsertedEquation(document, start);
                if (equation == null) throw new InvalidOperationException("无法定位 Word 新建的 OMath 对象。");
                equation.Type = display ? Word.WdOMathType.wdOMathDisplay : Word.WdOMathType.wdOMathInline;
                return equation.Range;
            }
            catch
            {
                Word.Range recovery = document.Range(start, start);
                recovery.Text = source;
                throw;
            }
            finally
            {
                try { if (File.Exists(packagePath)) File.Delete(packagePath); }
                catch { }
            }
        }

        private static Word.OMath FindInsertedEquation(Word.Document document, int insertionStart)
        {
            Word.OMath nearest = null;
            int nearestDistance = int.MaxValue;
            for (int index = 1; index <= document.OMaths.Count; index++)
            {
                Word.OMath candidate = document.OMaths[index];
                int distance = Math.Abs(candidate.Range.Start - insertionStart);
                if (distance < nearestDistance)
                {
                    nearest = candidate;
                    nearestDistance = distance;
                    if (distance == 0) break;
                }
            }
            return nearest;
        }

        private static void CreateEquationPackage(string path, string omml)
        {
            const string contentTypes = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>" +
                "</Types>";
            const string relationships = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>" +
                "</Relationships>";
            string documentXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><w:body><w:p>" +
                "<w:bookmarkStart w:id=\"0\" w:name=\"Equation\"/>" + omml +
                "<w:bookmarkEnd w:id=\"0\"/></w:p><w:sectPr/></w:body></w:document>";

            using (FileStream stream = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Create, false))
            {
                AddPackagePart(archive, "[Content_Types].xml", contentTypes);
                AddPackagePart(archive, "_rels/.rels", relationships);
                AddPackagePart(archive, "word/document.xml", documentXml);
            }
        }

        private static void AddPackagePart(ZipArchive archive, string name, string content)
        {
            ZipArchiveEntry entry = archive.CreateEntry(name, CompressionLevel.Optimal);
            using (Stream entryStream = entry.Open())
            using (StreamWriter writer = new StreamWriter(entryStream, new UTF8Encoding(false)))
            {
                writer.Write(content);
            }
        }
    }
}

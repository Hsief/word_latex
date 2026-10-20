using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using AfterMathCore;
using Word = Microsoft.Office.Interop.Word;

namespace WordLatexAddin
{
    /// <summary>Stops a batch if Word could not restore the affected source range.</summary>
    internal sealed class EquationRecoveryException : InvalidOperationException
    {
        public EquationRecoveryException(Exception conversionError, Exception recoveryError)
            : base("Word 未能完整恢复公式原文，请按 Ctrl+Z 撤销本次转换。",
                new AggregateException(conversionError, recoveryError)) { }
    }

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
            Word.Range edit = target.Duplicate;
            ValidateSource(edit);
            string source = WordRangeLocator.ReadText(edit);
            string originalXml = edit.WordOpenXML;
            string omml = _generator.Convert(latex);
            string packagePath = Path.Combine(Path.GetTempPath(), "WordLatexVSTO-" + Guid.NewGuid().ToString("N") + ".docx");
            string bookmarkName = "WL" + Guid.NewGuid().ToString("N");

            try
            {
                // Parsing, XML backup and package creation must all succeed before Word
                // is changed. A preparation failure must never insert another source copy.
                CreateEquationPackage(packagePath, omml, bookmarkName);
                return ReplacePrepared(edit, source, originalXml, packagePath, bookmarkName, display);
            }
            finally
            {
                try { if (File.Exists(packagePath)) File.Delete(packagePath); }
                catch { }
            }
        }

        private static void ValidateSource(Word.Range target)
        {
            if (target.Start == target.End) throw new InvalidOperationException("没有可转换的公式原文。");
            if (target.Document.TrackRevisions || target.Revisions.Count > 0)
                throw new InvalidOperationException("请先关闭修订并处理公式范围内的修订，再转换公式。");
            string source = WordRangeLocator.ReadText(target);
            if (string.IsNullOrWhiteSpace(source) || source.IndexOf('\a') >= 0)
                throw new InvalidOperationException("公式范围为空或跨越单元格，已保留原文。");
            if (source.EndsWith("\r", StringComparison.Ordinal) && target.End == GetStoryEnd(target))
                throw new InvalidOperationException("请只选择公式源码，不要包含文档末尾的段落标记。");
            WordRangeLocator.EnsurePlainSource(target, source);
        }

        private static Word.Range ReplacePrepared(Word.Range edit, string source, string originalXml,
            string packagePath, string bookmarkName, bool display)
        {
            int start = edit.Start;
            int originalEnd = edit.End;
            int storyEndBefore = GetStoryEnd(edit);
            try
            {
                // Keep the same story (body, header, footnote, etc.). Never reinterpret
                // these positions with Document.Range, which addresses the main body.
                int deleted = edit.Delete();
                if (deleted == 0 || GetStoryEnd(edit) != storyEndBefore - (originalEnd - start))
                    throw new InvalidOperationException("Word 未完整删除公式源码，已取消替换。");

                edit.SetRange(start, start);
                object bookmark = bookmarkName;
                object confirmConversions = false;
                object link = false;
                object attachment = false;
                edit.InsertFile(packagePath, ref bookmark, ref confirmConversions, ref link, ref attachment);

                Word.Range inserted = ChangedRange(edit, start, originalEnd, storyEndBefore);
                Word.OMath equation = RequireInsertedEquation(inserted);
                equation.Type = display ? Word.WdOMathType.wdOMathDisplay : Word.WdOMathType.wdOMathInline;
                inserted = ChangedRange(edit, start, originalEnd, storyEndBefore);
                return RequireInsertedEquation(inserted).Range.Duplicate;
            }
            catch (Exception conversionError)
            {
                try
                {
                    Word.Range affected = ChangedRange(edit, start, originalEnd, storyEndBefore);
                    bool unchanged = GetStoryEnd(edit) == storyEndBefore && affected.OMaths.Count == 0 &&
                        string.Equals(WordRangeLocator.ReadText(affected), source, StringComparison.Ordinal);
                    if (!unchanged)
                    {
                        // Replace exactly the deleted/partly inserted region, including
                        // any half-created OMath; preserve the original formatted XML.
                        if (affected.Start != affected.End && affected.Delete() == 0)
                            throw new InvalidOperationException("无法移除未完成的公式插入。");
                        affected.SetRange(start, start);
                        object transform = Type.Missing;
                        affected.InsertXML(originalXml, ref transform);
                        Word.Range restored = ChangedRange(affected, start, originalEnd, storyEndBefore);
                        if (GetStoryEnd(affected) != storyEndBefore || restored.OMaths.Count != 0 ||
                            !string.Equals(WordRangeLocator.ReadText(restored), source, StringComparison.Ordinal))
                            throw new InvalidOperationException("恢复后的公式范围与原文不一致。");
                    }
                }
                catch (Exception recoveryError)
                {
                    throw new EquationRecoveryException(conversionError, recoveryError);
                }
                throw;
            }
        }

        private static int GetStoryEnd(Word.Range range)
        {
            Word.Range story = range.Duplicate;
            story.Expand(Word.WdUnits.wdStory);
            return story.End;
        }

        private static Word.Range ChangedRange(Word.Range anchor, int start, int originalEnd, int storyEndBefore)
        {
            int storyEnd = GetStoryEnd(anchor);
            int end = originalEnd + (storyEnd - storyEndBefore);
            if (end < start || end > storyEnd)
                throw new InvalidOperationException("无法确定本次写入的局部范围。");
            Word.Range result = anchor.Duplicate;
            result.SetRange(start, end);
            return result;
        }

        private static Word.OMath RequireInsertedEquation(Word.Range inserted)
        {
            if (inserted.OMaths.Count != 1) throw new InvalidOperationException("Word 未在目标位置创建唯一公式对象。");
            Word.OMath equation = inserted.OMaths[1];
            Word.Range equationRange = equation.Range;
            if (equationRange.Start < inserted.Start || equationRange.End > inserted.End)
                throw new InvalidOperationException("Word 公式对象超出了本次插入范围。");
            Word.Range before = inserted.Duplicate;
            before.SetRange(inserted.Start, equationRange.Start);
            Word.Range after = inserted.Duplicate;
            after.SetRange(equationRange.End, inserted.End);
            if (!string.IsNullOrWhiteSpace(WordRangeLocator.ReadText(before)) ||
                !string.IsNullOrWhiteSpace(WordRangeLocator.ReadText(after)))
                throw new InvalidOperationException("公式插入范围内仍存在未转换的源码。");
            return equation;
        }

        private static void CreateEquationPackage(string path, string omml, string bookmarkName)
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
                "<w:bookmarkStart w:id=\"0\" w:name=\"" + bookmarkName + "\"/>" + omml +
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

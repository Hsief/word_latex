using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using AfterMathCore.LatexParser;

namespace AfterMathCore
{
    /// <summary>Converts the AfterMathCore equation tree into editable Office Math Markup Language.</summary>
    public sealed class OmmlGenerator
    {
        public const string MathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        public const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        private static readonly XNamespace M = MathNamespace;
        private static readonly XNamespace W = WordNamespace;
        private static readonly Regex ResearchTerms = new Regex("(SO\\(3\\)|SE\\(3\\)|FMCW|LiDAR|IMU)", RegexOptions.Compiled);

        public string Convert(string latex)
        {
            LatexNode tree = new LatexParser.LatexParser(latex).Parse();
            XElement root = new XElement(M + "oMath",
                new XAttribute(XNamespace.Xmlns + "m", M),
                new XAttribute(XNamespace.Xmlns + "w", W));
            AppendNode(root, tree, MathVariant.Default);
            return root.ToString(SaveOptions.DisableFormatting);
        }

        public string ConvertForInsertion(string latex, bool display)
        {
            string equation = Convert(latex);
            if (!display) return equation;
            XElement math = XElement.Parse(equation);
            XElement paragraph = new XElement(M + "oMathPara",
                new XAttribute(XNamespace.Xmlns + "m", M),
                new XAttribute(XNamespace.Xmlns + "w", W),
                math);
            return paragraph.ToString(SaveOptions.DisableFormatting);
        }

        private static void AppendNode(XContainer parent, LatexNode node, MathVariant inheritedVariant)
        {
            if (node.Kind == LatexNodeKind.Sequence)
            {
                foreach (LatexNode child in node.Children) AppendNode(parent, child, inheritedVariant);
                return;
            }
            if (node.Kind == LatexNodeKind.Text)
            {
                AppendText(parent, node.Value ?? string.Empty, node.Variant == MathVariant.Default ? inheritedVariant : node.Variant);
                return;
            }
            if (node.Kind == LatexNodeKind.Styled)
            {
                foreach (LatexNode child in node.Children) AppendNode(parent, child, node.Variant);
                return;
            }
            if (node.Kind == LatexNodeKind.Fraction)
            {
                XElement fraction = new XElement(M + "f");
                if (node.Value == "noBar")
                {
                    fraction.Add(new XElement(M + "fPr", new XElement(M + "type", new XAttribute(M + "val", "noBar"))));
                }
                XElement numerator = new XElement(M + "num");
                XElement denominator = new XElement(M + "den");
                AppendNode(numerator, node.Children[0], inheritedVariant);
                AppendNode(denominator, node.Children[1], inheritedVariant);
                fraction.Add(numerator, denominator);
                parent.Add(fraction);
                return;
            }
            if (node.Kind == LatexNodeKind.Radical)
            {
                XElement radical = new XElement(M + "rad");
                XElement degree = new XElement(M + "deg");
                LatexNode degreeNode;
                if (node.Slots.TryGetValue("degree", out degreeNode)) AppendNode(degree, degreeNode, inheritedVariant);
                XElement element = new XElement(M + "e");
                AppendNode(element, node.Children[0], inheritedVariant);
                radical.Add(degree, element);
                parent.Add(radical);
                return;
            }
            if (node.Kind == LatexNodeKind.Accent)
            {
                IDictionary<string, string> accents = new Dictionary<string, string>
                {
                    { "hat", "̂" }, { "bar", "̅" }, { "tilde", "̃" }, { "vec", "⃗" },
                    { "overrightarrow", "⃗" }, { "overleftarrow", "⃖" }, { "dot", "̇" }, { "ddot", "̈" },
                    { "check", "̌" }, { "breve", "̆" }, { "acute", "́" }, { "grave", "̀" }
                };
                XElement accent = new XElement(M + "acc",
                    new XElement(M + "accPr", new XElement(M + "chr", new XAttribute(M + "val", accents[node.Value]))));
                XElement element = new XElement(M + "e");
                AppendNode(element, node.Children[0], inheritedVariant);
                accent.Add(element);
                parent.Add(accent);
                return;
            }
            if (node.Kind == LatexNodeKind.Script)
            {
                AppendScript(parent, node, inheritedVariant);
                return;
            }
            if (node.Kind == LatexNodeKind.Nary)
            {
                XElement nary = new XElement(M + "nary",
                    new XElement(M + "naryPr", new XElement(M + "chr", new XAttribute(M + "val", node.Value ?? "∑"))));
                XElement sub = new XElement(M + "sub");
                XElement sup = new XElement(M + "sup");
                LatexNode slot;
                if (node.Slots.TryGetValue("sub", out slot)) AppendNode(sub, slot, inheritedVariant);
                if (node.Slots.TryGetValue("sup", out slot)) AppendNode(sup, slot, inheritedVariant);
                nary.Add(sub, sup, new XElement(M + "e"));
                parent.Add(nary);
                return;
            }
            if (node.Kind == LatexNodeKind.Matrix)
            {
                XElement matrix = new XElement(M + "m");
                foreach (LatexNode rowNode in node.Children)
                {
                    XElement row = new XElement(M + "mr");
                    foreach (LatexNode cellNode in rowNode.Children)
                    {
                        XElement cell = new XElement(M + "e");
                        AppendNode(cell, cellNode, inheritedVariant);
                        row.Add(cell);
                    }
                    matrix.Add(row);
                }
                parent.Add(matrix);
                return;
            }
            if (node.Kind == LatexNodeKind.Delimiter)
            {
                LatexNode endNode;
                string ending = node.Slots.TryGetValue("end", out endNode) ? endNode.Value ?? string.Empty : ")";
                XElement delimiter = new XElement(M + "d",
                    new XElement(M + "dPr",
                        new XElement(M + "begChr", new XAttribute(M + "val", node.Value ?? "(")),
                        new XElement(M + "endChr", new XAttribute(M + "val", ending))));
                XElement element = new XElement(M + "e");
                AppendNode(element, node.Children[0], inheritedVariant);
                delimiter.Add(element);
                parent.Add(delimiter);
                return;
            }
            if (node.Kind == LatexNodeKind.Bar)
            {
                XElement bar = new XElement(M + "bar",
                    new XElement(M + "barPr", new XElement(M + "pos", new XAttribute(M + "val", node.Value == "bot" ? "bot" : "top"))));
                XElement element = new XElement(M + "e");
                AppendNode(element, node.Children[0], inheritedVariant);
                bar.Add(element);
                parent.Add(bar);
                return;
            }
            if (node.Kind == LatexNodeKind.LimitUpper || node.Kind == LatexNodeKind.LimitLower)
            {
                XElement limit = new XElement(M + (node.Kind == LatexNodeKind.LimitUpper ? "limUpp" : "limLow"));
                XElement basis = new XElement(M + "e");
                AppendNode(basis, node.Children[0], inheritedVariant);
                XElement limitValue = new XElement(M + "lim");
                LatexNode limitNode;
                if (node.Slots.TryGetValue("limit", out limitNode)) AppendNode(limitValue, limitNode, inheritedVariant);
                limit.Add(basis, limitValue);
                parent.Add(limit);
                return;
            }
            if (node.Kind == LatexNodeKind.GroupCharacter)
            {
                bool bottom = node.Value == "bot";
                XElement group = new XElement(M + "groupChr",
                    new XElement(M + "groupChrPr",
                        new XElement(M + "chr", new XAttribute(M + "val", bottom ? "⏟" : "⏞")),
                        new XElement(M + "pos", new XAttribute(M + "val", bottom ? "bot" : "top")),
                        new XElement(M + "vertJc", new XAttribute(M + "val", bottom ? "bot" : "top"))));
                XElement element = new XElement(M + "e");
                AppendNode(element, node.Children[0], inheritedVariant);
                group.Add(element);
                parent.Add(group);
                return;
            }
            if (node.Kind == LatexNodeKind.BorderBox)
            {
                XElement box = new XElement(M + "borderBox");
                XElement element = new XElement(M + "e");
                AppendNode(element, node.Children[0], inheritedVariant);
                box.Add(element);
                parent.Add(box);
                return;
            }

            throw new LatexParseException("无法生成 OMML 节点：" + node.Kind);
        }

        private static void AppendScript(XContainer parent, LatexNode node, MathVariant variant)
        {
            LatexNode sub;
            LatexNode sup;
            bool hasSub = node.Slots.TryGetValue("sub", out sub);
            bool hasSup = node.Slots.TryGetValue("sup", out sup);
            string elementName = hasSub && hasSup ? "sSubSup" : hasSub ? "sSub" : "sSup";
            XElement script = new XElement(M + elementName);
            XElement basis = new XElement(M + "e");
            AppendNode(basis, node.Children[0], variant);
            script.Add(basis);
            if (hasSub)
            {
                XElement subElement = new XElement(M + "sub");
                AppendNode(subElement, sub, variant);
                script.Add(subElement);
            }
            if (hasSup)
            {
                XElement supElement = new XElement(M + "sup");
                AppendNode(supElement, sup, variant);
                script.Add(supElement);
            }
            parent.Add(script);
        }

        private static void AppendText(XContainer parent, string value, MathVariant variant)
        {
            value = Regex.Replace(value.Replace("~", " "), @"\s+", " ");
            if (value.Length == 0) return;
            if (variant != MathVariant.Default)
            {
                parent.Add(CreateRun(value, variant));
                return;
            }

            int start = 0;
            foreach (Match match in ResearchTerms.Matches(value))
            {
                if (match.Index > start) parent.Add(CreateRun(value.Substring(start, match.Index - start), MathVariant.Default));
                parent.Add(CreateRun(match.Value, MathVariant.Plain));
                start = match.Index + match.Length;
            }
            if (start < value.Length) parent.Add(CreateRun(value.Substring(start), MathVariant.Default));
        }

        private static XElement CreateRun(string value, MathVariant variant)
        {
            XElement run = new XElement(M + "r");
            if (variant != MathVariant.Default)
            {
                string style = variant == MathVariant.Plain ? "p" : variant == MathVariant.Bold ? "b" : variant == MathVariant.Italic ? "i" : "bi";
                run.Add(new XElement(M + "rPr", new XElement(M + "sty", new XAttribute(M + "val", style))));
            }
            run.Add(new XElement(M + "t", value));
            return run;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace AfterMathCore
{
    /// <summary>Converts Word Office Math (OMML) XML into readable, editable LaTeX.</summary>
    public sealed class OmmlToLatexConverter
    {
        private static readonly XNamespace M = OmmlGenerator.MathNamespace;

        private static readonly IDictionary<string, string> Symbols = new Dictionary<string, string>
        {
            { "α", @"\alpha" }, { "β", @"\beta" }, { "γ", @"\gamma" }, { "δ", @"\delta" },
            { "ε", @"\epsilon" }, { "ϵ", @"\varepsilon" }, { "ζ", @"\zeta" }, { "η", @"\eta" },
            { "θ", @"\theta" }, { "ϑ", @"\vartheta" }, { "κ", @"\kappa" }, { "λ", @"\lambda" },
            { "μ", @"\mu" }, { "ν", @"\nu" }, { "ξ", @"\xi" }, { "π", @"\pi" },
            { "ρ", @"\rho" }, { "σ", @"\sigma" }, { "τ", @"\tau" }, { "υ", @"\upsilon" },
            { "φ", @"\phi" }, { "ϕ", @"\varphi" }, { "χ", @"\chi" }, { "ψ", @"\psi" }, { "ω", @"\omega" },
            { "Γ", @"\Gamma" }, { "Δ", @"\Delta" }, { "Θ", @"\Theta" }, { "Λ", @"\Lambda" },
            { "Ξ", @"\Xi" }, { "Π", @"\Pi" }, { "Σ", @"\Sigma" }, { "Υ", @"\Upsilon" },
            { "Φ", @"\Phi" }, { "Ψ", @"\Psi" }, { "Ω", @"\Omega" },
            { "∞", @"\infty" }, { "∂", @"\partial" }, { "∇", @"\nabla" }, { "∈", @"\in" },
            { "∉", @"\notin" }, { "≤", @"\leq" }, { "≥", @"\geq" }, { "≠", @"\neq" },
            { "≈", @"\approx" }, { "≡", @"\equiv" }, { "∝", @"\propto" }, { "±", @"\pm" },
            { "∓", @"\mp" }, { "×", @"\times" }, { "·", @"\cdot" }, { "÷", @"\div" },
            { "→", @"\to" }, { "←", @"\leftarrow" }, { "↔", @"\leftrightarrow" },
            { "⇒", @"\Rightarrow" }, { "⇐", @"\Leftarrow" }, { "⇔", @"\Leftrightarrow" },
            { "∪", @"\cup" }, { "∩", @"\cap" }, { "⊂", @"\subset" }, { "⊆", @"\subseteq" },
            { "⊃", @"\supset" }, { "⊇", @"\supseteq" }, { "∅", @"\emptyset" },
            { "∧", @"\land" }, { "∨", @"\lor" }, { "¬", @"\neg" }, { "⊥", @"\perp" },
            { "∥", @"\parallel" }, { "⊕", @"\oplus" }, { "⊗", @"\otimes" },
            { "ℝ", @"\mathbb{R}" }, { "ℕ", @"\mathbb{N}" }, { "ℤ", @"\mathbb{Z}" },
            { "ℚ", @"\mathbb{Q}" }, { "ℂ", @"\mathbb{C}" }, { "ℏ", @"\hbar" },
            { "…", @"\ldots" }, { "⋯", @"\cdots" }, { "⋮", @"\vdots" }, { "⋱", @"\ddots" }
        };

        private static readonly ISet<string> Functions = new HashSet<string>(StringComparer.Ordinal)
        {
            "sin", "cos", "tan", "cot", "sec", "csc", "arcsin", "arccos", "arctan",
            "sinh", "cosh", "tanh", "log", "ln", "exp", "min", "max", "sup", "inf",
            "lim", "det", "dim", "ker", "gcd"
        };

        /// <summary>Converts an OMML fragment or a complete WordprocessingML package fragment.</summary>
        public string Convert(string xml)
        {
            if (string.IsNullOrWhiteSpace(xml)) throw new ArgumentException("OMML 内容为空。", "xml");
            XDocument document = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
            XElement root = document.Root;
            if (root == null) throw new FormatException("OMML XML 没有根节点。");

            IList<XElement> equations = root.DescendantsAndSelf(M + "oMath")
                .Where(element => !element.Ancestors(M + "oMath").Any()).ToList();
            if (equations.Count == 0 && root.Name.Namespace == M) equations.Add(root);
            if (equations.Count == 0) throw new FormatException("没有找到 Word 原生公式节点 m:oMath。");

            return string.Join(Environment.NewLine, equations.Select(Render).Select(value => value.Trim()).Where(value => value.Length > 0));
        }

        private static string Render(XElement element)
        {
            string name = element.Name.LocalName;
            if (IsPropertyElement(name)) return string.Empty;
            if (name == "r") return RenderRun(element);
            if (name == "t") return EscapeText(element.Value);
            if (name == "f") return RenderFraction(element);
            if (name == "rad") return RenderRadical(element);
            if (name == "sSub") return Script(element, true, false);
            if (name == "sSup") return Script(element, false, true);
            if (name == "sSubSup") return Script(element, true, true);
            if (name == "sPre") return RenderPreScript(element);
            if (name == "nary") return RenderNary(element);
            if (name == "d") return RenderDelimiter(element);
            if (name == "m") return RenderMatrix(element, "matrix");
            if (name == "eqArr") return RenderEquationArray(element);
            if (name == "acc") return RenderAccent(element);
            if (name == "bar") return RenderBar(element);
            if (name == "groupChr") return RenderGroupCharacter(element);
            if (name == "func") return RenderFunction(element);
            if (name == "limLow" || name == "limUpp") return RenderLimit(element, name == "limUpp");
            if (name == "borderBox") return @"\boxed{" + RenderChild(element, "e") + "}";
            return RenderChildren(element);
        }

        private static string RenderRun(XElement run)
        {
            string text = string.Concat(run.Descendants(M + "t").Select(node => node.Value));
            string latex = EscapeText(text);
            XElement style = run.Element(M + "rPr") == null ? null : run.Element(M + "rPr").Element(M + "sty");
            string value = AttributeValue(style, "val");
            if (string.IsNullOrEmpty(latex) || string.IsNullOrEmpty(value) || value == "i") return latex;
            if (value == "p") return NeedsTextStyle(text) ? @"\mathrm{" + latex + "}" : latex;
            if (value == "b") return (latex.Contains("\\") ? @"\boldsymbol{" : @"\mathbf{") + latex + "}";
            if (value == "bi") return @"\boldsymbol{" + latex + "}";
            return latex;
        }

        private static string RenderFraction(XElement fraction)
        {
            string numerator = RenderChild(fraction, "num");
            string denominator = RenderChild(fraction, "den");
            XElement type = fraction.Descendants(M + "type").FirstOrDefault();
            return AttributeValue(type, "val") == "noBar"
                ? @"\binom{" + numerator + "}{" + denominator + "}"
                : @"\frac{" + numerator + "}{" + denominator + "}";
        }

        private static string RenderRadical(XElement radical)
        {
            string degree = RenderChild(radical, "deg");
            string body = RenderChild(radical, "e");
            return string.IsNullOrEmpty(degree) ? @"\sqrt{" + body + "}" : @"\sqrt[" + degree + "]{" + body + "}";
        }

        private static string Script(XElement script, bool subscript, bool superscript)
        {
            string result = GroupBasis(RenderChild(script, "e"));
            if (subscript) result += "_{" + RenderChild(script, "sub") + "}";
            if (superscript) result += "^{" + RenderChild(script, "sup") + "}";
            return result;
        }

        private static string RenderPreScript(XElement script)
        {
            return "^{" + RenderChild(script, "sup") + "}_{" + RenderChild(script, "sub") + "}{" + RenderChild(script, "e") + "}";
        }

        private static string RenderNary(XElement nary)
        {
            XElement character = nary.Descendants(M + "chr").FirstOrDefault();
            string operation = NaryCommand(AttributeValue(character, "val"));
            string sub = RenderChild(nary, "sub");
            string sup = RenderChild(nary, "sup");
            string body = RenderChild(nary, "e");
            if (sub.Length > 0) operation += "_{" + sub + "}";
            if (sup.Length > 0) operation += "^{" + sup + "}";
            return operation + (body.Length > 0 ? " " + body : string.Empty);
        }

        private static string RenderDelimiter(XElement delimiter)
        {
            XElement properties = delimiter.Element(M + "dPr");
            string begin = AttributeValue(properties == null ? null : properties.Element(M + "begChr"), "val");
            string end = AttributeValue(properties == null ? null : properties.Element(M + "endChr"), "val");
            if (string.IsNullOrEmpty(begin)) begin = "(";
            if (string.IsNullOrEmpty(end)) end = ")";
            XElement content = delimiter.Elements(M + "e").FirstOrDefault();
            XElement matrix = content == null ? null : content.Elements(M + "m").FirstOrDefault();
            if (matrix != null && content.Elements().All(node => node == matrix || IsPropertyElement(node.Name.LocalName)))
            {
                string environment = MatrixEnvironment(begin, end);
                if (environment != null) return RenderMatrix(matrix, environment);
            }
            return @"\left" + DelimiterCommand(begin, true) + RenderChildren(content) + @"\right" + DelimiterCommand(end, false);
        }

        private static string RenderMatrix(XElement matrix, string environment)
        {
            IEnumerable<string> rows = matrix.Elements(M + "mr").Select(row =>
                string.Join(" & ", row.Elements(M + "e").Select(RenderChildren)));
            return @"\begin{" + environment + "}" + string.Join(@" \\ ", rows) + @"\end{" + environment + "}";
        }

        private static string RenderEquationArray(XElement array)
        {
            IEnumerable<string> rows = array.Elements(M + "e").Select(RenderChildren);
            return @"\begin{aligned}" + string.Join(@" \\ ", rows) + @"\end{aligned}";
        }

        private static string RenderAccent(XElement accent)
        {
            XElement character = accent.Descendants(M + "chr").FirstOrDefault();
            string command = AccentCommand(AttributeValue(character, "val"));
            return command + "{" + RenderChild(accent, "e") + "}";
        }

        private static string RenderBar(XElement bar)
        {
            XElement position = bar.Descendants(M + "pos").FirstOrDefault();
            string command = AttributeValue(position, "val") == "bot" ? @"\underline" : @"\overline";
            return command + "{" + RenderChild(bar, "e") + "}";
        }

        private static string RenderGroupCharacter(XElement group)
        {
            XElement position = group.Descendants(M + "pos").FirstOrDefault();
            string command = AttributeValue(position, "val") == "bot" ? @"\underbrace" : @"\overbrace";
            return command + "{" + RenderChild(group, "e") + "}";
        }

        private static string RenderFunction(XElement function)
        {
            string name = RenderChild(function, "fName");
            string argument = RenderChild(function, "e");
            string plain = Regex.Replace(name, @"\\mathrm\{([^{}]+)\}", "$1");
            string command = Functions.Contains(plain) ? "\\" + plain : @"\operatorname{" + plain + "}";
            return command + (argument.Length > 0 ? " " + argument : string.Empty);
        }

        private static string RenderLimit(XElement limit, bool upper)
        {
            string basis = RenderChild(limit, "e");
            string value = RenderChild(limit, "lim");
            if (basis.StartsWith("\\", StringComparison.Ordinal) && !basis.Contains(" "))
                return basis + (upper ? "^{" : "_{") + value + "}";
            return (upper ? @"\overset{" : @"\underset{") + value + "}{" + basis + "}";
        }

        private static string RenderChild(XElement parent, string name)
        {
            XElement child = parent.Element(M + name);
            return child == null ? string.Empty : RenderChildren(child);
        }

        private static string RenderChildren(XElement parent)
        {
            if (parent == null) return string.Empty;
            StringBuilder value = new StringBuilder();
            foreach (XElement child in parent.Elements()) value.Append(Render(child));
            return value.ToString();
        }

        private static bool IsPropertyElement(string name)
        {
            return name.EndsWith("Pr", StringComparison.Ordinal) || name == "ctrlPr" || name == "argPr";
        }

        private static string AttributeValue(XElement element, string localName)
        {
            if (element == null) return string.Empty;
            XAttribute attribute = element.Attributes().FirstOrDefault(item => item.Name.LocalName == localName);
            return attribute == null ? string.Empty : attribute.Value;
        }

        private static string GroupBasis(string basis)
        {
            return basis.Length <= 1 || (basis.StartsWith("\\", StringComparison.Ordinal) && !basis.Contains(" "))
                ? basis : "{" + basis + "}";
        }

        private static string NaryCommand(string value)
        {
            if (value == "∫") return @"\int";
            if (value == "∬") return @"\iint";
            if (value == "∭") return @"\iiint";
            if (value == "∮") return @"\oint";
            if (value == "∏") return @"\prod";
            if (value == "∐") return @"\coprod";
            if (value == "⋃") return @"\bigcup";
            if (value == "⋂") return @"\bigcap";
            return @"\sum";
        }

        private static string AccentCommand(string value)
        {
            if (value == "̅") return @"\bar";
            if (value == "̃") return @"\tilde";
            if (value == "⃗") return @"\vec";
            if (value == "⃖") return @"\overleftarrow";
            if (value == "⃡") return @"\overleftrightarrow";
            if (value == "̇") return @"\dot";
            if (value == "̈") return @"\ddot";
            if (value == "̌") return @"\check";
            if (value == "̆") return @"\breve";
            if (value == "̊") return @"\mathring";
            return @"\hat";
        }

        private static string MatrixEnvironment(string begin, string end)
        {
            if (begin == "(" && end == ")") return "pmatrix";
            if (begin == "[" && end == "]") return "bmatrix";
            if (begin == "{" && end == "}") return "Bmatrix";
            if (begin == "|" && end == "|") return "vmatrix";
            if (begin == "‖" && end == "‖") return "Vmatrix";
            if (begin == "{" && (end == "." || end.Length == 0)) return "cases";
            return null;
        }

        private static string DelimiterCommand(string value, bool opening)
        {
            if (value == "⌈") return opening ? @"\lceil" : @"\rceil";
            if (value == "⌉") return @"\rceil";
            if (value == "⌊") return opening ? @"\lfloor" : @"\rfloor";
            if (value == "⌋") return @"\rfloor";
            if (value == "⟨") return @"\langle";
            if (value == "⟩") return @"\rangle";
            if (value == "‖") return @"\|";
            if (value == "{") return @"\{";
            if (value == "}") return @"\}";
            return string.IsNullOrEmpty(value) ? "." : value;
        }

        private static bool NeedsTextStyle(string text)
        {
            return text.Any(char.IsLetter) && text.Length > 1;
        }

        private static string EscapeText(string text)
        {
            StringBuilder result = new StringBuilder();
            foreach (char character in text)
            {
                string replacement;
                if (Symbols.TryGetValue(character.ToString(), out replacement))
                {
                    result.Append(replacement).Append(' ');
                    continue;
                }
                if (character == '\\') result.Append(@"\backslash ");
                else if (character == '{') result.Append(@"\{");
                else if (character == '}') result.Append(@"\}");
                else if (character == '#') result.Append(@"\#");
                else if (character == '%') result.Append(@"\%");
                else if (character == '&') result.Append(@"\&");
                else if (character == '_') result.Append(@"\_");
                else if (character == '^') result.Append(@"\hat{} ");
                else if (character == '~') result.Append(@"\sim ");
                else result.Append(character);
            }
            // Preserve a trailing separator after a control word. Word frequently splits
            // symbols and following letters into separate runs; trimming here would join
            // "\in" + "R" into the unrelated command "\inR".
            return Regex.Replace(result.ToString(), @"\s+", " ").TrimStart();
        }
    }
}

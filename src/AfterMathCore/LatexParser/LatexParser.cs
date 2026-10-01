using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace AfterMathCore.LatexParser
{
    /// <summary>
    /// Parses the practical LaTeX subset used in research papers. This is a C# port and
    /// extension of AfterMath's MIT-licensed latex_to_omml parser, not a full TeX engine.
    /// </summary>
    public sealed class LatexParser
    {
        private static readonly IDictionary<string, string> Greek = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            { "alpha", "α" }, { "beta", "β" }, { "gamma", "γ" }, { "delta", "δ" },
            { "epsilon", "ε" }, { "varepsilon", "ϵ" }, { "zeta", "ζ" }, { "eta", "η" },
            { "theta", "θ" }, { "vartheta", "ϑ" }, { "iota", "ι" }, { "kappa", "κ" },
            { "lambda", "λ" }, { "mu", "μ" }, { "nu", "ν" }, { "xi", "ξ" },
            { "pi", "π" }, { "varpi", "ϖ" }, { "rho", "ρ" }, { "varrho", "ϱ" },
            { "sigma", "σ" }, { "varsigma", "ς" }, { "tau", "τ" }, { "upsilon", "υ" },
            { "phi", "φ" }, { "varphi", "ϕ" }, { "chi", "χ" }, { "psi", "ψ" }, { "omega", "ω" },
            { "Gamma", "Γ" }, { "Delta", "Δ" }, { "Theta", "Θ" }, { "Lambda", "Λ" },
            { "Xi", "Ξ" }, { "Pi", "Π" }, { "Sigma", "Σ" }, { "Upsilon", "Υ" },
            { "Phi", "Φ" }, { "Psi", "Ψ" }, { "Omega", "Ω" }
        };

        private static readonly IDictionary<string, string> Symbols = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            { "in", "∈" }, { "notin", "∉" }, { "ni", "∋" }, { "times", "×" }, { "cdot", "·" },
            { "pm", "±" }, { "mp", "∓" }, { "le", "≤" }, { "leq", "≤" }, { "ge", "≥" },
            { "geq", "≥" }, { "neq", "≠" }, { "approx", "≈" }, { "equiv", "≡" },
            { "to", "→" }, { "rightarrow", "→" }, { "leftarrow", "←" }, { "leftrightarrow", "↔" },
            { "Rightarrow", "⇒" }, { "Leftarrow", "⇐" }, { "infty", "∞" }, { "partial", "∂" },
            { "nabla", "∇" }, { "forall", "∀" }, { "exists", "∃" }, { "propto", "∝" },
            { "parallel", "∥" }, { "perp", "⊥" }, { "subset", "⊂" }, { "subseteq", "⊆" },
            { "supset", "⊃" }, { "supseteq", "⊇" }, { "cup", "∪" }, { "cap", "∩" },
            { "ldots", "…" }, { "cdots", "⋯" }, { "vdots", "⋮" }, { "ddots", "⋱" }
        };

        private static readonly IDictionary<string, string> Blackboard = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            { "R", "ℝ" }, { "N", "ℕ" }, { "Z", "ℤ" }, { "Q", "ℚ" }, { "C", "ℂ" }
        };

        private static readonly ISet<string> UprightFunctions = new HashSet<string>(StringComparer.Ordinal)
        {
            "sin", "cos", "tan", "log", "ln", "exp", "min", "max", "det", "dim", "ker", "rank", "tr"
        };

        private readonly string _source;
        private int _position;

        public LatexParser(string source)
        {
            _source = LatexTokenizer.NormalizeSource(source);
        }

        public LatexNode Parse()
        {
            if (_source.Length == 0)
            {
                throw new LatexParseException("公式为空。");
            }

            LatexNode node = ParseUntil('\0');
            SkipSpaces();
            if (_position != _source.Length)
            {
                throw Error("存在无法解析的内容");
            }

            return node;
        }

        private LatexNode ParseUntil(char stop)
        {
            List<LatexNode> children = new List<LatexNode>();
            StringBuilder text = new StringBuilder();
            while (_position < _source.Length)
            {
                char current = _source[_position];
                if ((stop != '\0' && current == stop) || current == '}')
                {
                    break;
                }

                if (current == '\\')
                {
                    FlushText(children, text);
                    children.Add(ParseCommand());
                    continue;
                }

                if (current == '^' || current == '_')
                {
                    FlushText(children, text);
                    if (children.Count == 0)
                    {
                        throw Error("上下标缺少主体");
                    }

                    char marker = current;
                    _position++;
                    LatexNode script = ParseAtom();
                    LatexNode basis = children[children.Count - 1];
                    children.RemoveAt(children.Count - 1);
                    children.Add(AttachScript(basis, marker, script));
                    continue;
                }

                if (current == '{')
                {
                    FlushText(children, text);
                    children.Add(ParseGroup());
                    continue;
                }

                if (current == '&')
                {
                    text.Append(' ');
                    _position++;
                    continue;
                }

                text.Append(current);
                _position++;
            }

            FlushText(children, text);
            return LatexNode.Sequence(children);
        }

        private LatexNode ParseAtom()
        {
            SkipSpaces();
            if (_position >= _source.Length)
            {
                throw Error("缺少参数");
            }

            char current = _source[_position];
            if (current == '{') return ParseGroup();
            if (current == '\\') return ParseCommand();
            _position++;
            return LatexNode.Text(current.ToString());
        }

        private LatexNode ParseGroup()
        {
            if (_position >= _source.Length || _source[_position] != '{')
            {
                throw Error("此处应为大括号参数");
            }

            _position++;
            LatexNode node = ParseUntil('\0');
            if (_position >= _source.Length || _source[_position] != '}')
            {
                throw Error("大括号未闭合");
            }
            _position++;
            return node;
        }

        private LatexNode ParseCommand()
        {
            _position++;
            if (_position >= _source.Length) return LatexNode.Text("\\");

            char current = _source[_position];
            if ("\\{}[](),;:! ".IndexOf(current) >= 0)
            {
                _position++;
                if (",;:! ".IndexOf(current) >= 0) return LatexNode.Text(" ");
                return LatexNode.Text(current.ToString());
            }

            int start = _position;
            while (_position < _source.Length && char.IsLetter(_source[_position])) _position++;
            string name = _source.Substring(start, _position - start);
            if (name.Length == 0)
            {
                return LatexNode.Text(_source[_position++].ToString());
            }

            if (name == "left" || name == "right")
            {
                SkipSpaces();
                if (_position >= _source.Length) return LatexNode.Text(string.Empty);
                string delimiter = ReadDelimiter();
                return LatexNode.Text(delimiter == "." ? string.Empty : delimiter);
            }

            if (name == "frac")
            {
                LatexNode fraction = new LatexNode(LatexNodeKind.Fraction);
                fraction.Children.Add(ParseRequiredGroup(name));
                fraction.Children.Add(ParseRequiredGroup(name));
                return fraction;
            }

            if (name == "sqrt")
            {
                LatexNode radical = new LatexNode(LatexNodeKind.Radical);
                SkipSpaces();
                if (_position < _source.Length && _source[_position] == '[')
                {
                    _position++;
                    int degreeStart = _position;
                    while (_position < _source.Length && _source[_position] != ']') _position++;
                    if (_position >= _source.Length) throw Error("根指数未闭合");
                    radical.Slots["degree"] = new LatexParser(_source.Substring(degreeStart, _position - degreeStart)).Parse();
                    _position++;
                }
                radical.Children.Add(ParseRequiredGroup(name));
                return radical;
            }

            if (name == "hat" || name == "bar" || name == "tilde" || name == "vec" || name == "dot" || name == "ddot")
            {
                LatexNode accent = new LatexNode(LatexNodeKind.Accent);
                accent.Value = name;
                accent.Children.Add(ParseRequiredGroup(name));
                return accent;
            }

            if (name == "mathbf" || name == "boldsymbol" || name == "mathrm" || name == "mathit" || name == "operatorname" || name == "text")
            {
                LatexNode styled = new LatexNode(LatexNodeKind.Styled);
                styled.Variant = name == "mathbf" ? MathVariant.Bold :
                    name == "boldsymbol" ? MathVariant.BoldItalic :
                    name == "mathit" ? MathVariant.Italic : MathVariant.Plain;
                styled.Children.Add(ParseRequiredGroup(name));
                return styled;
            }

            if (name == "mathbb")
            {
                string plain = PlainText(ParseRequiredGroup(name));
                StringBuilder converted = new StringBuilder();
                foreach (char letter in plain)
                {
                    string replacement;
                    converted.Append(Blackboard.TryGetValue(letter.ToString(), out replacement) ? replacement : letter.ToString());
                }
                return LatexNode.Text(converted.ToString());
            }

            if (name == "sum" || name == "int" || name == "prod" || name == "bigcup" || name == "bigcap")
            {
                LatexNode nary = new LatexNode(LatexNodeKind.Nary);
                nary.Value = name == "sum" ? "∑" : name == "int" ? "∫" : name == "prod" ? "∏" : name == "bigcup" ? "⋃" : "⋂";
                return nary;
            }

            if (name == "begin")
            {
                return ParseEnvironment(PlainText(ParseRequiredGroup(name)));
            }

            string symbol;
            if (Greek.TryGetValue(name, out symbol)) return LatexNode.Text(symbol);
            if (Symbols.TryGetValue(name, out symbol)) return LatexNode.Text(symbol);
            if (UprightFunctions.Contains(name))
            {
                LatexNode function = new LatexNode(LatexNodeKind.Styled);
                function.Variant = MathVariant.Plain;
                function.Children.Add(LatexNode.Text(name));
                return function;
            }

            throw Error("暂不支持 LaTeX 命令 \\" + name);
        }

        private LatexNode ParseEnvironment(string environment)
        {
            string closing = "\\end{" + environment + "}";
            int end = _source.IndexOf(closing, _position, StringComparison.Ordinal);
            if (end < 0) throw Error("环境 " + environment + " 未闭合");
            string raw = _source.Substring(_position, end - _position);
            _position = end + closing.Length;

            if (environment != "matrix" && environment != "pmatrix" && environment != "bmatrix" &&
                environment != "cases" && environment != "aligned" && environment != "align")
            {
                throw Error("暂不支持环境 " + environment);
            }

            LatexNode matrix = new LatexNode(LatexNodeKind.Matrix);
            foreach (string rowText in SplitRows(raw))
            {
                LatexNode row = new LatexNode(LatexNodeKind.Sequence);
                foreach (string cell in SplitCells(rowText))
                {
                    row.Children.Add(string.IsNullOrWhiteSpace(cell) ? LatexNode.Text(string.Empty) : new LatexParser(cell.Trim()).Parse());
                }
                matrix.Children.Add(row);
            }

            if (environment == "matrix" || environment == "aligned" || environment == "align") return matrix;
            LatexNode delimiterNode = new LatexNode(LatexNodeKind.Delimiter);
            delimiterNode.Value = environment == "pmatrix" ? "(" : environment == "bmatrix" ? "[" : "{";
            delimiterNode.Slots["end"] = LatexNode.Text(environment == "pmatrix" ? ")" : environment == "bmatrix" ? "]" : string.Empty);
            delimiterNode.Children.Add(matrix);
            return delimiterNode;
        }

        private LatexNode ParseRequiredGroup(string command)
        {
            SkipSpaces();
            if (_position >= _source.Length || _source[_position] != '{')
            {
                throw Error("\\" + command + " 需要大括号参数");
            }
            return ParseGroup();
        }

        private LatexNode AttachScript(LatexNode basis, char marker, LatexNode script)
        {
            if (basis.Kind == LatexNodeKind.Nary)
            {
                basis.Slots[marker == '^' ? "sup" : "sub"] = script;
                return basis;
            }

            LatexNode result = basis.Kind == LatexNodeKind.Script ? basis : new LatexNode(LatexNodeKind.Script);
            if (basis.Kind != LatexNodeKind.Script) result.Children.Add(basis);
            result.Slots[marker == '^' ? "sup" : "sub"] = script;
            return result;
        }

        private string ReadDelimiter()
        {
            if (_source[_position] != '\\') return _source[_position++].ToString();
            _position++;
            int start = _position;
            while (_position < _source.Length && char.IsLetter(_source[_position])) _position++;
            string name = _source.Substring(start, _position - start);
            if (name == "langle") return "⟨";
            if (name == "rangle") return "⟩";
            if (name == "lbrace") return "{";
            if (name == "rbrace") return "}";
            if (name == "vert" || name == "Vert") return "|";
            return name;
        }

        private static IList<string> SplitRows(string raw)
        {
            string normalized = raw.Replace("\r\n", "\n").Replace("\r", "\n");
            normalized = Regex.Replace(normalized, @"\\\\", "\n");
            string[] pieces = normalized.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
            List<string> rows = new List<string>();
            foreach (string piece in pieces) if (!string.IsNullOrWhiteSpace(piece)) rows.Add(piece.Trim());
            return rows;
        }

        private static IList<string> SplitCells(string row)
        {
            List<string> cells = new List<string>();
            int level = 0;
            int start = 0;
            for (int i = 0; i < row.Length; i++)
            {
                if (row[i] == '{') level++;
                else if (row[i] == '}') level--;
                else if (row[i] == '&' && level == 0)
                {
                    cells.Add(row.Substring(start, i - start));
                    start = i + 1;
                }
            }
            cells.Add(row.Substring(start));
            return cells;
        }

        private static string PlainText(LatexNode node)
        {
            if (node.Kind == LatexNodeKind.Text) return node.Value ?? string.Empty;
            if (node.Kind == LatexNodeKind.Sequence)
            {
                StringBuilder result = new StringBuilder();
                foreach (LatexNode child in node.Children) result.Append(PlainText(child));
                return result.ToString();
            }
            throw new LatexParseException("此参数应为纯文本。");
        }

        private void SkipSpaces()
        {
            while (_position < _source.Length && char.IsWhiteSpace(_source[_position])) _position++;
        }

        private static void FlushText(ICollection<LatexNode> children, StringBuilder text)
        {
            if (text.Length == 0) return;
            children.Add(LatexNode.Text(text.ToString()));
            text.Clear();
        }

        private LatexParseException Error(string message)
        {
            return new LatexParseException(message + "（位置 " + _position + "）。");
        }
    }
}

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
            { "omicron", "ο" }, { "varkappa", "ϰ" }, { "digamma", "ϝ" },
            { "Gamma", "Γ" }, { "Delta", "Δ" }, { "Theta", "Θ" }, { "Lambda", "Λ" },
            { "Xi", "Ξ" }, { "Pi", "Π" }, { "Sigma", "Σ" }, { "Upsilon", "Υ" },
            { "Phi", "Φ" }, { "Psi", "Ψ" }, { "Omega", "Ω" }, { "varTheta", "ϴ" }
        };

        private static readonly IDictionary<string, string> Symbols = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            { "in", "∈" }, { "notin", "∉" }, { "ni", "∋" }, { "times", "×" }, { "cdot", "·" },
            { "pm", "±" }, { "mp", "∓" }, { "le", "≤" }, { "leq", "≤" }, { "ge", "≥" },
            { "leqslant", "⩽" }, { "geq", "≥" }, { "geqslant", "⩾" }, { "neq", "≠" }, { "ne", "≠" },
            { "lt", "<" }, { "gt", ">" }, { "approx", "≈" }, { "sim", "∼" }, { "simeq", "≃" },
            { "cong", "≅" }, { "equiv", "≡" }, { "ll", "≪" }, { "gg", "≫" },
            { "prec", "≺" }, { "preceq", "⪯" }, { "succ", "≻" }, { "succeq", "⪰" },
            { "to", "→" }, { "rightarrow", "→" }, { "leftarrow", "←" }, { "leftrightarrow", "↔" },
            { "gets", "←" }, { "mapsto", "↦" }, { "hookrightarrow", "↪" }, { "hookleftarrow", "↩" },
            { "uparrow", "↑" }, { "downarrow", "↓" }, { "updownarrow", "↕" },
            { "Rightarrow", "⇒" }, { "Leftarrow", "⇐" }, { "Leftrightarrow", "⇔" },
            { "longrightarrow", "⟶" }, { "longleftarrow", "⟵" }, { "longleftrightarrow", "⟷" },
            { "Longrightarrow", "⟹" }, { "Longleftarrow", "⟸" }, { "Longleftrightarrow", "⟺" },
            { "infty", "∞" }, { "partial", "∂" },
            { "nabla", "∇" }, { "forall", "∀" }, { "exists", "∃" }, { "propto", "∝" },
            { "parallel", "∥" }, { "perp", "⊥" }, { "subset", "⊂" }, { "subseteq", "⊆" },
            { "supset", "⊃" }, { "supseteq", "⊇" }, { "nsubseteq", "⊈" }, { "nsupseteq", "⊉" },
            { "cup", "∪" }, { "cap", "∩" }, { "setminus", "∖" }, { "smallsetminus", "∖" },
            { "emptyset", "∅" }, { "varnothing", "∅" }, { "land", "∧" }, { "wedge", "∧" },
            { "lor", "∨" }, { "vee", "∨" }, { "neg", "¬" }, { "lnot", "¬" },
            { "implies", "⇒" }, { "iff", "⇔" }, { "therefore", "∴" }, { "because", "∵" },
            { "doteq", "≐" }, { "triangleq", "≜" }, { "coloneqq", "≔" }, { "eqqcolon", "≕" },
            { "asymp", "≍" }, { "bowtie", "⋈" }, { "models", "⊧" }, { "vdash", "⊢" },
            { "dashv", "⊣" }, { "vDash", "⊨" }, { "Vdash", "⊩" }, { "Vvdash", "⊪" },
            { "mid", "∣" }, { "nmid", "∤" }, { "nparallel", "∦" },
            { "lesssim", "≲" }, { "gtrsim", "≳" }, { "lessapprox", "⪅" }, { "gtrapprox", "⪆" },
            { "subsetneq", "⊊" }, { "supsetneq", "⊋" }, { "sqsubset", "⊏" }, { "sqsupset", "⊐" },
            { "sqsubseteq", "⊑" }, { "sqsupseteq", "⊒" },
            { "div", "÷" }, { "ast", "∗" }, { "star", "⋆" }, { "circ", "∘" }, { "bullet", "∙" },
            { "oplus", "⊕" }, { "ominus", "⊖" }, { "otimes", "⊗" }, { "oslash", "⊘" }, { "odot", "⊙" },
            { "circledast", "⊛" }, { "circledcirc", "⊚" }, { "circleddash", "⊝" },
            { "boxplus", "⊞" }, { "boxminus", "⊟" }, { "boxtimes", "⊠" }, { "boxdot", "⊡" },
            { "diamond", "⋄" }, { "bigcirc", "◯" }, { "uplus", "⊎" }, { "sqcap", "⊓" }, { "sqcup", "⊔" },
            { "amalg", "⨿" }, { "wr", "≀" }, { "dagger", "†" }, { "ddagger", "‡" },
            { "triangleleft", "◁" }, { "triangleright", "▷" }, { "unlhd", "⊴" }, { "unrhd", "⊵" },
            { "triangle", "△" }, { "angle", "∠" }, { "measuredangle", "∡" },
            { "ell", "ℓ" }, { "hbar", "ℏ" }, { "Re", "ℜ" }, { "Im", "ℑ" },
            { "prime", "′" }, { "degree", "°" }, { "top", "⊤" }, { "bot", "⊥" },
            { "aleph", "ℵ" }, { "beth", "ℶ" }, { "gimel", "ℷ" }, { "daleth", "ℸ" },
            { "imath", "ı" }, { "jmath", "ȷ" }, { "wp", "℘" }, { "mho", "℧" }, { "eth", "ð" },
            { "complement", "∁" }, { "surd", "√" }, { "backslash", "∖" }, { "colon", ":" },
            { "flat", "♭" }, { "natural", "♮" }, { "sharp", "♯" }, { "checkmark", "✓" },
            { "square", "□" }, { "Box", "□" }, { "blacksquare", "■" }, { "lozenge", "◊" },
            { "clubsuit", "♣" }, { "diamondsuit", "♢" }, { "heartsuit", "♡" }, { "spadesuit", "♠" },
            { "nearrow", "↗" }, { "searrow", "↘" }, { "swarrow", "↙" }, { "nwarrow", "↖" },
            { "rightharpoonup", "⇀" }, { "rightharpoondown", "⇁" }, { "leftharpoonup", "↼" },
            { "leftharpoondown", "↽" }, { "rightleftharpoons", "⇌" }, { "leftrightharpoons", "⇋" },
            { "leadsto", "⇝" }, { "twoheadrightarrow", "↠" }, { "twoheadleftarrow", "↞" },
            { "ldots", "…" }, { "dots", "…" }, { "dotsc", "…" }, { "dotsb", "⋯" },
            { "dotsm", "⋯" }, { "dotsi", "⋯" }, { "dotso", "…" }, { "cdots", "⋯" },
            { "vdots", "⋮" }, { "ddots", "⋱" }
        };

        private static readonly IDictionary<string, string> Blackboard = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            { "C", "ℂ" }, { "H", "ℍ" }, { "N", "ℕ" }, { "P", "ℙ" },
            { "Q", "ℚ" }, { "R", "ℝ" }, { "Z", "ℤ" }
        };

        private static readonly IDictionary<string, string> ScriptLetters = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            { "B", "ℬ" }, { "E", "ℰ" }, { "F", "ℱ" }, { "H", "ℋ" }, { "I", "ℐ" },
            { "L", "ℒ" }, { "M", "ℳ" }, { "R", "ℛ" }, { "e", "ℯ" }, { "g", "ℊ" }, { "o", "ℴ" }
        };

        private static readonly IDictionary<string, string> FrakturLetters = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            { "C", "ℭ" }, { "H", "ℌ" }, { "I", "ℑ" }, { "R", "ℜ" }, { "Z", "ℨ" }
        };

        private static readonly ISet<string> UprightFunctions = new HashSet<string>(StringComparer.Ordinal)
        {
            "sin", "cos", "tan", "cot", "sec", "csc", "arcsin", "arccos", "arctan",
            "sinh", "cosh", "tanh", "coth", "log", "ln", "lg", "exp",
            "min", "max", "sup", "inf", "lim", "limsup", "liminf",
            "det", "dim", "ker", "rank", "tr", "trace", "diag", "gcd", "lcm", "Pr", "arg", "argmin", "argmax", "hom", "sgn"
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
            return ParseUntil(stop, false);
        }

        private LatexNode ParseUntil(char stop, bool stopAtRightDelimiter)
        {
            List<LatexNode> children = new List<LatexNode>();
            StringBuilder text = new StringBuilder();
            while (_position < _source.Length)
            {
                if (stopAtRightDelimiter && IsCommandAtCurrentPosition("right"))
                {
                    break;
                }

                char current = _source[_position];
                if ((stop != '\0' && current == stop) || current == '}')
                {
                    break;
                }

                if (current == '\\')
                {
                    FlushText(children, text);
                    LatexNode commandNode = ParseCommand();
                    if (commandNode.Kind != LatexNodeKind.Text || !string.IsNullOrEmpty(commandNode.Value))
                    {
                        children.Add(commandNode);
                    }
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
                if (current == '!') return LatexNode.Text(string.Empty);
                if (",;: ".IndexOf(current) >= 0) return LatexNode.Text(" ");
                return LatexNode.Text(current.ToString());
            }

            int start = _position;
            while (_position < _source.Length && char.IsLetter(_source[_position])) _position++;
            string name = _source.Substring(start, _position - start);
            if (name.Length == 0)
            {
                return LatexNode.Text(_source[_position++].ToString());
            }

            if (name == "left")
            {
                return ParseLeftRight();
            }

            if (name == "right")
            {
                SkipSpaces();
                return LatexNode.Text(_position >= _source.Length ? string.Empty : ReadDelimiter());
            }

            if (name == "middle")
            {
                SkipSpaces();
                return LatexNode.Text(_position >= _source.Length ? string.Empty : ReadDelimiter());
            }

            if (name == "frac" || name == "dfrac" || name == "tfrac" || name == "cfrac" ||
                name == "sfrac" || name == "nicefrac")
            {
                LatexNode fraction = new LatexNode(LatexNodeKind.Fraction);
                fraction.Children.Add(ParseRequiredArgument(name));
                fraction.Children.Add(ParseRequiredArgument(name));
                return fraction;
            }

            if (name == "binom" || name == "dbinom" || name == "tbinom")
            {
                LatexNode fraction = new LatexNode(LatexNodeKind.Fraction);
                fraction.Value = "noBar";
                fraction.Children.Add(ParseRequiredArgument(name));
                fraction.Children.Add(ParseRequiredArgument(name));
                return CreateDelimiter("(", ")", fraction);
            }

            if (name == "genfrac")
            {
                string beginning = PlainText(ParseRequiredGroup(name));
                string ending = PlainText(ParseRequiredGroup(name));
                string thickness = PlainText(ParseRequiredGroup(name));
                ParseRequiredGroup(name); // style selector; Word chooses its own native equation size.
                LatexNode fraction = new LatexNode(LatexNodeKind.Fraction);
                if (thickness == "0" || thickness == "0pt") fraction.Value = "noBar";
                fraction.Children.Add(ParseRequiredArgument(name));
                fraction.Children.Add(ParseRequiredArgument(name));
                return string.IsNullOrEmpty(beginning) && string.IsNullOrEmpty(ending)
                    ? fraction : CreateDelimiter(beginning, ending, fraction);
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
                radical.Children.Add(ParseRequiredArgument(name));
                return radical;
            }

            if (name == "hat" || name == "widehat" || name == "bar" || name == "tilde" || name == "widetilde" ||
                name == "vec" || name == "overrightarrow" || name == "overleftarrow" || name == "overleftrightarrow" ||
                name == "dot" || name == "ddot" || name == "check" || name == "breve" || name == "acute" ||
                name == "grave" || name == "mathring")
            {
                LatexNode accent = new LatexNode(LatexNodeKind.Accent);
                accent.Value = name == "widehat" ? "hat" : name == "widetilde" ? "tilde" : name;
                accent.Children.Add(ParseRequiredArgument(name));
                return accent;
            }

            if (name == "overline" || name == "underline")
            {
                LatexNode bar = new LatexNode(LatexNodeKind.Bar);
                bar.Value = name == "overline" ? "top" : "bot";
                bar.Children.Add(ParseRequiredArgument(name));
                return bar;
            }

            if (name == "overbrace" || name == "underbrace")
            {
                LatexNode groupCharacter = new LatexNode(LatexNodeKind.GroupCharacter);
                groupCharacter.Value = name == "overbrace" ? "top" : "bot";
                groupCharacter.Children.Add(ParseRequiredArgument(name));
                return groupCharacter;
            }

            if (name == "overset" || name == "stackrel" || name == "underset")
            {
                LatexNode limit = new LatexNode(name == "underset" ? LatexNodeKind.LimitLower : LatexNodeKind.LimitUpper);
                limit.Slots["limit"] = ParseRequiredArgument(name);
                limit.Children.Add(ParseRequiredArgument(name));
                return limit;
            }

            if (name == "prescript")
            {
                LatexNode script = new LatexNode(LatexNodeKind.PreScript);
                script.Slots["sup"] = ParseRequiredArgument(name);
                script.Slots["sub"] = ParseRequiredArgument(name);
                script.Children.Add(ParseRequiredArgument(name));
                return script;
            }

            if (name == "xrightarrow" || name == "xleftarrow" || name == "xleftrightarrow" ||
                name == "xRightarrow" || name == "xLeftarrow" || name == "xLeftrightarrow" || name == "xmapsto")
            {
                LatexNode below = ParseOptionalBracket();
                LatexNode above = ParseRequiredArgument(name);
                string arrow = name == "xrightarrow" ? "→" : name == "xleftarrow" ? "←" :
                    name == "xleftrightarrow" ? "↔" : name == "xRightarrow" ? "⇒" :
                    name == "xLeftarrow" ? "⇐" : name == "xLeftrightarrow" ? "⇔" : "↦";
                LatexNode upper = CreateLimit(LatexNodeKind.LimitUpper, LatexNode.Text(arrow), above);
                return below == null ? upper : CreateLimit(LatexNodeKind.LimitLower, upper, below);
            }

            if (name == "boxed")
            {
                LatexNode border = new LatexNode(LatexNodeKind.BorderBox);
                border.Children.Add(ParseRequiredArgument(name));
                return border;
            }

            if (name == "mathbf" || name == "bm" || name == "boldsymbol" || name == "symbf" || name == "pmb" || name == "mathbfit" ||
                name == "mathrm" || name == "textrm" || name == "textnormal" || name == "textup" || name == "mbox" ||
                name == "mathit" || name == "textit" || name == "textbf" || name == "mathsf" || name == "textsf" ||
                name == "mathtt" || name == "texttt" ||
                name == "mathnormal" || name == "operatorname" || name == "text")
            {
                if (name == "operatorname" && _position < _source.Length && _source[_position] == '*') _position++;
                LatexNode styled = new LatexNode(LatexNodeKind.Styled);
                styled.Variant = name == "mathbf" || name == "textbf" || name == "pmb" ? MathVariant.Bold :
                    name == "boldsymbol" || name == "bm" || name == "symbf" || name == "mathbfit" ? MathVariant.BoldItalic :
                    name == "mathit" || name == "textit" ? MathVariant.Italic : MathVariant.Plain;
                styled.Children.Add(name == "operatorname" || name == "text" || name == "mbox"
                    ? ParseRequiredGroup(name) : ParseRequiredArgument(name));
                return styled;
            }

            if (name == "mathbb" || name == "mathbbm" || name == "mathds" || name == "mathcal" || name == "mathscr" || name == "mathfrak")
            {
                string plain = PlainText(ParseRequiredArgument(name));
                string alphabet = name == "mathbbm" || name == "mathds" ? "mathbb" : name;
                return LatexNode.Text(ConvertMathematicalAlphabet(plain, alphabet));
            }

            if (name == "abs" || name == "norm" || name == "floor" || name == "ceil" || name == "avg" ||
                name == "bra" || name == "ket" || name == "braket" || name == "expval")
            {
                // Several paper toolchains emit \ceil x\rceil and \floor x\rfloor instead
                // of the conventional \lceil x\rceil / \lfloor x\rfloor pair.  Keep
                // supporting the braced semantic macros, but treat an unbraced opening
                // command as a delimiter alias so the rest of the expression is not
                // incorrectly consumed as a one-token macro argument.
                if (!HasGroupArgument() && (name == "ceil" || name == "floor"))
                {
                    return LatexNode.Text(name == "ceil" ? "⌈" : "⌊");
                }

                LatexNode argument = ParseRequiredArgument(name);
                if (name == "norm") return CreateDelimiter("‖", "‖", argument);
                if (name == "floor") return CreateDelimiter("⌊", "⌋", argument);
                if (name == "ceil") return CreateDelimiter("⌈", "⌉", argument);
                if (name == "avg" || name == "braket" || name == "expval") return CreateDelimiter("⟨", "⟩", argument);
                if (name == "bra") return CreateDelimiter("⟨", "|", argument);
                if (name == "ket") return CreateDelimiter("|", "⟩", argument);
                return CreateDelimiter("|", "|", argument);
            }

            if (name == "commutator" || name == "anticommutator" || name == "poissonbracket")
            {
                LatexNode first = ParseRequiredArgument(name);
                LatexNode second = ParseRequiredArgument(name);
                LatexNode content = LatexNode.Sequence(new List<LatexNode> { first, LatexNode.Text(","), second });
                return name == "commutator" ? CreateDelimiter("[", "]", content) : CreateDelimiter("{", "}", content);
            }

            if (name == "dv" || name == "odv" || name == "pdv")
            {
                return ParseDerivative(name);
            }

            if (name == "dd" || name == "differential")
            {
                LatexNode differential = CreateStyledText("d", MathVariant.Plain);
                return HasGroupArgument()
                    ? LatexNode.Sequence(new List<LatexNode> { differential, ParseRequiredArgument(name) })
                    : differential;
            }

            if (name == "qty")
            {
                return ParseQuantity(name);
            }

            if (name == "mathop" || name == "mathrel" || name == "mathbin" || name == "mathord" ||
                name == "mathopen" || name == "mathclose" || name == "mathpunct" || name == "mathinner" ||
                name == "ensuremath")
            {
                return ParseRequiredArgument(name);
            }

            if (name == "sum" || name == "int" || name == "iint" || name == "iiint" || name == "oint" ||
                name == "prod" || name == "coprod" || name == "bigcup" || name == "bigcap" ||
                name == "bigvee" || name == "bigwedge" || name == "biguplus")
            {
                LatexNode nary = new LatexNode(LatexNodeKind.Nary);
                nary.Value = NaryCharacter(name);
                return nary;
            }

            if (name == "substack")
            {
                return ParseSubstack(name);
            }

            if (name == "pmod")
            {
                LatexNode content = LatexNode.Sequence(new List<LatexNode>
                {
                    LatexNode.Text("mod "), ParseRequiredArgument(name)
                });
                return CreateDelimiter("(", ")", content);
            }

            if (name == "mod" || name == "bmod")
            {
                LatexNode mod = new LatexNode(LatexNodeKind.Styled);
                mod.Variant = MathVariant.Plain;
                mod.Children.Add(LatexNode.Text("mod"));
                return mod;
            }

            if (name == "quad" || name == "qquad" || name == "enspace" || name == "enskip" ||
                name == "thinspace" || name == "medspace" || name == "thickspace")
            {
                return LatexNode.Text(name == "qquad" ? "  " : " ");
            }

            if (name == "limits" || name == "nolimits" || name == "displaystyle" || name == "textstyle" ||
                name == "scriptstyle" || name == "scriptscriptstyle" || name == "rm" || name == "bf" ||
                name == "it" || name == "cal" || name == "sf" || name == "tt")
            {
                return LatexNode.Text(string.Empty);
            }

            if (name == "label" || name == "tag" || name == "notag" || name == "nonumber")
            {
                if ((name == "label" || name == "tag") && HasGroupArgument()) ParseGroup();
                return LatexNode.Text(string.Empty);
            }

            if (name == "hspace")
            {
                if (_position < _source.Length && _source[_position] == '*') _position++;
                if (HasGroupArgument()) ParseGroup();
                return LatexNode.Text(" ");
            }

            if (name == "textcolor" || name == "colorbox")
            {
                ParseRequiredGroup(name);
                return ParseRequiredArgument(name);
            }

            if (name == "color")
            {
                ParseRequiredGroup(name);
                return LatexNode.Text(string.Empty);
            }

            if (name == "multicolumn")
            {
                ParseRequiredGroup(name);
                ParseRequiredGroup(name);
                return ParseRequiredArgument(name);
            }

            if (name == "big" || name == "Big" || name == "bigg" || name == "Bigg" ||
                name == "bigl" || name == "Bigl" || name == "biggl" || name == "Biggl" ||
                name == "bigr" || name == "Bigr" || name == "biggr" || name == "Biggr" ||
                name == "bigm" || name == "Bigm" || name == "biggm" || name == "Biggm")
            {
                SkipSpaces();
                return LatexNode.Text(_position >= _source.Length ? string.Empty : ReadDelimiter());
            }

            if (name == "not")
            {
                return ParseNegatedRelation();
            }

            if (name == "phantom" || name == "hphantom" || name == "vphantom" || name == "smash")
            {
                LatexNode hidden = ParseRequiredArgument(name);
                return name == "smash" ? hidden : LatexNode.Text(" ");
            }

            if (name == "begin")
            {
                return ParseEnvironment(PlainText(ParseRequiredGroup(name)));
            }

            string symbol;
            if (Greek.TryGetValue(name, out symbol)) return LatexNode.Text(symbol);
            if (Symbols.TryGetValue(name, out symbol)) return LatexNode.Text(symbol);
            if (name == "lceil") return LatexNode.Text("⌈");
            if (name == "rceil") return LatexNode.Text("⌉");
            if (name == "lfloor") return LatexNode.Text("⌊");
            if (name == "rfloor") return LatexNode.Text("⌋");
            if (name == "langle") return LatexNode.Text("⟨");
            if (name == "rangle") return LatexNode.Text("⟩");
            if (name == "vert" || name == "lvert" || name == "rvert") return LatexNode.Text("|");
            if (name == "Vert" || name == "lVert" || name == "rVert") return LatexNode.Text("‖");
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
            string baseEnvironment = environment.EndsWith("*", StringComparison.Ordinal)
                ? environment.Substring(0, environment.Length - 1) : environment;

            if (baseEnvironment == "equation" || baseEnvironment == "displaymath")
            {
                return string.IsNullOrWhiteSpace(raw) ? LatexNode.Text(string.Empty) : new LatexParser(raw.Trim()).Parse();
            }

            ISet<string> matrixEnvironments = new HashSet<string>(StringComparer.Ordinal)
            {
                "matrix", "smallmatrix", "pmatrix", "bmatrix", "Bmatrix", "vmatrix", "Vmatrix",
                "cases", "dcases", "array", "aligned", "alignedat", "gathered", "split", "align",
                "gather", "multline", "multlined", "eqnarray"
            };
            if (!matrixEnvironments.Contains(baseEnvironment))
            {
                throw Error("暂不支持环境 " + environment);
            }

            if (baseEnvironment == "array" || baseEnvironment == "alignedat") raw = StripLeadingGroup(raw);
            raw = StripLeadingOptionalBracket(raw);
            raw = Regex.Replace(raw, @"\\(?:hline|toprule|midrule|bottomrule)\b", string.Empty);
            raw = Regex.Replace(raw, @"\\cline\s*\{[^}]*\}", string.Empty);

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

            if (baseEnvironment == "pmatrix") return CreateDelimiter("(", ")", matrix);
            if (baseEnvironment == "bmatrix") return CreateDelimiter("[", "]", matrix);
            if (baseEnvironment == "Bmatrix") return CreateDelimiter("{", "}", matrix);
            if (baseEnvironment == "vmatrix") return CreateDelimiter("|", "|", matrix);
            if (baseEnvironment == "Vmatrix") return CreateDelimiter("‖", "‖", matrix);
            if (baseEnvironment == "cases" || baseEnvironment == "dcases") return CreateDelimiter("{", string.Empty, matrix);
            return matrix;
        }

        private LatexNode ParseLeftRight()
        {
            SkipSpaces();
            if (_position >= _source.Length) throw Error("\\left 缺少定界符");
            string beginning = ReadDelimiter();
            LatexNode content = ParseUntil('\0', true);
            if (!IsCommandAtCurrentPosition("right")) throw Error("\\left 缺少对应的 \\right");
            _position += "\\right".Length;
            SkipSpaces();
            if (_position >= _source.Length) throw Error("\\right 缺少定界符");
            string ending = ReadDelimiter();
            return CreateDelimiter(beginning == "." ? string.Empty : beginning, ending == "." ? string.Empty : ending, content);
        }

        private LatexNode ParseSubstack(string command)
        {
            string raw = ReadRawRequiredGroup(command);
            LatexNode matrix = new LatexNode(LatexNodeKind.Matrix);
            foreach (string rowText in SplitRows(raw))
            {
                LatexNode row = new LatexNode(LatexNodeKind.Sequence);
                row.Children.Add(string.IsNullOrWhiteSpace(rowText) ? LatexNode.Text(string.Empty) : new LatexParser(rowText.Trim()).Parse());
                matrix.Children.Add(row);
            }
            return matrix;
        }

        private LatexNode ParseDerivative(string command)
        {
            LatexNode order = ParseOptionalBracket();
            LatexNode expression = ParseRequiredArgument(command);
            LatexNode variable = ParseRequiredArgument(command);
            string symbol = command == "pdv" ? "∂" : "d";
            LatexNode numeratorSymbol = CreateStyledText(symbol, MathVariant.Plain);
            LatexNode denominatorVariable = variable;
            if (order != null)
            {
                numeratorSymbol = AttachScript(numeratorSymbol, '^', order);
                denominatorVariable = AttachScript(variable, '^', order);
            }

            LatexNode numerator = LatexNode.Sequence(new List<LatexNode> { numeratorSymbol, expression });
            LatexNode denominator = LatexNode.Sequence(new List<LatexNode>
            {
                CreateStyledText(symbol, MathVariant.Plain), denominatorVariable
            });
            LatexNode fraction = new LatexNode(LatexNodeKind.Fraction);
            fraction.Children.Add(numerator);
            fraction.Children.Add(denominator);
            return fraction;
        }

        private LatexNode ParseQuantity(string command)
        {
            SkipSpaces();
            if (_position >= _source.Length) throw Error("\\" + command + " 缺少参数");
            if (_source[_position] == '{') return ParseGroup();

            char opening = _source[_position];
            string closings = opening == '(' ? ")" : opening == '[' ? "]" : opening == '|' ? "|" : string.Empty;
            if (closings.Length == 0) return ParseAtom();
            _position++;
            LatexNode content = ParseUntil(closings[0]);
            if (_position >= _source.Length || _source[_position] != closings[0]) throw Error("\\qty 定界符未闭合");
            _position++;
            return CreateDelimiter(opening.ToString(), closings, content);
        }

        private LatexNode ParseNegatedRelation()
        {
            SkipSpaces();
            if (_position >= _source.Length) return LatexNode.Text("¬");
            LatexNode relation;
            if (_source[_position] == '\\') relation = ParseCommand();
            else relation = LatexNode.Text(_source[_position++].ToString());
            string value = PlainText(relation);
            IDictionary<string, string> negated = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                { "=", "≠" }, { "<", "≮" }, { ">", "≯" }, { "∈", "∉" }, { "∋", "∌" },
                { "≤", "≰" }, { "≥", "≱" }, { "∼", "≁" }, { "≈", "≉" }, { "≡", "≢" },
                { "⊂", "⊄" }, { "⊃", "⊅" }, { "⊆", "⊈" }, { "⊇", "⊉" },
                { "∥", "∦" }, { "∣", "∤" }
            };
            string replacement;
            return LatexNode.Text(negated.TryGetValue(value, out replacement) ? replacement : value + "̸");
        }

        private LatexNode ParseOptionalBracket()
        {
            SkipSpaces();
            if (_position >= _source.Length || _source[_position] != '[') return null;
            _position++;
            int start = _position;
            int braceLevel = 0;
            while (_position < _source.Length)
            {
                char current = _source[_position];
                if (current == '{') braceLevel++;
                else if (current == '}') braceLevel--;
                else if (current == ']' && braceLevel == 0)
                {
                    string raw = _source.Substring(start, _position - start);
                    _position++;
                    return string.IsNullOrWhiteSpace(raw) ? LatexNode.Text(string.Empty) : new LatexParser(raw).Parse();
                }
                _position++;
            }
            throw Error("可选参数未闭合");
        }

        private static LatexNode CreateLimit(LatexNodeKind kind, LatexNode basis, LatexNode limitValue)
        {
            LatexNode limit = new LatexNode(kind);
            limit.Children.Add(basis);
            limit.Slots["limit"] = limitValue;
            return limit;
        }

        private static LatexNode CreateStyledText(string value, MathVariant variant)
        {
            LatexNode styled = new LatexNode(LatexNodeKind.Styled);
            styled.Variant = variant;
            styled.Children.Add(LatexNode.Text(value));
            return styled;
        }

        private string ReadRawRequiredGroup(string command)
        {
            SkipSpaces();
            if (_position >= _source.Length || _source[_position] != '{')
            {
                throw Error("\\" + command + " 需要大括号参数");
            }

            int contentStart = ++_position;
            int level = 1;
            while (_position < _source.Length)
            {
                char current = _source[_position];
                if (current == '\\' && _position + 1 < _source.Length &&
                    (_source[_position + 1] == '{' || _source[_position + 1] == '}'))
                {
                    _position += 2;
                    continue;
                }
                if (current == '{') level++;
                if (current == '}')
                {
                    level--;
                    if (level == 0)
                    {
                        string result = _source.Substring(contentStart, _position - contentStart);
                        _position++;
                        return result;
                    }
                }
                _position++;
            }
            throw Error("\\" + command + " 的大括号未闭合");
        }

        private LatexNode ParseRequiredArgument(string command)
        {
            SkipSpaces();
            if (_position >= _source.Length) throw Error("\\" + command + " 缺少参数");
            return ParseAtom();
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

        private bool HasGroupArgument()
        {
            SkipSpaces();
            return _position < _source.Length && _source[_position] == '{';
        }

        private bool IsCommandAtCurrentPosition(string name)
        {
            string command = "\\" + name;
            if (_position + command.Length > _source.Length ||
                string.Compare(_source, _position, command, 0, command.Length, StringComparison.Ordinal) != 0)
            {
                return false;
            }
            int next = _position + command.Length;
            return next >= _source.Length || !char.IsLetter(_source[next]);
        }

        private static LatexNode CreateDelimiter(string beginning, string ending, LatexNode content)
        {
            LatexNode delimiter = new LatexNode(LatexNodeKind.Delimiter);
            delimiter.Value = beginning;
            delimiter.Slots["end"] = LatexNode.Text(ending);
            delimiter.Children.Add(content);
            return delimiter;
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
            if (_position >= _source.Length) return string.Empty;
            if (!char.IsLetter(_source[_position]))
            {
                char escaped = _source[_position++];
                return escaped == '|' ? "‖" : escaped.ToString();
            }
            int start = _position;
            while (_position < _source.Length && char.IsLetter(_source[_position])) _position++;
            string name = _source.Substring(start, _position - start);
            if (name == "langle") return "⟨";
            if (name == "rangle") return "⟩";
            if (name == "lbrace") return "{";
            if (name == "rbrace") return "}";
            if (name == "vert" || name == "lvert" || name == "rvert") return "|";
            if (name == "Vert" || name == "lVert" || name == "rVert") return "‖";
            if (name == "lfloor") return "⌊";
            if (name == "rfloor") return "⌋";
            if (name == "lceil") return "⌈";
            if (name == "rceil") return "⌉";
            if (name == "lbrack") return "[";
            if (name == "rbrack") return "]";
            return name;
        }

        private static string NaryCharacter(string name)
        {
            if (name == "sum") return "∑";
            if (name == "int") return "∫";
            if (name == "iint") return "∬";
            if (name == "iiint") return "∭";
            if (name == "oint") return "∮";
            if (name == "prod") return "∏";
            if (name == "coprod") return "∐";
            if (name == "bigcup") return "⋃";
            if (name == "bigcap") return "⋂";
            if (name == "bigvee") return "⋁";
            if (name == "bigwedge") return "⋀";
            return "⨄";
        }

        private static string ConvertMathematicalAlphabet(string plain, string command)
        {
            StringBuilder converted = new StringBuilder();
            foreach (char letter in plain)
            {
                string replacement;
                if (command == "mathbb")
                {
                    if (Blackboard.TryGetValue(letter.ToString(), out replacement)) converted.Append(replacement);
                    else if (letter >= 'A' && letter <= 'Z') converted.Append(char.ConvertFromUtf32(0x1D538 + letter - 'A'));
                    else if (letter >= 'a' && letter <= 'z') converted.Append(char.ConvertFromUtf32(0x1D552 + letter - 'a'));
                    else if (letter >= '0' && letter <= '9') converted.Append(char.ConvertFromUtf32(0x1D7D8 + letter - '0'));
                    else converted.Append(letter);
                    continue;
                }

                if (command == "mathcal" || command == "mathscr")
                {
                    if (ScriptLetters.TryGetValue(letter.ToString(), out replacement)) converted.Append(replacement);
                    else if (letter >= 'A' && letter <= 'Z') converted.Append(char.ConvertFromUtf32(0x1D49C + letter - 'A'));
                    else if (letter >= 'a' && letter <= 'z') converted.Append(char.ConvertFromUtf32(0x1D4B6 + letter - 'a'));
                    else converted.Append(letter);
                    continue;
                }

                if (FrakturLetters.TryGetValue(letter.ToString(), out replacement)) converted.Append(replacement);
                else if (letter >= 'A' && letter <= 'Z') converted.Append(char.ConvertFromUtf32(0x1D504 + letter - 'A'));
                else if (letter >= 'a' && letter <= 'z') converted.Append(char.ConvertFromUtf32(0x1D51E + letter - 'a'));
                else converted.Append(letter);
            }
            return converted.ToString();
        }

        private static string StripLeadingGroup(string raw)
        {
            int position = 0;
            while (position < raw.Length && char.IsWhiteSpace(raw[position])) position++;
            if (position >= raw.Length || raw[position] != '{') return raw;
            int level = 0;
            for (int index = position; index < raw.Length; index++)
            {
                if (raw[index] == '{') level++;
                else if (raw[index] == '}')
                {
                    level--;
                    if (level == 0) return raw.Substring(index + 1);
                }
            }
            return raw;
        }

        private static string StripLeadingOptionalBracket(string raw)
        {
            int position = 0;
            while (position < raw.Length && char.IsWhiteSpace(raw[position])) position++;
            if (position >= raw.Length || raw[position] != '[') return raw;
            int end = raw.IndexOf(']', position + 1);
            return end < 0 ? raw : raw.Substring(end + 1);
        }

        private static IList<string> SplitRows(string raw)
        {
            string normalized = raw.Replace("\r\n", "\n").Replace("\r", "\n");
            normalized = Regex.Replace(normalized, @"\\\\(?:\s*\[[^\]]*\])?", "\n");
            normalized = Regex.Replace(normalized, @"\\cr\b", "\n");
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

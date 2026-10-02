using System;
using System.Collections.Generic;
using System.Xml.Linq;
using WordLatexAddin;

namespace AfterMathCore.Tests
{
    /// <summary>Dependency-free smoke tests so CI can validate the converter on a clean Windows runner.</summary>
    internal static class Program
    {
        private static int _failures;

        private static int Main()
        {
            OmmlGenerator generator = new OmmlGenerator();
            AssertContains(generator.Convert(@"E=mc^2+x_i"), "sSup", "superscript");
            AssertContains(generator.Convert(@"E=mc^2+x_i"), "sSub", "subscript");
            AssertContains(generator.Convert(@"\frac{a+b}{c}"), "<m:f>", "fraction");
            AssertContains(generator.Convert(@"\sqrt{x}"), "<m:rad>", "radical");
            AssertContains(generator.Convert(@"\alpha+\beta=\omega"), "α", "Greek letters");
            AssertContains(generator.Convert(@"\sum_{i=1}^{n}x_i"), "<m:nary>", "n-ary operator");
            AssertContains(generator.Convert(@"\begin{bmatrix}a & b \\ c & d\end{bmatrix}"), "<m:m>", "matrix");
            AssertContains(generator.Convert(@"\mathbf{x}"), "m:val=\"b\"", "bold vector");
            AssertContains(generator.Convert(@"\hat{\mathbf{x}}"), "<m:acc>", "accented vector");
            AssertContains(generator.Convert(@"\boldsymbol{\omega}"), "m:val=\"bi\"", "bold symbol");
            AssertContains(generator.Convert(@"\operatorname{argmin}"), "m:val=\"p\"", "operator name");
            AssertContains(generator.Convert(@"\mathrm{diag}"), "diag", "roman function");
            AssertContains(generator.Convert(@"SO(3)+SE(3)+FMCW+LiDAR+IMU"), "m:val=\"p\"", "research terms");
            AssertXml(generator.Convert(@"\left(\frac{x}{y}\right)"), "well-formed OMML");

            AssertContains(generator.Convert(@"\dfrac{1}{2}+\tfrac12"), "<m:f>", "display and text fractions");
            AssertContains(generator.Convert(@"\binom{n}{k}"), "m:val=\"noBar\"", "binomial no-bar fraction");
            AssertContains(generator.Convert(@"\sqrt[3]{x}"), "<m:deg>", "radical degree");
            AssertContains(generator.Convert(@"\left\langle x,y\right\rangle"), "m:val=\"⟨\"", "angle delimiter start");
            AssertContains(generator.Convert(@"\left\langle x,y\right\rangle"), "m:val=\"⟩\"", "angle delimiter end");
            AssertContains(generator.Convert(@"\left\|\mathbf{x}\right\|_2"), "m:val=\"‖\"", "norm delimiters");
            AssertContains(generator.Convert(@"\overline{AB}+\underline{x}"), "<m:bar>", "overline and underline");
            AssertContains(generator.Convert(@"\overset{!}{=}+\underset{x}{\min}"), "<m:limUpp>", "upper annotation");
            AssertContains(generator.Convert(@"\overset{!}{=}+\underset{x}{\min}"), "<m:limLow>", "lower annotation");
            AssertContains(generator.Convert(@"\underbrace{x_1+\cdots+x_n}_{n\text{ terms}}"), "<m:groupChr>", "underbrace");
            AssertContains(generator.Convert(@"\boxed{x+y}"), "<m:borderBox>", "boxed equation");
            AssertContains(generator.Convert(@"\iint_\Omega f\,dA+\oint_C x\,dx"), "∬", "double integral");
            AssertContains(generator.Convert(@"\iint_\Omega f\,dA+\oint_C x\,dx"), "∮", "contour integral");
            AssertContains(generator.Convert(@"\sum\limits_{\substack{i=1\\i\ne j}}^n a_i"), "<m:m>", "substack limits");
            AssertContains(generator.Convert(@"\begin{vmatrix}a&b\\c&d\end{vmatrix}"), "m:val=\"|\"", "determinant matrix");
            AssertContains(generator.Convert(@"\begin{Vmatrix}a&b\\c&d\end{Vmatrix}"), "m:val=\"‖\"", "double-bar matrix");
            AssertContains(generator.Convert(@"\begin{array}{cc}a&b\\c&d\end{array}"), "<m:m>", "array environment");
            AssertContains(generator.Convert(@"\begin{equation*}E=mc^2\end{equation*}"), "<m:sSup>", "equation environment");
            AssertContains(generator.Convert(@"\mathbb{R}^{3\times3}"), "ℝ", "blackboard alphabet");
            AssertContains(generator.Convert(@"\mathcal{L}+\mathfrak{g}"), "ℒ", "script alphabet");
            AssertContains(generator.Convert(@"\bm{\theta}+\textbf{x}"), "m:val=\"bi\"", "bold math alias");
            AssertContains(generator.Convert(@"\operatorname*{argmin}_{x}\;f(x)"), "argmin", "starred operator name");
            AssertContains(generator.Convert(@"x\leqslant y\implies y\notin\emptyset"), "⩽", "extended relations");
            AssertContains(generator.Convert(@"x\leqslant y\implies y\notin\emptyset"), "⇒", "logic implication");
            AssertContains(generator.Convert(@"f:A\mapsto B,\quad x\mapsto f(x)"), "↦", "mapsto arrow");
            AssertContains(generator.Convert(@"a\oplus b\otimes c"), "⊕", "extended operators");
            AssertContains(generator.Convert(@"\sin x+\arctan y+\limsup_{n\to\infty}a_n"), "limsup", "extended functions");
            AssertContains(generator.Convert(@"a\pmod{n}"), "mod ", "parenthesized modulo");
            AssertXml(generator.Convert(@"\left[\begin{array}{cc}\hat{x}&\dot{y}\\\vec{v}&\overline{z}\end{array}\right]"), "complex scientific equation");
            AssertContains(generator.Convert(@"\genfrac{[}{]}{0pt}{}{n}{k}"), "m:val=\"noBar\"", "generalized fraction");
            AssertContains(generator.Convert(@"\prescript{14}{6}{\mathrm{C}}"), "<m:sPre>", "prescript isotope");
            AssertContains(generator.Convert(@"A\xrightarrow[n\to\infty]{a.s.}B"), "<m:limUpp>", "annotated arrow upper label");
            AssertContains(generator.Convert(@"A\xrightarrow[n\to\infty]{a.s.}B"), "<m:limLow>", "annotated arrow lower label");
            AssertContains(generator.Convert(@"x\not\equiv y,\;x\not=0"), "≢", "negated relation command");
            AssertContains(generator.Convert(@"x\not\equiv y,\;x\not=0"), "≠", "negated literal relation");
            AssertContains(generator.Convert(@"\dv[2]{f}{x}+\pdv{g}{t}"), "<m:f>", "derivative macros");
            AssertContains(generator.Convert(@"\dv[2]{f}{x}"), "<m:sSup>", "derivative order");
            AssertContains(generator.Convert(@"\abs{x}+\norm{\mathbf{x}}+\floor{y}+\ceil{z}"), "m:val=\"⌊\"", "semantic delimiters");
            AssertContains(generator.Convert(@"h_l=w_l=\ceil L/\Delta_b\rceil"), "⌈", "bare ceil opening alias");
            AssertContains(generator.Convert(@"h_l=w_l=\ceil L/\Delta_b\rceil"), "⌉", "bare ceil closing delimiter");
            AssertContains(generator.Convert(@"D_f=\lceil L_z/\Delta_z\rceil"), "⌉", "standard ceil delimiter pair");
            AssertContains(generator.Convert(@"q=\floor x\rfloor+\lfloor y\rfloor"), "⌋", "floor delimiter aliases");
            AssertContains(generator.Convert(@"\bra{\psi}+\ket{\phi}+\braket{\phi|\psi}"), "m:val=\"⟨\"", "quantum notation");
            AssertContains(generator.Convert(@"\commutator{A}{B}+\anticommutator{A}{B}"), "m:val=\"[\"", "commutator notation");
            AssertContains(generator.Convert(@"\qty(1+\frac12)"), "<m:d>", "physics quantity delimiter");
            AssertContains(generator.Convert(@"\begin{Bmatrix*}[r]a&b\\c&d\end{Bmatrix*}"), "m:val=\"{\"", "starred brace matrix");
            AssertContains(generator.Convert(@"\begin{pmatrix*}[c]a\\b\end{pmatrix*}"), "m:val=\"(\"", "starred parenthesis matrix");
            AssertContains(generator.Convert(@"\begin{dcases*}x&x>0\\-x&x\le0\end{dcases*}"), "m:val=\"{\"", "dcases environment");
            AssertContains(generator.Convert(@"\begin{gather*}x=1\\y=2\end{gather*}"), "<m:m>", "gather environment");
            AssertContains(generator.Convert(@"\begin{matrix}\hline a&b\\[2pt]\cline{1-2}c&d\end{matrix}"), "<m:mr>", "matrix rules and row spacing");
            AssertContains(generator.Convert(@"\begin{array}{cc}\multicolumn{2}{c}{x}\\a&b\end{array}"), "x", "multicolumn content");
            AssertContains(generator.Convert(@"\textcolor{red}{x}+\color{blue}y"), "x", "ignored color with preserved content");
            AssertContains(generator.Convert(@"\overleftrightarrow{AB}+\mathring{x}"), "⃡", "extended accents");
            AssertContains(generator.Convert(@"A\rightleftharpoons B\therefore C"), "⇌", "extended arrows");
            AssertContains(generator.Convert(@"\aleph+\beth+\checkmark+\blacksquare"), "ℵ", "miscellaneous symbols");
            AssertContains(generator.Convert(@"\argmin_x f(x)+\diag(A)+\trace(A)"), "argmin", "direct research operators");
            AssertContains(generator.Convert(@"\mathbbm{1}+\mathds{R}"), "ℝ", "blackboard aliases");
            AssertXml(generator.Convert(@"\begin{aligned}R&=\exp(\hat{\boldsymbol{\omega}}\theta)\\\mathbf{t}&=\int_0^T\mathbf{v}(t)\,dt\end{aligned}"), "robotics equation");

            OmmlToLatexConverter reverse = new OmmlToLatexConverter();
            AssertContains(reverse.Convert(generator.Convert(@"\frac{a+b}{c}")), @"\frac{a+b}{c}", "OMML fraction to LaTeX");
            AssertContains(reverse.Convert(generator.Convert(@"x_i^2")), "x_{i}^{2}", "OMML scripts to LaTeX");
            AssertContains(reverse.Convert(generator.Convert(@"\sqrt[3]{x}")), @"\sqrt[3]{x}", "OMML radical to LaTeX");
            AssertContains(reverse.Convert(generator.Convert(@"\sum_{i=1}^{n}x_i")), @"\sum_{i=1}^{n}", "OMML n-ary to LaTeX");
            AssertContains(reverse.Convert(generator.Convert(@"\left(\frac{x}{y}\right)")), @"\left(\frac{x}{y}\right)", "OMML delimiters to LaTeX");
            AssertContains(reverse.Convert(generator.Convert(@"\begin{bmatrix}a&b\\c&d\end{bmatrix}")), @"\begin{bmatrix}", "OMML matrix to LaTeX");
            AssertContains(reverse.Convert(generator.Convert(@"\hat{\mathbf{x}}")), @"\hat{\mathbf{x}}", "OMML accent and bold to LaTeX");
            AssertContains(reverse.Convert(generator.Convert(@"\alpha+\beta=\omega")), @"\alpha", "OMML Greek to LaTeX");
            AssertContains(reverse.Convert(generator.Convert(@"\boxed{x+y}")), @"\boxed{x+y}", "OMML box to LaTeX");
            AssertContains(reverse.Convert(generator.Convert(@"\operatorname{argmin}_{x}f(x)")), "argmin", "OMML function to LaTeX");
            AssertContains(reverse.Convert(@"<m:oMath xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math""><m:r><m:t>∈</m:t></m:r><m:r><m:t>R</m:t></m:r></m:oMath>"), @"\in R", "OMML split-run command separator");
            AssertXml(generator.Convert(reverse.Convert(generator.Convert(@"\frac{\hat{\mathbf{x}}_i}{\sqrt{\alpha+1}}"))), "fraction round trip");
            AssertXml(generator.Convert(reverse.Convert(generator.Convert(@"\sum_{i=1}^{n}\left\|\mathbf{x}_i\right\|_2"))), "n-ary round trip");
            AssertXml(generator.Convert(reverse.Convert(generator.Convert(@"\begin{bmatrix}a&b\\c&d\end{bmatrix}"))), "matrix round trip");

            IList<LatexToken> tokens = new LatexTokenizer(@"x_i^2+\alpha").Tokenize();
            Assert(tokens.Count > 5, "tokenizer");

            LatexScanner scanner = new LatexScanner();
            IList<LatexSpan> spans = scanner.FindAll("正文 $E=mc^2$，以及\\[\\frac{a}{b}\\]。");
            Assert(spans.Count == 2, "inline and display scanning");
            Assert(!spans[0].IsDisplay && spans[1].IsDisplay, "delimiter classification");
            Assert(scanner.FindAll(@"价格 \$5，不是公式").Count == 0, "escaped dollar");

            Console.WriteLine(_failures == 0 ? "All AfterMathCore tests passed." : _failures + " test(s) failed.");
            return _failures == 0 ? 0 : 1;
        }

        private static void AssertContains(string actual, string expected, string name)
        {
            Assert(actual.IndexOf(expected, StringComparison.Ordinal) >= 0, name + " should contain " + expected);
        }

        private static void AssertXml(string xml, string name)
        {
            try
            {
                XElement.Parse(xml);
                Assert(true, name);
            }
            catch (Exception exception)
            {
                Assert(false, name + ": " + exception.Message);
            }
        }

        private static void Assert(bool condition, string name)
        {
            if (condition)
            {
                Console.WriteLine("PASS " + name);
                return;
            }
            _failures++;
            Console.Error.WriteLine("FAIL " + name);
        }
    }
}

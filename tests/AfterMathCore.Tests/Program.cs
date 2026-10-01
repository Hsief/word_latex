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

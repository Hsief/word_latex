using System;

namespace AfterMathCore.LatexParser
{
    /// <summary>Describes unsupported or malformed LaTeX input without hiding the source text.</summary>
    public sealed class LatexParseException : Exception
    {
        public LatexParseException(string message) : base(message) { }
    }
}

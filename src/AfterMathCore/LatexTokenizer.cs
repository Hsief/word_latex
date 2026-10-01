using System;
using System.Collections.Generic;

namespace AfterMathCore
{
    /// <summary>Token kinds exposed for diagnostics and future parser extensions.</summary>
    public enum LatexTokenKind
    {
        Text,
        Command,
        OpenBrace,
        CloseBrace,
        Superscript,
        Subscript,
        Alignment,
        End
    }

    /// <summary>Represents one lexical unit in a LaTeX formula.</summary>
    public sealed class LatexToken
    {
        public LatexToken(LatexTokenKind kind, string value, int position)
        {
            Kind = kind;
            Value = value;
            Position = position;
        }

        public LatexTokenKind Kind { get; private set; }
        public string Value { get; private set; }
        public int Position { get; private set; }
    }

    /// <summary>
    /// Lightweight tokenizer for the practical TeX subset used by AfterMathCore.
    /// The parser remains deliberately small instead of pretending to be a full TeX engine.
    /// </summary>
    public sealed class LatexTokenizer
    {
        private readonly string _source;
        private int _position;

        public LatexTokenizer(string source)
        {
            _source = NormalizeSource(source);
        }

        public static string NormalizeSource(string source)
        {
            if (source == null)
            {
                throw new ArgumentNullException("source");
            }

            return source.Replace("\r\n", "\n").Replace("\r", "\n").Trim();
        }

        public IList<LatexToken> Tokenize()
        {
            List<LatexToken> tokens = new List<LatexToken>();
            while (_position < _source.Length)
            {
                int start = _position;
                char current = _source[_position++];
                switch (current)
                {
                    case '{': tokens.Add(new LatexToken(LatexTokenKind.OpenBrace, "{", start)); break;
                    case '}': tokens.Add(new LatexToken(LatexTokenKind.CloseBrace, "}", start)); break;
                    case '^': tokens.Add(new LatexToken(LatexTokenKind.Superscript, "^", start)); break;
                    case '_': tokens.Add(new LatexToken(LatexTokenKind.Subscript, "_", start)); break;
                    case '&': tokens.Add(new LatexToken(LatexTokenKind.Alignment, "&", start)); break;
                    case '\\':
                        tokens.Add(ReadCommand(start));
                        break;
                    default:
                        tokens.Add(ReadText(start));
                        break;
                }
            }

            tokens.Add(new LatexToken(LatexTokenKind.End, string.Empty, _source.Length));
            return tokens;
        }

        private LatexToken ReadCommand(int start)
        {
            if (_position >= _source.Length)
            {
                return new LatexToken(LatexTokenKind.Text, "\\", start);
            }

            if (!char.IsLetter(_source[_position]))
            {
                return new LatexToken(LatexTokenKind.Command, _source[_position++].ToString(), start);
            }

            int commandStart = _position;
            while (_position < _source.Length && char.IsLetter(_source[_position]))
            {
                _position++;
            }

            return new LatexToken(LatexTokenKind.Command, _source.Substring(commandStart, _position - commandStart), start);
        }

        private LatexToken ReadText(int start)
        {
            while (_position < _source.Length && "{}^_&\\".IndexOf(_source[_position]) < 0)
            {
                _position++;
            }

            return new LatexToken(LatexTokenKind.Text, _source.Substring(start, _position - start), start);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

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

            return StripInvisible(source).Replace("\r\n", "\n").Replace("\r", "\n").Trim();
        }

        // Word text pasted from the web carries zero-width marks, no-break spaces and broken surrogates that split commands like \hat.
        // Text read back from a Word equation uses Unicode math italic letters (and U+210E for h); map them to plain ASCII.
        private static string StripInvisible(string source)
        {
            StringBuilder builder = new StringBuilder(source.Length);
            for (int i = 0; i < source.Length; i++)
            {
                char current = source[i];
                if (char.IsHighSurrogate(current) && i + 1 < source.Length && char.IsLowSurrogate(source[i + 1]))
                {
                    int codePoint = char.ConvertToUtf32(current, source[i + 1]);
                    if (codePoint >= 0x1D400 && codePoint <= 0x1D6A3)
                    {
                        int offset = (codePoint - 0x1D400) % 52;
                        builder.Append((char)(offset < 26 ? 'A' + offset : 'a' + offset - 26));
                    }
                    else if (codePoint >= 0x1D7CE && codePoint <= 0x1D7FF)
                    {
                        builder.Append((char)('0' + (codePoint - 0x1D7CE) % 10));
                    }
                    else if (CharUnicodeInfo.GetUnicodeCategory(source, i) != UnicodeCategory.Format)
                    {
                        builder.Append(current).Append(source[i + 1]);
                    }
                    i++;
                    continue;
                }
                if (current == '\u210E')
                {
                    builder.Append('h');
                    continue;
                }
                if (char.IsSurrogate(current)) continue;
                UnicodeCategory category = char.GetUnicodeCategory(current);
                if (category == UnicodeCategory.Format) continue;
                builder.Append(category == UnicodeCategory.SpaceSeparator ? ' ' : current);
            }
            return builder.ToString();
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

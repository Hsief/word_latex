using System.Collections.Generic;

namespace AfterMathCore.LatexParser
{
    /// <summary>Node kinds in the compact equation intermediate representation.</summary>
    public enum LatexNodeKind
    {
        Sequence,
        Text,
        Fraction,
        Radical,
        Accent,
        Script,
        Nary,
        Matrix,
        Delimiter,
        Styled,
        Bar,
        LimitUpper,
        LimitLower,
        GroupCharacter,
        BorderBox
    }

    /// <summary>Math style values that map directly to OMML run styles.</summary>
    public enum MathVariant
    {
        Default,
        Plain,
        Bold,
        Italic,
        BoldItalic
    }

    /// <summary>A small, immutable-by-convention intermediate representation for equations.</summary>
    public sealed class LatexNode
    {
        public LatexNode(LatexNodeKind kind)
        {
            Kind = kind;
            Children = new List<LatexNode>();
            Slots = new Dictionary<string, LatexNode>();
            Variant = MathVariant.Default;
        }

        public LatexNodeKind Kind { get; set; }
        public string Value { get; set; }
        public MathVariant Variant { get; set; }
        public IList<LatexNode> Children { get; private set; }
        public IDictionary<string, LatexNode> Slots { get; private set; }

        public static LatexNode Text(string value)
        {
            LatexNode node = new LatexNode(LatexNodeKind.Text);
            node.Value = value;
            return node;
        }

        public static LatexNode Sequence(IList<LatexNode> children)
        {
            List<LatexNode> flattened = new List<LatexNode>();
            foreach (LatexNode child in children)
            {
                if (child.Kind == LatexNodeKind.Sequence)
                {
                    foreach (LatexNode nested in child.Children)
                    {
                        flattened.Add(nested);
                    }
                }
                else if (child.Kind != LatexNodeKind.Text || !string.IsNullOrEmpty(child.Value))
                {
                    flattened.Add(child);
                }
            }

            if (flattened.Count == 1)
            {
                return flattened[0];
            }

            LatexNode result = new LatexNode(LatexNodeKind.Sequence);
            foreach (LatexNode child in flattened)
            {
                result.Children.Add(child);
            }
            return result;
        }
    }
}

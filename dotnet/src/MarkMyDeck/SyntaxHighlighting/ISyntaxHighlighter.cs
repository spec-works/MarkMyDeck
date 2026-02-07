using System.Collections.Generic;

namespace MarkMyDeck.SyntaxHighlighting;

/// <summary>
/// Interface for syntax highlighters that tokenize code into colored segments.
/// </summary>
public interface ISyntaxHighlighter
{
    /// <summary>
    /// Highlights code by breaking it into syntax tokens.
    /// </summary>
    IEnumerable<SyntaxToken> Highlight(string code, string language);

    /// <summary>
    /// Checks if this highlighter supports the specified language.
    /// </summary>
    bool SupportsLanguage(string language);
}

/// <summary>
/// Represents a syntax token with its text content and classification type.
/// </summary>
public class SyntaxToken
{
    public string Text { get; }
    public TokenType Type { get; }

    public SyntaxToken(string text, TokenType type)
    {
        Text = text;
        Type = type;
    }
}

/// <summary>
/// Classification types for syntax tokens.
/// </summary>
public enum TokenType
{
    Keyword,
    String,
    Number,
    Comment,
    Operator,
    Identifier,
    Type,
    Function,
    Property,
    Default
}

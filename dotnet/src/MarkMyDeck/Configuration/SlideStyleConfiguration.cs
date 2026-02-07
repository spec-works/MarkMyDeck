using System;

namespace MarkMyDeck.Configuration;

/// <summary>
/// Configuration for presentation styling.
/// </summary>
public class SlideStyleConfiguration
{
    /// <summary>
    /// Default font name for body text.
    /// </summary>
    public string DefaultFontName { get; set; } = "Calibri";

    /// <summary>
    /// Default font size for body text (in points).
    /// </summary>
    public int DefaultFontSize { get; set; } = 18;

    /// <summary>
    /// Font size for slide titles (in points).
    /// </summary>
    public int TitleFontSize { get; set; } = 36;

    /// <summary>
    /// Font size for slide subtitles / H2 headings (in points).
    /// </summary>
    public int SubtitleFontSize { get; set; } = 28;

    /// <summary>
    /// Font size for H3 headings (in points).
    /// </summary>
    public int Heading3FontSize { get; set; } = 24;

    /// <summary>
    /// Font size for H4 headings (in points).
    /// </summary>
    public int Heading4FontSize { get; set; } = 22;

    /// <summary>
    /// Font size for H5 headings (in points).
    /// </summary>
    public int Heading5FontSize { get; set; } = 20;

    /// <summary>
    /// Font size for H6 headings (in points).
    /// </summary>
    public int Heading6FontSize { get; set; } = 18;

    /// <summary>
    /// Title text color (hex color without #).
    /// </summary>
    public string TitleColor { get; set; } = "2E74B5";

    /// <summary>
    /// Body text color (hex color without #).
    /// </summary>
    public string BodyColor { get; set; } = "333333";

    /// <summary>
    /// Font name for code blocks and inline code.
    /// </summary>
    public string CodeFontName { get; set; } = "Consolas";

    /// <summary>
    /// Font size for code blocks and inline code (in points).
    /// </summary>
    public int CodeFontSize { get; set; } = 14;

    /// <summary>
    /// Background color for code blocks (hex color without #).
    /// </summary>
    public string CodeBackgroundColor { get; set; } = "F5F5F5";

    /// <summary>
    /// Color scheme for syntax highlighting (hex colors without #).
    /// </summary>
    public SyntaxColorScheme? SyntaxColorScheme { get; set; }

    /// <summary>
    /// Gets the font size in points for a heading level.
    /// </summary>
    public int GetHeadingFontSize(int level)
    {
        return level switch
        {
            1 => TitleFontSize,
            2 => SubtitleFontSize,
            3 => Heading3FontSize,
            4 => Heading4FontSize,
            5 => Heading5FontSize,
            6 => Heading6FontSize,
            _ => DefaultFontSize
        };
    }
}

/// <summary>
/// Color scheme for syntax highlighting.
/// All colors are hex format without # prefix (e.g., "569CD6" for blue).
/// </summary>
public class SyntaxColorScheme
{
    public string KeywordColor { get; set; } = "569CD6";
    public string StringColor { get; set; } = "CE9178";
    public string NumberColor { get; set; } = "098658";
    public string CommentColor { get; set; } = "6A9955";
    public string OperatorColor { get; set; } = "4A4A4A";
    public string TypeColor { get; set; } = "4EC9B0";
    public string FunctionColor { get; set; } = "C4A000";
    public string PropertyColor { get; set; } = "4FC1FF";
    public string IdentifierColor { get; set; } = "383838";
    public string DefaultColor { get; set; } = "383838";

    /// <summary>
    /// Gets the color for a specific token type.
    /// </summary>
    public string GetColorForTokenType(MarkMyDeck.SyntaxHighlighting.TokenType type)
    {
        return type switch
        {
            MarkMyDeck.SyntaxHighlighting.TokenType.Keyword => KeywordColor,
            MarkMyDeck.SyntaxHighlighting.TokenType.String => StringColor,
            MarkMyDeck.SyntaxHighlighting.TokenType.Number => NumberColor,
            MarkMyDeck.SyntaxHighlighting.TokenType.Comment => CommentColor,
            MarkMyDeck.SyntaxHighlighting.TokenType.Operator => OperatorColor,
            MarkMyDeck.SyntaxHighlighting.TokenType.Type => TypeColor,
            MarkMyDeck.SyntaxHighlighting.TokenType.Function => FunctionColor,
            MarkMyDeck.SyntaxHighlighting.TokenType.Property => PropertyColor,
            MarkMyDeck.SyntaxHighlighting.TokenType.Identifier => IdentifierColor,
            _ => DefaultColor
        };
    }
}

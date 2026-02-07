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
    public string DefaultFontName { get; set; } = "Segoe UI";

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
    public int SubtitleFontSize { get; set; } = 30;

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
    public string TitleColor { get; set; } = "FFFFFF";

    /// <summary>
    /// Body text color (hex color without #).
    /// </summary>
    public string BodyColor { get; set; } = "2D2D2D";

    /// <summary>
    /// Accent color used for title bar, table headers, links (hex without #).
    /// </summary>
    public string AccentColor { get; set; } = "1B3A5C";

    /// <summary>
    /// Secondary accent color for highlights and emphasis (hex without #).
    /// </summary>
    public string AccentColor2 { get; set; } = "2E86DE";

    /// <summary>
    /// Slide background color (hex without #).
    /// </summary>
    public string SlideBackgroundColor { get; set; } = "FAFBFC";

    /// <summary>
    /// Title bar background color (hex without #).
    /// </summary>
    public string TitleBarColor { get; set; } = "1B3A5C";

    /// <summary>
    /// Subtle border/separator color (hex without #).
    /// </summary>
    public string BorderColor { get; set; } = "DEE2E6";

    /// <summary>
    /// Table header background color (hex without #).
    /// </summary>
    public string TableHeaderColor { get; set; } = "1B3A5C";

    /// <summary>
    /// Table header text color (hex without #).
    /// </summary>
    public string TableHeaderTextColor { get; set; } = "FFFFFF";

    /// <summary>
    /// Table alternate row background (hex without #).
    /// </summary>
    public string TableStripeColor { get; set; } = "F0F4F8";

    /// <summary>
    /// Font name for code blocks and inline code.
    /// </summary>
    public string CodeFontName { get; set; } = "Cascadia Code";

    /// <summary>
    /// Font size for code blocks and inline code (in points).
    /// </summary>
    public int CodeFontSize { get; set; } = 13;

    /// <summary>
    /// Background color for code blocks (hex color without #).
    /// </summary>
    public string CodeBackgroundColor { get; set; } = "1E1E2E";

    /// <summary>
    /// Default text color inside code blocks (hex without #).
    /// </summary>
    public string CodeForegroundColor { get; set; } = "CDD6F4";

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
/// Color scheme for syntax highlighting — Catppuccin Mocha inspired for dark code backgrounds.
/// All colors are hex format without # prefix.
/// </summary>
public class SyntaxColorScheme
{
    public string KeywordColor { get; set; } = "CBA6F7";
    public string StringColor { get; set; } = "A6E3A1";
    public string NumberColor { get; set; } = "FAB387";
    public string CommentColor { get; set; } = "6C7086";
    public string OperatorColor { get; set; } = "89DCEB";
    public string TypeColor { get; set; } = "F9E2AF";
    public string FunctionColor { get; set; } = "89B4FA";
    public string PropertyColor { get; set; } = "74C7EC";
    public string IdentifierColor { get; set; } = "CDD6F4";
    public string DefaultColor { get; set; } = "CDD6F4";

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

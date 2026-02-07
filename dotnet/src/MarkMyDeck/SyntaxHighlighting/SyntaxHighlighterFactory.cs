using System;
using System.Collections.Generic;
using System.Linq;

namespace MarkMyDeck.SyntaxHighlighting;

/// <summary>
/// Factory for creating and selecting appropriate syntax highlighters based on language.
/// </summary>
public class SyntaxHighlighterFactory
{
    private readonly List<ISyntaxHighlighter> _highlighters;

    public SyntaxHighlighterFactory()
    {
        _highlighters = new List<ISyntaxHighlighter>
        {
            new HttpHighlighter(),
            new TypeSpecHighlighter(),
            new BashHighlighter(),
            new ColorCodeHighlighter()
        };
    }

    public IEnumerable<SyntaxToken> Highlight(string code, string language)
    {
        if (string.IsNullOrEmpty(code))
            return Enumerable.Empty<SyntaxToken>();

        if (string.IsNullOrWhiteSpace(language))
            return CreateDefaultTokens(code);

        var highlighter = _highlighters.FirstOrDefault(h => h.SupportsLanguage(language));

        if (highlighter != null)
            return highlighter.Highlight(code, language);

        return CreateDefaultTokens(code);
    }

    public bool IsLanguageSupported(string language)
    {
        if (string.IsNullOrWhiteSpace(language))
            return false;

        return _highlighters.Any(h => h.SupportsLanguage(language));
    }

    private IEnumerable<SyntaxToken> CreateDefaultTokens(string code)
    {
        yield return new SyntaxToken(code, TokenType.Default);
    }
}

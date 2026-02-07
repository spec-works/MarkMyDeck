using System.Linq;
using System.Text;
using Markdig.Syntax;
using MarkMyDeck.SyntaxHighlighting;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace MarkMyDeck.Converters.BlockRenderers;

/// <summary>
/// Renderer for code blocks — creates a standalone shape with background color.
/// </summary>
public class CodeBlockRenderer : OpenXmlObjectRenderer<CodeBlock>
{
    private readonly SyntaxHighlighterFactory _highlighterFactory = new();

    protected override void Write(OpenXmlPresentationRenderer renderer, CodeBlock obj)
    {
        string? language = null;
        if (obj is FencedCodeBlock fencedBlock && !string.IsNullOrEmpty(fencedBlock.Info))
        {
            language = fencedBlock.Info.Trim();
        }

        bool useSyntaxHighlighting = renderer.Options.EnableSyntaxHighlighting &&
                                      !string.IsNullOrWhiteSpace(language) &&
                                      _highlighterFactory.IsLanguageSupported(language!);

        var slide = renderer.CurrentSlide;
        var styles = slide.Styles;

        // Calculate height based on number of lines
        var lines = obj.Lines.Lines;
        int lineCount = 0;
        int lastNonEmptyLine = -1;

        if (lines != null)
        {
            for (int i = 0; i < lines.Length; i++)
            {
                if (!string.IsNullOrWhiteSpace(lines[i].Slice.ToString()))
                    lastNonEmptyLine = i;
            }
            lineCount = lastNonEmptyLine + 1;
        }

        if (lineCount == 0) lineCount = 1;

        // 1 point = 12700 EMU; line height ~ 1.4x font size
        var lineHeightEmu = (long)(styles.CodeFontSize * 12700 * 1.4);
        var totalHeight = lineHeightEmu * lineCount + 182880; // + padding

        // Code blocks are standalone shapes with background
        var shape = slide.AddCodeBlockShape(totalHeight, styles.CodeBackgroundColor);
        renderer.CurrentShape = shape;

        if (lines != null)
        {
            for (int i = 0; i <= lastNonEmptyLine; i++)
            {
                var text = lines[i].Slice.ToString();
                var paragraph = slide.AddParagraphToShape(shape);

                // Set tight line spacing
                var pProps = new D.ParagraphProperties();
                pProps.Append(new D.LineSpacing(new D.SpacingPercent { Val = 100000 }));
                pProps.Append(new D.SpaceBefore(new D.SpacingPoints { Val = 0 }));
                pProps.Append(new D.SpaceAfter(new D.SpacingPoints { Val = 0 }));
                paragraph.Append(pProps);

                if (useSyntaxHighlighting)
                {
                    var tokens = _highlighterFactory.Highlight(text, language!);
                    var colorScheme = styles.SyntaxColorScheme ?? new Configuration.SyntaxColorScheme();

                    foreach (var token in tokens)
                    {
                        var color = colorScheme.GetColorForTokenType(token.Type);
                        var run = slide.CreateRun(token.Text, styles.CodeFontName, styles.CodeFontSize, color);
                        paragraph.Append(run);
                    }
                }
                else
                {
                    var run = slide.CreateRun(text, styles.CodeFontName, styles.CodeFontSize, styles.BodyColor);
                    paragraph.Append(run);
                }
            }
        }

        renderer.CurrentParagraph = null;
    }
}

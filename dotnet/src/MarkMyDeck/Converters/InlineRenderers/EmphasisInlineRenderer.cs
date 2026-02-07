using Markdig.Syntax.Inlines;
using D = DocumentFormat.OpenXml.Drawing;

namespace MarkMyDeck.Converters.InlineRenderers;

/// <summary>
/// Renderer for emphasis (bold/italic) inline elements.
/// </summary>
public class EmphasisInlineRenderer : OpenXmlObjectRenderer<EmphasisInline>
{
    protected override void Write(OpenXmlPresentationRenderer renderer, EmphasisInline obj)
    {
        var paragraph = renderer.CurrentParagraph;
        if (paragraph == null || renderer.CurrentShape == null)
            return;

        var slide = renderer.CurrentSlide;
        var styles = slide.Styles;

        bool bold = obj.DelimiterCount >= 2;
        bool italic = obj.DelimiterCount == 1 || obj.DelimiterCount >= 3;

        // Render children with emphasis applied
        var child = obj.FirstChild;
        while (child != null)
        {
            if (child is LiteralInline literal)
            {
                var run = slide.CreateRun(literal.Content.ToString(), styles.DefaultFontName,
                    styles.DefaultFontSize, styles.BodyColor, bold: bold, italic: italic);
                paragraph.Append(run);
            }
            else
            {
                renderer.Write(child);
            }
            child = child.NextSibling;
        }
    }
}

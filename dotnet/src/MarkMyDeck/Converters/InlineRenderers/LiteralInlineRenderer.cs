using System.Linq;
using Markdig.Syntax.Inlines;
using D = DocumentFormat.OpenXml.Drawing;

namespace MarkMyDeck.Converters.InlineRenderers;

/// <summary>
/// Renderer for literal inline text.
/// </summary>
public class LiteralInlineRenderer : OpenXmlObjectRenderer<LiteralInline>
{
    protected override void Write(OpenXmlPresentationRenderer renderer, LiteralInline obj)
    {
        if (obj.Content.IsEmpty)
            return;

        var text = obj.Content.ToString();
        if (string.IsNullOrEmpty(text))
            return;

        var slide = renderer.CurrentSlide;
        var styles = slide.Styles;

        // Get the current paragraph to add the run to
        var paragraph = renderer.CurrentParagraph;
        if (paragraph == null || renderer.CurrentShape == null)
            return;

        var run = slide.CreateRun(text, styles.DefaultFontName, styles.DefaultFontSize, styles.BodyColor);
        paragraph.Append(run);
    }
}

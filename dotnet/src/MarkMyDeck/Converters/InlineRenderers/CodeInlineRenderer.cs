using Markdig.Syntax.Inlines;
using D = DocumentFormat.OpenXml.Drawing;

namespace MarkMyDeck.Converters.InlineRenderers;

/// <summary>
/// Renderer for inline code elements.
/// </summary>
public class CodeInlineRenderer : OpenXmlObjectRenderer<CodeInline>
{
    protected override void Write(OpenXmlPresentationRenderer renderer, CodeInline obj)
    {
        var paragraph = renderer.CurrentParagraph;
        if (paragraph == null || renderer.CurrentShape == null)
            return;

        var slide = renderer.CurrentSlide;
        var styles = slide.Styles;

        var run = slide.CreateRun(obj.Content, styles.CodeFontName, styles.CodeFontSize, styles.BodyColor);
        paragraph.Append(run);
    }
}

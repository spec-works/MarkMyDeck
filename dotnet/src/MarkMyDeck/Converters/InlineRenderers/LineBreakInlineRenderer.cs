using Markdig.Syntax.Inlines;
using D = DocumentFormat.OpenXml.Drawing;

namespace MarkMyDeck.Converters.InlineRenderers;

/// <summary>
/// Renderer for line break inline elements.
/// </summary>
public class LineBreakInlineRenderer : OpenXmlObjectRenderer<LineBreakInline>
{
    protected override void Write(OpenXmlPresentationRenderer renderer, LineBreakInline obj)
    {
        var paragraph = renderer.CurrentParagraph;
        if (paragraph == null || renderer.CurrentShape == null)
            return;

        if (obj.IsHard)
        {
            var br = new D.Break();
            paragraph.Append(br);
        }
    }
}

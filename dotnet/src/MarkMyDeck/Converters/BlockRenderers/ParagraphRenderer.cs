using Markdig.Syntax;
using D = DocumentFormat.OpenXml.Drawing;

namespace MarkMyDeck.Converters.BlockRenderers;

/// <summary>
/// Renderer for paragraph blocks — adds to the content shape.
/// </summary>
public class ParagraphRenderer : OpenXmlObjectRenderer<ParagraphBlock>
{
    protected override void Write(OpenXmlPresentationRenderer renderer, ParagraphBlock obj)
    {
        var slide = renderer.CurrentSlide;

        var paragraph = slide.AddContentParagraph();
        renderer.CurrentShape = slide.GetOrCreateContentShape();
        renderer.CurrentParagraph = paragraph;

        if (obj.Inline != null)
        {
            renderer.WriteChildren(obj.Inline);
        }
    }
}

using Markdig.Syntax;
using D = DocumentFormat.OpenXml.Drawing;

namespace MarkMyDeck.Converters.BlockRenderers;

/// <summary>
/// Renderer for paragraph blocks.
/// </summary>
public class ParagraphRenderer : OpenXmlObjectRenderer<ParagraphBlock>
{
    protected override void Write(OpenXmlPresentationRenderer renderer, ParagraphBlock obj)
    {
        var slide = renderer.CurrentSlide;
        var styles = slide.Styles;
        var height = (long)(styles.DefaultFontSize * 100 * 1.6);

        var shape = slide.AddTextBox(height);
        renderer.CurrentShape = shape;

        var paragraph = slide.AddParagraphToShape(shape);
        renderer.CurrentParagraph = paragraph;

        if (obj.Inline != null)
        {
            renderer.WriteChildren(obj.Inline);
        }
    }
}

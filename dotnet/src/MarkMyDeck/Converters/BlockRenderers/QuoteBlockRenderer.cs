using Markdig.Syntax;
using D = DocumentFormat.OpenXml.Drawing;

namespace MarkMyDeck.Converters.BlockRenderers;

/// <summary>
/// Renderer for quote blocks.
/// </summary>
public class QuoteBlockRenderer : OpenXmlObjectRenderer<QuoteBlock>
{
    protected override void Write(OpenXmlPresentationRenderer renderer, QuoteBlock obj)
    {
        var slide = renderer.CurrentSlide;
        var styles = slide.Styles;

        foreach (var child in obj)
        {
            if (child is ParagraphBlock paragraphBlock)
            {
                var height = (long)(styles.DefaultFontSize * 100 * 1.6);

                // Indent quote blocks and use italic styling
                var shape = slide.AddTextBox(height, xOffset: 914400, width: slide.ContentWidth - 457200);
                renderer.CurrentShape = shape;

                var paragraph = slide.AddParagraphToShape(shape);
                renderer.CurrentParagraph = paragraph;

                if (paragraphBlock.Inline != null)
                {
                    renderer.WriteChildren(paragraphBlock.Inline);
                }

                // Apply italic to all runs in the paragraph
                foreach (var run in paragraph.Elements<D.Run>())
                {
                    if (run.RunProperties == null)
                        run.RunProperties = new D.RunProperties();
                    run.RunProperties.Italic = true;
                }
            }
            else
            {
                renderer.Write(child);
            }
        }
    }
}

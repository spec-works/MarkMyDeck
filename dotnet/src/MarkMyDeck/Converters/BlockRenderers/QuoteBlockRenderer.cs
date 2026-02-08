using Markdig.Syntax;
using D = DocumentFormat.OpenXml.Drawing;

namespace MarkMyDeck.Converters.BlockRenderers;

/// <summary>
/// Renderer for quote blocks — adds italic paragraphs with indent to the content shape.
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
                if (slide.WouldOverflowWithParagraph)
                {
                    slide = renderer.NewContinuationSlide();
                }

                var paragraph = slide.AddContentParagraph();
                renderer.CurrentShape = slide.GetOrCreateContentShape();
                renderer.CurrentParagraph = paragraph;

                // Indent and style as quote
                var pProps = new D.ParagraphProperties();
                pProps.Append(new D.SpaceBefore(new D.SpacingPoints { Val = 200 }));
                paragraph.Append(pProps);

                // Add indent via a run with spaces (MarginL not available on ParagraphProperties in Drawing)
                var indentRun = slide.CreateRun("    ", styles.DefaultFontName, styles.DefaultFontSize);
                paragraph.Append(indentRun);

                if (paragraphBlock.Inline != null)
                {
                    renderer.WriteChildren(paragraphBlock.Inline);
                }

                // Apply italic to all runs
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

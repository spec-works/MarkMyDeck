using System.Linq;
using Markdig.Syntax;
using D = DocumentFormat.OpenXml.Drawing;

namespace MarkMyDeck.Converters.BlockRenderers;

/// <summary>
/// Renderer for list blocks — adds bullet/number paragraphs to the content shape.
/// </summary>
public class ListRenderer : OpenXmlObjectRenderer<ListBlock>
{
    protected override void Write(OpenXmlPresentationRenderer renderer, ListBlock listBlock)
    {
        RenderList(renderer, listBlock, 0);
    }

    private void RenderList(OpenXmlPresentationRenderer renderer, ListBlock listBlock, int level)
    {
        var slide = renderer.CurrentSlide;
        var styles = slide.Styles;
        int itemIndex = 1;

        foreach (var item in listBlock)
        {
            if (item is ListItemBlock listItem)
            {
                foreach (var block in listItem)
                {
                    if (block is ParagraphBlock paragraphBlock)
                    {
                        // Check overflow before adding list item
                        if (slide.WouldOverflowWithParagraph)
                        {
                            slide = renderer.NewContinuationSlide();
                        }

                        var paragraph = slide.AddContentParagraph();
                        renderer.CurrentShape = slide.GetOrCreateContentShape();
                        renderer.CurrentParagraph = paragraph;

                        // Set indentation
                        var indent = 457200L * (level + 1); // 0.5 inch per level
                        var pProps = new D.ParagraphProperties { Indent = -228600, LeftMargin = (int)indent };
                        pProps.Append(new D.SpaceBefore(new D.SpacingPoints { Val = 100 }));
                        paragraph.Append(pProps);

                        // Add bullet/number prefix
                        string prefix = listBlock.IsOrdered
                            ? $"{itemIndex}. "
                            : GetBulletChar(level) + " ";

                        var prefixRun = slide.CreateRun(prefix, styles.DefaultFontName, styles.DefaultFontSize, styles.BodyColor);
                        paragraph.Append(prefixRun);

                        if (paragraphBlock.Inline != null)
                        {
                            renderer.WriteChildren(paragraphBlock.Inline);
                        }
                    }
                    else if (block is ListBlock nestedList)
                    {
                        RenderList(renderer, nestedList, level + 1);
                    }
                    else
                    {
                        renderer.Write(block);
                    }
                }

                itemIndex++;
            }
        }
    }

    private string GetBulletChar(int level)
    {
        return (level % 3) switch
        {
            0 => "•",
            1 => "○",
            2 => "■",
            _ => "•"
        };
    }
}

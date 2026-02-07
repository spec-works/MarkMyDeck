using System.Linq;
using Markdig.Syntax;
using D = DocumentFormat.OpenXml.Drawing;

namespace MarkMyDeck.Converters.BlockRenderers;

/// <summary>
/// Renderer for list blocks (ordered and unordered).
/// </summary>
public class ListRenderer : OpenXmlObjectRenderer<ListBlock>
{
    protected override void Write(OpenXmlPresentationRenderer renderer, ListBlock listBlock)
    {
        var slide = renderer.CurrentSlide;
        var styles = slide.Styles;

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
                        var height = (long)(styles.DefaultFontSize * 100 * 1.5);
                        var indent = 457200L * (level + 1); // 0.5 inch per level

                        var shape = slide.AddTextBox(height, xOffset: indent, width: slide.ContentWidth - indent + 457200);
                        renderer.CurrentShape = shape;

                        var paragraph = slide.AddParagraphToShape(shape);
                        renderer.CurrentParagraph = paragraph;

                        // Add bullet/number prefix
                        string prefix;
                        if (listBlock.IsOrdered)
                        {
                            prefix = $"{itemIndex}. ";
                        }
                        else
                        {
                            prefix = GetBulletChar(level) + " ";
                        }

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

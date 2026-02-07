using System.Linq;
using Markdig.Syntax;
using D = DocumentFormat.OpenXml.Drawing;

namespace MarkMyDeck.Converters.BlockRenderers;

/// <summary>
/// Renderer for heading blocks. H1 and H2 create new slides.
/// </summary>
public class HeadingRenderer : OpenXmlObjectRenderer<HeadingBlock>
{
    protected override void Write(OpenXmlPresentationRenderer renderer, HeadingBlock obj)
    {
        // H1 and H2 create new slides
        if (obj.Level <= 2)
        {
            // If the current slide already has content, create a new one
            // Always create a new slide for headings to separate content
            var slide = renderer.NewSlide();

            var fontSize = slide.Styles.GetHeadingFontSize(obj.Level);
            var height = (long)(fontSize * 100 * 1.8); // approximate height in EMUs

            var shape = slide.AddTextBox(height);
            renderer.CurrentShape = shape;

            var paragraph = slide.AddParagraphToShape(shape);
            renderer.CurrentParagraph = paragraph;

            // Render inline content with title styling
            if (obj.Inline != null)
            {
                renderer.WriteChildren(obj.Inline);
            }

            // Apply bold and color to all runs in the paragraph
            ApplyHeadingStyle(paragraph, slide.Styles, obj.Level);
        }
        else
        {
            // H3-H6 rendered as styled text within current slide
            var slide = renderer.CurrentSlide;
            var fontSize = slide.Styles.GetHeadingFontSize(obj.Level);
            var height = (long)(fontSize * 100 * 1.6);

            var shape = slide.AddTextBox(height);
            renderer.CurrentShape = shape;

            var paragraph = slide.AddParagraphToShape(shape);
            renderer.CurrentParagraph = paragraph;

            if (obj.Inline != null)
            {
                renderer.WriteChildren(obj.Inline);
            }

            ApplyHeadingStyle(paragraph, slide.Styles, obj.Level);
        }
    }

    private void ApplyHeadingStyle(D.Paragraph paragraph, Configuration.SlideStyleConfiguration styles, int level)
    {
        var fontSize = styles.GetHeadingFontSize(level);

        // Apply default run properties to the paragraph
        var pProps = new D.ParagraphProperties();
        var defRunProps = new D.DefaultRunProperties
        {
            FontSize = fontSize * 100,
            Bold = true,
            Dirty = false
        };
        defRunProps.Append(new D.SolidFill(new D.RgbColorModelHex { Val = styles.TitleColor }));
        defRunProps.Append(new D.LatinFont { Typeface = styles.DefaultFontName });
        pProps.Append(defRunProps);
        paragraph.InsertAt(pProps, 0);

        // Also update any existing runs
        foreach (var run in paragraph.Elements<D.Run>())
        {
            if (run.RunProperties == null)
            {
                run.RunProperties = new D.RunProperties();
            }
            run.RunProperties.FontSize = fontSize * 100;
            run.RunProperties.Bold = true;
            run.RunProperties.Dirty = false;

            // Add color if not already present
            if (run.RunProperties.Elements<D.SolidFill>().FirstOrDefault() == null)
            {
                run.RunProperties.Append(new D.SolidFill(new D.RgbColorModelHex { Val = styles.TitleColor }));
            }

            // Add font if not already present
            if (run.RunProperties.Elements<D.LatinFont>().FirstOrDefault() == null)
            {
                run.RunProperties.Append(new D.LatinFont { Typeface = styles.DefaultFontName });
            }
        }
    }
}

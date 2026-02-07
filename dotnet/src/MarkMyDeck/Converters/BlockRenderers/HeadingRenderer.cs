using System.Linq;
using Markdig.Syntax;
using D = DocumentFormat.OpenXml.Drawing;

namespace MarkMyDeck.Converters.BlockRenderers;

/// <summary>
/// Renderer for heading blocks. H1 and H2 create new slides with a title shape.
/// </summary>
public class HeadingRenderer : OpenXmlObjectRenderer<HeadingBlock>
{
    protected override void Write(OpenXmlPresentationRenderer renderer, HeadingBlock obj)
    {
        if (obj.Level <= 2)
        {
            // H1/H2 create new slides — consume pending break flag
            renderer.PendingSlideBreak = false;
            var slide = renderer.NewSlide();
            var styles = slide.Styles;
            var fontSize = styles.GetHeadingFontSize(obj.Level);

            var paragraph = slide.AddTitleParagraph();
            renderer.CurrentShape = slide.GetOrCreateTitleShape();
            renderer.CurrentParagraph = paragraph;

            if (obj.Inline != null)
            {
                renderer.WriteChildren(obj.Inline);
            }

            // Style the title runs
            foreach (var run in paragraph.Elements<D.Run>())
            {
                if (run.RunProperties == null)
                    run.RunProperties = new D.RunProperties();
                run.RunProperties.FontSize = fontSize * 100;
                run.RunProperties.Bold = true;
                run.RunProperties.Dirty = false;
                if (!run.RunProperties.Elements<D.SolidFill>().Any())
                    run.RunProperties.Append(new D.SolidFill(new D.RgbColorModelHex { Val = styles.TitleColor }));
                if (!run.RunProperties.Elements<D.LatinFont>().Any())
                    run.RunProperties.Append(new D.LatinFont { Typeface = styles.DefaultFontName });
            }
        }
        else
        {
            // H3-H6 rendered as styled text in the content area
            var slide = renderer.CurrentSlide;
            var styles = slide.Styles;
            var fontSize = styles.GetHeadingFontSize(obj.Level);

            var paragraph = slide.AddContentParagraph();
            renderer.CurrentShape = slide.GetOrCreateContentShape();
            renderer.CurrentParagraph = paragraph;

            // Add spacing before sub-heading
            var pProps = new D.ParagraphProperties();
            pProps.Append(new D.SpaceBefore(new D.SpacingPoints { Val = 600 }));
            paragraph.Append(pProps);

            if (obj.Inline != null)
            {
                renderer.WriteChildren(obj.Inline);
            }

            foreach (var run in paragraph.Elements<D.Run>())
            {
                if (run.RunProperties == null)
                    run.RunProperties = new D.RunProperties();
                run.RunProperties.FontSize = fontSize * 100;
                run.RunProperties.Bold = true;
                run.RunProperties.Dirty = false;
                if (!run.RunProperties.Elements<D.SolidFill>().Any())
                    run.RunProperties.Append(new D.SolidFill(new D.RgbColorModelHex { Val = styles.TitleColor }));
                if (!run.RunProperties.Elements<D.LatinFont>().Any())
                    run.RunProperties.Append(new D.LatinFont { Typeface = styles.DefaultFontName });
            }
        }
    }
}

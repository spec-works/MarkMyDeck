using Markdig.Syntax;

namespace MarkMyDeck.Converters.BlockRenderers;

/// <summary>
/// Renderer for thematic breaks (horizontal rules). Creates a new slide.
/// </summary>
public class ThematicBreakRenderer : OpenXmlObjectRenderer<ThematicBreakBlock>
{
    protected override void Write(OpenXmlPresentationRenderer renderer, ThematicBreakBlock obj)
    {
        // Thematic break forces a new slide
        renderer.NewSlide();
    }
}

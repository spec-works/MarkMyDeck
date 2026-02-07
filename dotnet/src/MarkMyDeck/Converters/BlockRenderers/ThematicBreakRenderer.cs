using Markdig.Syntax;

namespace MarkMyDeck.Converters.BlockRenderers;

/// <summary>
/// Renderer for thematic breaks (horizontal rules).
/// Sets a flag so the next heading doesn't create a duplicate slide.
/// If no heading follows, creates a new slide.
/// </summary>
public class ThematicBreakRenderer : OpenXmlObjectRenderer<ThematicBreakBlock>
{
    protected override void Write(OpenXmlPresentationRenderer renderer, ThematicBreakBlock obj)
    {
        // Mark that a slide break is pending.
        // HeadingRenderer will consume this flag instead of creating a double slide.
        renderer.PendingSlideBreak = true;
    }
}

using Markdig.Renderers;
using Markdig.Syntax;

namespace MarkMyDeck.Converters;

/// <summary>
/// Base class for OpenXML Presentation object renderers.
/// </summary>
/// <typeparam name="TObject">The type of markdown object to render.</typeparam>
public abstract class OpenXmlObjectRenderer<TObject> : MarkdownObjectRenderer<OpenXmlPresentationRenderer, TObject>
    where TObject : MarkdownObject
{
}

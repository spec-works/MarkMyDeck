using System;
using System.Collections.Generic;
using System.IO;
using Markdig.Renderers;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using MarkMyDeck.Configuration;
using MarkMyDeck.Converters.BlockRenderers;
using MarkMyDeck.Converters.InlineRenderers;
using MarkMyDeck.OpenXml;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace MarkMyDeck.Converters;

/// <summary>
/// Renders Markdig AST to OpenXML PowerPoint format.
/// </summary>
public class OpenXmlPresentationRenderer : RendererBase, IDisposable
{
    private readonly PresentationBuilder _presentationBuilder;
    private SlideManager? _currentSlide;

    public PresentationBuilder PresentationBuilder => _presentationBuilder;
    public ConversionOptions Options => _presentationBuilder.Options;

    /// <summary>
    /// Gets or creates the current slide. If no slide exists yet, creates the first one.
    /// </summary>
    public SlideManager CurrentSlide
    {
        get
        {
            if (_currentSlide == null)
            {
                _currentSlide = _presentationBuilder.AddSlide();
            }
            return _currentSlide;
        }
    }

    /// <summary>
    /// The current shape being rendered into. Block renderers set this before rendering inline content.
    /// </summary>
    public P.Shape? CurrentShape { get; set; }

    /// <summary>
    /// The current paragraph being rendered into.
    /// </summary>
    public D.Paragraph? CurrentParagraph { get; set; }

    /// <summary>
    /// Flag set by ThematicBreakRenderer. Consumed by HeadingRenderer or next block.
    /// </summary>
    public bool PendingSlideBreak { get; set; }

    public OpenXmlPresentationRenderer(Stream outputStream, ConversionOptions? options = null)
    {
        _presentationBuilder = new PresentationBuilder(outputStream, options, leaveOpen: true);

        if (options?.DocumentTitle != null || options?.Author != null || options?.Subject != null)
        {
            _presentationBuilder.SetDocumentProperties(options?.DocumentTitle, options?.Author, options?.Subject);
        }

        // Register block renderers
        ObjectRenderers.Add(new HeadingRenderer());
        ObjectRenderers.Add(new ParagraphRenderer());
        ObjectRenderers.Add(new CodeBlockRenderer());
        ObjectRenderers.Add(new ThematicBreakRenderer());
        ObjectRenderers.Add(new QuoteBlockRenderer());
        ObjectRenderers.Add(new ListRenderer());
        ObjectRenderers.Add(new TableRenderer());

        // Register inline renderers
        ObjectRenderers.Add(new LiteralInlineRenderer());
        ObjectRenderers.Add(new EmphasisInlineRenderer());
        ObjectRenderers.Add(new CodeInlineRenderer());
        ObjectRenderers.Add(new LineBreakInlineRenderer());
        ObjectRenderers.Add(new LinkInlineRenderer());
    }

    /// <summary>
    /// Creates a new slide and sets it as the current slide.
    /// </summary>
    public SlideManager NewSlide()
    {
        _currentSlide = _presentationBuilder.AddSlide();
        CurrentShape = null;
        CurrentParagraph = null;
        return _currentSlide;
    }

    /// <summary>
    /// Creates a continuation slide with the same title + " (cont.)".
    /// Used when content overflows the current slide.
    /// </summary>
    public SlideManager NewContinuationSlide()
    {
        var previousTitle = _currentSlide?.GetTitleText();
        var slide = NewSlide();
        var styles = slide.Styles;

        if (!string.IsNullOrEmpty(previousTitle))
        {
            var titleText = previousTitle!;
            if (!titleText.EndsWith(" (cont.)"))
                titleText += " (cont.)";

            var paragraph = slide.AddTitleParagraph();
            CurrentShape = slide.GetOrCreateTitleShape();
            CurrentParagraph = paragraph;

            var run = slide.CreateRun(titleText, styles.TitleFontName,
                styles.TitleFontSize, styles.TitleColor);
            paragraph.Append(run);
        }

        return slide;
    }

    public override object Render(MarkdownObject markdownObject)
    {
        Write(markdownObject);
        return null!;
    }

    /// <summary>
    /// Override Write to check for pending slide breaks before non-heading blocks.
    /// </summary>
    public new void Write(MarkdownObject obj)
    {
        if (PendingSlideBreak && obj is Block && !(obj is HeadingBlock) && !(obj is ThematicBreakBlock))
        {
            // Thematic break was not followed by a heading; force a new slide now
            PendingSlideBreak = false;
            NewSlide();
        }
        base.Write(obj);
    }

    /// <summary>
    /// Finalizes the presentation.
    /// </summary>
    public void FinalizeDocument()
    {
        // Ensure at least one slide exists
        if (_presentationBuilder.SlideCount == 0)
        {
            _presentationBuilder.AddSlide();
        }

        _presentationBuilder.Save();
    }

    public void Dispose()
    {
        _presentationBuilder?.Dispose();
    }
}

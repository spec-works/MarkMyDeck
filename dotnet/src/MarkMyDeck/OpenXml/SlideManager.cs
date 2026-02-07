using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Drawing;
using MarkMyDeck.Configuration;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace MarkMyDeck.OpenXml;

/// <summary>
/// Manages the current slide's content using a title shape and a content shape.
/// </summary>
public class SlideManager
{
    private readonly SlidePart _slidePart;
    private readonly PresentationBuilder _builder;
    private int _shapeIdCounter = 2;

    // Layout constants in EMUs (914400 EMUs = 1 inch)
    private const long LeftMargin = 457200;     // 0.5 inch
    private const long TopMargin = 274638;      // ~0.3 inch
    private const long RightMargin = 457200;    // 0.5 inch
    private const long TitleHeight = 990600;    // ~1.08 inch
    private const long TitleContentGap = 182880; // ~0.2 inch

    private long _contentWidth;
    private long _slideWidth;
    private long _slideHeight;
    private long _contentTop;
    private long _contentHeight;
    private long _currentY; // tracks position for standalone shapes (tables, images, code blocks)

    private P.Shape? _titleShape;
    private P.Shape? _contentShape;
    private int _contentParagraphCount;

    public SlidePart SlidePart => _slidePart;
    public SlideStyleConfiguration Styles => _builder.Options.Styles;
    public ConversionOptions Options => _builder.Options;
    public PresentationBuilder Builder => _builder;

    public SlideManager(SlidePart slidePart, PresentationBuilder builder)
    {
        _slidePart = slidePart;
        _builder = builder;
        _slideWidth = (long)(builder.Options.SlideWidthInches * 914400);
        _slideHeight = (long)(builder.Options.SlideHeightInches * 914400);
        _contentWidth = _slideWidth - LeftMargin - RightMargin;
        _contentTop = TopMargin + TitleHeight + TitleContentGap;
        _contentHeight = _slideHeight - _contentTop - 274638; // bottom margin
        _currentY = _contentTop;
    }

    /// <summary>
    /// Gets the shape tree of the current slide.
    /// </summary>
    public ShapeTree GetShapeTree()
    {
        return _slidePart.Slide.CommonSlideData!.ShapeTree!;
    }

    /// <summary>
    /// Gets the available content width in EMUs.
    /// </summary>
    public long ContentWidth => _contentWidth;

    /// <summary>
    /// Gets or creates the title shape (top of slide).
    /// </summary>
    public P.Shape GetOrCreateTitleShape()
    {
        if (_titleShape == null)
        {
            _titleShape = new P.Shape(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties { Id = (uint)_shapeIdCounter++, Name = "Title" },
                    new P.NonVisualShapeDrawingProperties(new D.ShapeLocks { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties()),
                new P.ShapeProperties(
                    new D.Transform2D(
                        new D.Offset { X = LeftMargin, Y = TopMargin },
                        new D.Extents { Cx = _contentWidth, Cy = TitleHeight }),
                    new D.PresetGeometry(new D.AdjustValueList()) { Preset = D.ShapeTypeValues.Rectangle }),
                new P.TextBody(
                    new D.BodyProperties
                    {
                        Wrap = D.TextWrappingValues.Square,
                        RightToLeftColumns = false,
                        Anchor = D.TextAnchoringTypeValues.Bottom
                    },
                    new D.ListStyle())
            );
            GetShapeTree().Append(_titleShape);
        }
        return _titleShape;
    }

    /// <summary>
    /// Gets or creates the main content shape (body of slide, below title).
    /// </summary>
    public P.Shape GetOrCreateContentShape()
    {
        if (_contentShape == null)
        {
            _contentShape = new P.Shape(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties { Id = (uint)_shapeIdCounter++, Name = "Content" },
                    new P.NonVisualShapeDrawingProperties(new D.ShapeLocks { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties()),
                new P.ShapeProperties(
                    new D.Transform2D(
                        new D.Offset { X = LeftMargin, Y = _contentTop },
                        new D.Extents { Cx = _contentWidth, Cy = _contentHeight }),
                    new D.PresetGeometry(new D.AdjustValueList()) { Preset = D.ShapeTypeValues.Rectangle }),
                new P.TextBody(
                    new D.BodyProperties
                    {
                        Wrap = D.TextWrappingValues.Square,
                        RightToLeftColumns = false,
                        Anchor = D.TextAnchoringTypeValues.Top
                    },
                    new D.ListStyle())
            );
            GetShapeTree().Append(_contentShape);
        }
        return _contentShape;
    }

    /// <summary>
    /// Whether the content shape has been created (i.e., there is body content).
    /// </summary>
    public bool HasContentShape => _contentShape != null;

    /// <summary>
    /// Adds a paragraph to the title shape.
    /// </summary>
    public D.Paragraph AddTitleParagraph()
    {
        var shape = GetOrCreateTitleShape();
        var paragraph = new D.Paragraph();
        shape.TextBody!.Append(paragraph);
        return paragraph;
    }

    /// <summary>
    /// Adds a paragraph to the content shape and tracks estimated height.
    /// </summary>
    public D.Paragraph AddContentParagraph()
    {
        var shape = GetOrCreateContentShape();
        var paragraph = new D.Paragraph();
        shape.TextBody!.Append(paragraph);
        _contentParagraphCount++;
        // Update _currentY to account for content paragraphs
        // Estimate ~0.35 inch per paragraph
        _currentY = _contentTop + (long)(_contentParagraphCount * 320040);
        return paragraph;
    }

    /// <summary>
    /// Adds a standalone text box with a solid background fill (for code blocks).
    /// If a content shape exists, shrinks it and positions the code block below.
    /// </summary>
    public P.Shape AddCodeBlockShape(long height, string bgColorHex)
    {
        // If content shape exists, resize it to fit its paragraphs
        if (_contentShape != null)
        {
            var estimatedContentHeight = (long)(_contentParagraphCount * 320040); // ~0.35in per paragraph
            if (estimatedContentHeight < 182880) estimatedContentHeight = 182880; // min 0.2in
            var xfrm = _contentShape.ShapeProperties!.GetFirstChild<D.Transform2D>()!;
            xfrm.Extents!.Cy = estimatedContentHeight;
            _currentY = _contentTop + estimatedContentHeight + 91440; // gap
        }

        var shape = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = (uint)_shapeIdCounter++, Name = $"Code {_shapeIdCounter}" },
                new P.NonVisualShapeDrawingProperties(new D.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties()),
            new P.ShapeProperties(
                new D.Transform2D(
                    new D.Offset { X = LeftMargin, Y = _currentY },
                    new D.Extents { Cx = _contentWidth, Cy = height }),
                new D.PresetGeometry(new D.AdjustValueList()) { Preset = D.ShapeTypeValues.Rectangle },
                new D.SolidFill(new D.RgbColorModelHex { Val = bgColorHex })),
            new P.TextBody(
                new D.BodyProperties
                {
                    Wrap = D.TextWrappingValues.Square,
                    RightToLeftColumns = false,
                    LeftInset = 91440,   // 0.1 inch padding
                    TopInset = 45720,
                    RightInset = 91440,
                    BottomInset = 45720
                },
                new D.ListStyle())
        );

        GetShapeTree().Append(shape);
        _currentY += height;

        return shape;
    }

    /// <summary>
    /// Advances the Y cursor past the content shape so standalone elements appear below it.
    /// Call this before adding code blocks, tables, or images if content shape exists.
    /// </summary>
    public void SyncCursorAfterContent()
    {
        // Place standalone elements after the content area
        // We estimate based on content shape bottom
        if (_contentShape != null)
        {
            _currentY = _contentTop + _contentHeight + 91440; // small gap
        }
    }

    /// <summary>
    /// Adds a paragraph to an existing shape's TextBody.
    /// </summary>
    public D.Paragraph AddParagraphToShape(P.Shape shape)
    {
        var textBody = shape.TextBody!;
        var paragraph = new D.Paragraph();
        textBody.Append(paragraph);
        return paragraph;
    }

    /// <summary>
    /// Creates a run with specified text and formatting.
    /// </summary>
    public D.Run CreateRun(string text, string fontName, int fontSizePt, string? colorHex = null,
        bool bold = false, bool italic = false, bool underline = false)
    {
        var runProps = new D.RunProperties { Language = "en-US", FontSize = fontSizePt * 100, Dirty = false };

        if (bold) runProps.Bold = true;
        if (italic) runProps.Italic = true;
        if (underline) runProps.Underline = D.TextUnderlineValues.Single;

        if (!string.IsNullOrEmpty(colorHex))
        {
            runProps.Append(new D.SolidFill(new D.RgbColorModelHex { Val = colorHex }));
        }

        runProps.Append(new D.LatinFont { Typeface = fontName });

        var run = new D.Run(runProps, new D.Text(text));
        return run;
    }

    /// <summary>
    /// Creates a run with hyperlink.
    /// </summary>
    public D.Run CreateHyperlinkRun(string text, string url, string fontName, int fontSizePt)
    {
        var relationshipId = AddHyperlinkRelationship(url);

        var runProps = new D.RunProperties { Language = "en-US", FontSize = fontSizePt * 100, Dirty = false };
        runProps.Append(new D.SolidFill(new D.RgbColorModelHex { Val = "0563C1" }));
        runProps.Append(new D.LatinFont { Typeface = fontName });
        runProps.Underline = D.TextUnderlineValues.Single;
        runProps.Append(new D.HyperlinkOnClick { Id = relationshipId });

        return new D.Run(runProps, new D.Text(text));
    }

    /// <summary>
    /// Adds a hyperlink relationship to the slide part.
    /// </summary>
    public string AddHyperlinkRelationship(string url)
    {
        var rel = _slidePart.AddHyperlinkRelationship(new Uri(url, UriKind.RelativeOrAbsolute), true);
        return rel.Id;
    }

    /// <summary>
    /// Adds an image to the slide.
    /// </summary>
    public void AddImage(byte[] imageData, string contentType, long widthEmu, long heightEmu)
    {
        var partType = contentType.ToLowerInvariant() switch
        {
            "image/png" => ImagePartType.Png,
            "image/jpeg" or "image/jpg" => ImagePartType.Jpeg,
            "image/gif" => ImagePartType.Gif,
            "image/bmp" => ImagePartType.Bmp,
            "image/tiff" => ImagePartType.Tiff,
            _ => ImagePartType.Png
        };

        var imagePart = _slidePart.AddImagePart(partType);
        using (var stream = new MemoryStream(imageData))
        {
            imagePart.FeedData(stream);
        }

        var relationshipId = _slidePart.GetIdOfPart(imagePart);

        var picture = new P.Picture(
            new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties { Id = (uint)_shapeIdCounter++, Name = $"Image {_shapeIdCounter}" },
                new P.NonVisualPictureDrawingProperties(new D.PictureLocks { NoChangeAspect = true }),
                new ApplicationNonVisualDrawingProperties()),
            new P.BlipFill(
                new D.Blip { Embed = relationshipId },
                new D.Stretch(new D.FillRectangle())),
            new P.ShapeProperties(
                new D.Transform2D(
                    new D.Offset { X = LeftMargin, Y = _currentY },
                    new D.Extents { Cx = widthEmu, Cy = heightEmu }),
                new D.PresetGeometry(new D.AdjustValueList()) { Preset = D.ShapeTypeValues.Rectangle })
        );

        GetShapeTree().Append(picture);
        _currentY += heightEmu;
    }

    /// <summary>
    /// Adds a table to the slide at the current position.
    /// </summary>
    public D.Table AddTable(int rows, int cols, long height)
    {
        var colWidth = _contentWidth / cols;

        var tableGrid = new D.TableGrid();
        for (int c = 0; c < cols; c++)
        {
            tableGrid.Append(new D.GridColumn { Width = colWidth });
        }

        var table = new D.Table(
            new D.TableProperties { FirstRow = true, BandRow = true },
            tableGrid
        );

        var graphicFrame = new P.GraphicFrame(
            new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = (uint)_shapeIdCounter++, Name = $"Table {_shapeIdCounter}" },
                new P.NonVisualGraphicFrameDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new P.Transform(
                new D.Offset { X = LeftMargin, Y = _currentY },
                new D.Extents { Cx = _contentWidth, Cy = height }),
            new D.Graphic(new D.GraphicData(table)
            {
                Uri = "http://schemas.openxmlformats.org/drawingml/2006/table"
            })
        );

        GetShapeTree().Append(graphicFrame);
        _currentY += height;

        return table;
    }
}

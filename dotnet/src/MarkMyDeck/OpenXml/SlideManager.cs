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
/// Manages the current slide's content, positioning elements vertically.
/// </summary>
public class SlideManager
{
    private readonly SlidePart _slidePart;
    private readonly PresentationBuilder _builder;
    private int _shapeIdCounter = 2;

    // Content area margins in EMUs (914400 EMUs = 1 inch)
    private const long LeftMargin = 457200;    // 0.5 inch
    private const long TopMargin = 457200;     // 0.5 inch
    private const long RightMargin = 457200;   // 0.5 inch
    private const long BottomMargin = 457200;  // 0.5 inch

    // Current Y position for content placement
    private long _currentY;
    private long _contentWidth;
    private long _slideWidth;
    private long _slideHeight;

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
        _currentY = TopMargin;
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
    /// Adds a text box shape to the slide at the current Y position.
    /// Returns the shape so callers can add paragraphs.
    /// </summary>
    public P.Shape AddTextBox(long height, long? xOffset = null, long? width = null)
    {
        var x = xOffset ?? LeftMargin;
        var w = width ?? _contentWidth;

        var shape = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = (uint)_shapeIdCounter++, Name = $"TextBox {_shapeIdCounter}" },
                new P.NonVisualShapeDrawingProperties(new D.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties()),
            new P.ShapeProperties(
                new D.Transform2D(
                    new D.Offset { X = x, Y = _currentY },
                    new D.Extents { Cx = w, Cy = height }),
                new D.PresetGeometry(new D.AdjustValueList()) { Preset = D.ShapeTypeValues.Rectangle }),
            new P.TextBody(
                new D.BodyProperties { Wrap = D.TextWrappingValues.Square, RightToLeftColumns = false },
                new D.ListStyle())
        );

        GetShapeTree().Append(shape);
        _currentY += height;

        return shape;
    }

    /// <summary>
    /// Adds a text box with a solid background fill.
    /// </summary>
    public P.Shape AddTextBoxWithBackground(long height, string bgColorHex, long? xOffset = null, long? width = null)
    {
        var x = xOffset ?? LeftMargin;
        var w = width ?? _contentWidth;

        var shape = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = (uint)_shapeIdCounter++, Name = $"TextBox {_shapeIdCounter}" },
                new P.NonVisualShapeDrawingProperties(new D.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties()),
            new P.ShapeProperties(
                new D.Transform2D(
                    new D.Offset { X = x, Y = _currentY },
                    new D.Extents { Cx = w, Cy = height }),
                new D.PresetGeometry(new D.AdjustValueList()) { Preset = D.ShapeTypeValues.Rectangle },
                new D.SolidFill(new D.RgbColorModelHex { Val = bgColorHex })),
            new P.TextBody(
                new D.BodyProperties { Wrap = D.TextWrappingValues.Square, RightToLeftColumns = false },
                new D.ListStyle())
        );

        GetShapeTree().Append(shape);
        _currentY += height;

        return shape;
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
    /// Adds vertical spacing.
    /// </summary>
    public void AddSpacing(long emus)
    {
        _currentY += emus;
    }

    /// <summary>
    /// Adds a horizontal line (thematic break) to the slide.
    /// </summary>
    public void AddHorizontalLine()
    {
        var shape = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = (uint)_shapeIdCounter++, Name = $"Line {_shapeIdCounter}" },
                new P.NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new P.ShapeProperties(
                new D.Transform2D(
                    new D.Offset { X = LeftMargin, Y = _currentY + 45000 },
                    new D.Extents { Cx = _contentWidth, Cy = 0 }),
                new D.PresetGeometry(new D.AdjustValueList()) { Preset = D.ShapeTypeValues.Line },
                new D.Outline(
                    new D.SolidFill(new D.RgbColorModelHex { Val = "AAAAAA" })
                ) { Width = 12700 })
        );

        GetShapeTree().Append(shape);
        _currentY += 90000;
    }

    /// <summary>
    /// Adds a table to the slide.
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

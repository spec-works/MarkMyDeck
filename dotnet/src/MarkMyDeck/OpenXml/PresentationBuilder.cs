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
/// Manages the creation and manipulation of OpenXML PowerPoint presentations.
/// </summary>
public class PresentationBuilder : IDisposable
{
    private readonly Stream _outputStream;
    private readonly bool _leaveOpen;
    private bool _disposed;
    private int _slideIdCounter = 256;
    private int _relationshipIdCounter = 1;

    public PresentationDocument PresentationDocument { get; private set; }
    public PresentationPart PresentationPart { get; private set; }
    public ConversionOptions Options { get; }

    private readonly List<SlidePart> _slideParts = new();

    public PresentationBuilder(Stream outputStream, ConversionOptions? options = null, bool leaveOpen = false)
    {
        _outputStream = outputStream ?? throw new ArgumentNullException(nameof(outputStream));
        _leaveOpen = leaveOpen;
        Options = options ?? new ConversionOptions();

        PresentationDocument = PresentationDocument.Create(_outputStream, PresentationDocumentType.Presentation, autoSave: false);
        PresentationPart = PresentationDocument.AddPresentationPart();
        PresentationPart.Presentation = new P.Presentation();

        // Create slide master and layout
        CreateSlideMaster();

        // Initialize slide ID list and slide size
        PresentationPart.Presentation.SlideIdList = new SlideIdList();
        PresentationPart.Presentation.SlideSize = new SlideSize
        {
            Cx = (int)(Options.SlideWidthInches * 914400),
            Cy = (int)(Options.SlideHeightInches * 914400),
            Type = SlideSizeValues.Custom
        };
        PresentationPart.Presentation.NotesSize = new NotesSize
        {
            Cx = (int)(Options.SlideHeightInches * 914400),
            Cy = (int)(Options.SlideWidthInches * 914400)
        };
    }

    /// <summary>
    /// Creates a new slide and returns the SlideManager for it.
    /// </summary>
    public SlideManager AddSlide()
    {
        var slidePart = PresentationPart.AddNewPart<SlidePart>(GetNextRelationshipId());
        var slide = new Slide(new CommonSlideData(new ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                new P.NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties(new D.TransformGroup(
                new D.Offset { X = 0, Y = 0 },
                new D.Extents { Cx = 0, Cy = 0 },
                new D.ChildOffset { X = 0, Y = 0 },
                new D.ChildExtents { Cx = 0, Cy = 0 }))
        )), new ColorMapOverride(new D.MasterColorMapping()));

        slidePart.Slide = slide;

        // Add to slide ID list
        var slideId = new SlideId
        {
            Id = (uint)_slideIdCounter++,
            RelationshipId = PresentationPart.GetIdOfPart(slidePart)
        };
        PresentationPart.Presentation.SlideIdList!.Append(slideId);

        _slideParts.Add(slidePart);

        // Add slide layout relationship
        var layoutPart = PresentationPart.SlideMasterParts.First().SlideLayoutParts.First();
        slidePart.AddPart(layoutPart, "rId1");

        return new SlideManager(slidePart, this);
    }

    /// <summary>
    /// Gets the number of slides in the presentation.
    /// </summary>
    public int SlideCount => _slideParts.Count;

    /// <summary>
    /// Sets presentation metadata properties.
    /// </summary>
    public void SetDocumentProperties(string? title = null, string? author = null, string? subject = null)
    {
        var props = PresentationDocument.PackageProperties;
        if (!string.IsNullOrEmpty(title)) props.Title = title;
        if (!string.IsNullOrEmpty(author)) props.Creator = author;
        if (!string.IsNullOrEmpty(subject)) props.Subject = subject;
    }

    /// <summary>
    /// Saves the presentation.
    /// </summary>
    public void Save()
    {
        PresentationDocument.Save();
    }

    private string GetNextRelationshipId()
    {
        return $"rId{_relationshipIdCounter++}";
    }

    private void CreateSlideMaster()
    {
        var slideMasterPart = PresentationPart.AddNewPart<SlideMasterPart>(GetNextRelationshipId());
        var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>("rId1");

        // Create minimal slide layout
        slideLayoutPart.SlideLayout = new SlideLayout(
            new CommonSlideData(new ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(new D.TransformGroup(
                    new D.Offset { X = 0, Y = 0 },
                    new D.Extents { Cx = 0, Cy = 0 },
                    new D.ChildOffset { X = 0, Y = 0 },
                    new D.ChildExtents { Cx = 0, Cy = 0 }))
            )),
            new ColorMapOverride(new D.MasterColorMapping()))
        {
            Type = SlideLayoutValues.Blank
        };

        // Slide layout must reference back to its slide master
        slideLayoutPart.AddPart(slideMasterPart, "rId1");

        // Create slide master
        slideMasterPart.SlideMaster = new SlideMaster(
            new CommonSlideData(new ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(new D.TransformGroup(
                    new D.Offset { X = 0, Y = 0 },
                    new D.Extents { Cx = 0, Cy = 0 },
                    new D.ChildOffset { X = 0, Y = 0 },
                    new D.ChildExtents { Cx = 0, Cy = 0 }))
            )),
            new P.ColorMap
            {
                Background1 = D.ColorSchemeIndexValues.Light1,
                Text1 = D.ColorSchemeIndexValues.Dark1,
                Background2 = D.ColorSchemeIndexValues.Light2,
                Text2 = D.ColorSchemeIndexValues.Dark2,
                Accent1 = D.ColorSchemeIndexValues.Accent1,
                Accent2 = D.ColorSchemeIndexValues.Accent2,
                Accent3 = D.ColorSchemeIndexValues.Accent3,
                Accent4 = D.ColorSchemeIndexValues.Accent4,
                Accent5 = D.ColorSchemeIndexValues.Accent5,
                Accent6 = D.ColorSchemeIndexValues.Accent6,
                Hyperlink = D.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink
            },
            new SlideLayoutIdList(
                new SlideLayoutId { Id = 2147483649, RelationshipId = "rId1" }
            ));

        // Create theme part
        var themePart = slideMasterPart.AddNewPart<ThemePart>("rId2");
        themePart.Theme = CreateDefaultTheme();

        // Add slide master ID list to presentation
        PresentationPart.Presentation.SlideMasterIdList = new SlideMasterIdList(
            new SlideMasterId
            {
                Id = 2147483648,
                RelationshipId = PresentationPart.GetIdOfPart(slideMasterPart)
            }
        );
    }

    private D.Theme CreateDefaultTheme()
    {
        var s = Options.Styles;
        return new D.Theme(
            new D.ThemeElements(
                new D.ColorScheme(
                    new D.Dark1Color(new D.RgbColorModelHex { Val = s.AccentColor }),
                    new D.Light1Color(new D.RgbColorModelHex { Val = s.SlideBackgroundColor }),
                    new D.Dark2Color(new D.RgbColorModelHex { Val = s.BodyColor }),
                    new D.Light2Color(new D.RgbColorModelHex { Val = s.TableStripeColor }),
                    new D.Accent1Color(new D.RgbColorModelHex { Val = s.AccentColor2 }),
                    new D.Accent2Color(new D.RgbColorModelHex { Val = "10AC84" }),
                    new D.Accent3Color(new D.RgbColorModelHex { Val = "EE5A24" }),
                    new D.Accent4Color(new D.RgbColorModelHex { Val = "6C5CE7" }),
                    new D.Accent5Color(new D.RgbColorModelHex { Val = "FDA7DF" }),
                    new D.Accent6Color(new D.RgbColorModelHex { Val = "F9CA24" }),
                    new D.Hyperlink(new D.RgbColorModelHex { Val = s.AccentColor2 }),
                    new D.FollowedHyperlinkColor(new D.RgbColorModelHex { Val = "6C5CE7" })
                )
                { Name = "MarkMyDeck" },
                new D.FontScheme(
                    new D.MajorFont(
                        new D.LatinFont { Typeface = s.TitleFontName },
                        new D.EastAsianFont { Typeface = "" },
                        new D.ComplexScriptFont { Typeface = "" }),
                    new D.MinorFont(
                        new D.LatinFont { Typeface = s.DefaultFontName },
                        new D.EastAsianFont { Typeface = "" },
                        new D.ComplexScriptFont { Typeface = "" })
                )
                { Name = "MarkMyDeck" },
                new D.FormatScheme(
                    new D.FillStyleList(
                        new D.SolidFill(new D.SchemeColor { Val = D.SchemeColorValues.PhColor }),
                        new D.SolidFill(new D.SchemeColor { Val = D.SchemeColorValues.PhColor }),
                        new D.SolidFill(new D.SchemeColor { Val = D.SchemeColorValues.PhColor })),
                    new D.LineStyleList(
                        new D.Outline(new D.SolidFill(new D.SchemeColor { Val = D.SchemeColorValues.PhColor })) { Width = 6350, CapType = D.LineCapValues.Flat, CompoundLineType = D.CompoundLineValues.Single, Alignment = D.PenAlignmentValues.Center },
                        new D.Outline(new D.SolidFill(new D.SchemeColor { Val = D.SchemeColorValues.PhColor })) { Width = 12700, CapType = D.LineCapValues.Flat, CompoundLineType = D.CompoundLineValues.Single, Alignment = D.PenAlignmentValues.Center },
                        new D.Outline(new D.SolidFill(new D.SchemeColor { Val = D.SchemeColorValues.PhColor })) { Width = 19050, CapType = D.LineCapValues.Flat, CompoundLineType = D.CompoundLineValues.Single, Alignment = D.PenAlignmentValues.Center }),
                    new D.EffectStyleList(
                        new D.EffectStyle(new D.EffectList()),
                        new D.EffectStyle(new D.EffectList()),
                        new D.EffectStyle(new D.EffectList())),
                    new D.BackgroundFillStyleList(
                        new D.SolidFill(new D.SchemeColor { Val = D.SchemeColorValues.PhColor }),
                        new D.SolidFill(new D.SchemeColor { Val = D.SchemeColorValues.PhColor }),
                        new D.SolidFill(new D.SchemeColor { Val = D.SchemeColorValues.PhColor }))
                )
                { Name = "MarkMyDeck" }
            )
        )
        { Name = "MarkMyDeck Theme" };
    }

    public void Dispose()
    {
        if (_disposed) return;
        PresentationDocument?.Dispose();
        if (!_leaveOpen) _outputStream?.Dispose();
        _disposed = true;
    }
}

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using FluentAssertions;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace MarkMyDeck.Tests;

public class BasicConversionTests
{
    [Fact]
    public void HelloWorld_ShouldConvertSuccessfully()
    {
        // Arrange
        var markdown = "Hello World";

        // Act
        var pptxBytes = MarkdownConverter.ConvertToPptxBytes(markdown);

        // Assert
        pptxBytes.Should().NotBeNull();
        pptxBytes.Length.Should().BeGreaterThan(0);

        using var ms = new MemoryStream(pptxBytes);
        using var doc = PresentationDocument.Open(ms, false);
        doc.Should().NotBeNull();
        doc.PresentationPart.Should().NotBeNull();
    }

    [Fact]
    public void Heading1_ShouldCreateNewSlide()
    {
        // Arrange
        var markdown = "# My Title";

        // Act
        var pptxBytes = MarkdownConverter.ConvertToPptxBytes(markdown);

        // Assert
        using var ms = new MemoryStream(pptxBytes);
        using var doc = PresentationDocument.Open(ms, false);

        var slideIds = doc.PresentationPart!.Presentation.SlideIdList!.Elements<SlideId>().ToList();
        slideIds.Should().HaveCountGreaterThanOrEqualTo(1);
    }

    [Fact]
    public void MultipleH1_ShouldCreateMultipleSlides()
    {
        // Arrange
        var markdown = "# Slide 1\n\nContent for slide 1\n\n# Slide 2\n\nContent for slide 2";

        // Act
        var pptxBytes = MarkdownConverter.ConvertToPptxBytes(markdown);

        // Assert
        using var ms = new MemoryStream(pptxBytes);
        using var doc = PresentationDocument.Open(ms, false);

        var slideIds = doc.PresentationPart!.Presentation.SlideIdList!.Elements<SlideId>().ToList();
        slideIds.Should().HaveCountGreaterThanOrEqualTo(2);
    }

    [Fact]
    public void ThematicBreak_ShouldCreateNewSlide()
    {
        // Arrange
        var markdown = "# First Slide\n\nSome content\n\n---\n\n# Second Slide";

        // Act
        var pptxBytes = MarkdownConverter.ConvertToPptxBytes(markdown);

        // Assert
        using var ms = new MemoryStream(pptxBytes);
        using var doc = PresentationDocument.Open(ms, false);

        var slideIds = doc.PresentationPart!.Presentation.SlideIdList!.Elements<SlideId>().ToList();
        // With the --- followed by # heading, the thematic break is consumed by the heading
        // so we get exactly 2 slides: "First Slide" and "Second Slide"
        slideIds.Should().HaveCountGreaterThanOrEqualTo(2);
    }

    [Fact]
    public void BoldText_ShouldCreateBoldRun()
    {
        // Arrange
        var markdown = "# Title\n\nThis is **bold** text";

        // Act
        var pptxBytes = MarkdownConverter.ConvertToPptxBytes(markdown);

        // Assert
        using var ms = new MemoryStream(pptxBytes);
        using var doc = PresentationDocument.Open(ms, false);

        // Find runs with bold
        var slideParts = doc.PresentationPart!.SlideParts.ToList();
        slideParts.Should().HaveCountGreaterThan(0);

        var boldRuns = slideParts
            .SelectMany(sp => sp.Slide.Descendants<D.Run>())
            .Where(r => r.RunProperties?.Bold != null && r.RunProperties.Bold.HasValue && r.RunProperties.Bold.Value == true)
            .ToList();
        boldRuns.Should().HaveCountGreaterThan(0);
    }

    [Fact]
    public void InlineCode_ShouldCreateCodeRun()
    {
        // Arrange
        var markdown = "# Title\n\nThis is `inline code` text";

        // Act
        var pptxBytes = MarkdownConverter.ConvertToPptxBytes(markdown);

        // Assert
        using var ms = new MemoryStream(pptxBytes);
        using var doc = PresentationDocument.Open(ms, false);

        var slideParts = doc.PresentationPart!.SlideParts.ToList();
        slideParts.Should().HaveCountGreaterThan(0);

        // Find runs with Consolas font (code font)
        var codeRuns = slideParts
            .SelectMany(sp => sp.Slide.Descendants<D.Run>())
            .Where(r => r.RunProperties?.Elements<D.LatinFont>().Any(f => f.Typeface == "Consolas") == true)
            .ToList();
        codeRuns.Should().HaveCountGreaterThan(0);
    }

    [Fact]
    public void CodeBlock_ShouldCreateShapeWithBackground()
    {
        // Arrange
        var markdown = "# Title\n\n```json\n{\"key\": \"value\"}\n```";

        // Act
        var pptxBytes = MarkdownConverter.ConvertToPptxBytes(markdown);

        // Assert
        using var ms = new MemoryStream(pptxBytes);
        using var doc = PresentationDocument.Open(ms, false);

        var slideParts = doc.PresentationPart!.SlideParts.ToList();
        slideParts.Should().HaveCountGreaterThan(0);

        // Find shapes with solid fill (code block background)
        var shapesWithBg = slideParts
            .SelectMany(sp => sp.Slide.Descendants<P.Shape>())
            .Where(s => s.ShapeProperties?.Elements<D.SolidFill>().Any() == true)
            .ToList();
        shapesWithBg.Should().HaveCountGreaterThan(0);
    }

    [Fact]
    public void Hyperlink_ShouldCreateHyperlinkRun()
    {
        // Arrange
        var markdown = "# Title\n\n[Google](https://www.google.com)";

        // Act
        var pptxBytes = MarkdownConverter.ConvertToPptxBytes(markdown);

        // Assert
        using var ms = new MemoryStream(pptxBytes);
        using var doc = PresentationDocument.Open(ms, false);

        var slideParts = doc.PresentationPart!.SlideParts.ToList();
        slideParts.Should().HaveCountGreaterThan(0);

        // Find runs with hyperlink
        var hyperlinkRuns = slideParts
            .SelectMany(sp => sp.Slide.Descendants<D.Run>())
            .Where(r => r.RunProperties?.Elements<D.HyperlinkOnClick>().Any() == true)
            .ToList();
        hyperlinkRuns.Should().HaveCountGreaterThan(0);
    }

    [Fact]
    public void Table_ShouldCreateTable()
    {
        // Arrange
        var markdown = "# Title\n\n| Name | Value |\n|------|-------|\n| A | 1 |\n| B | 2 |";

        // Act
        var pptxBytes = MarkdownConverter.ConvertToPptxBytes(markdown);

        // Assert
        using var ms = new MemoryStream(pptxBytes);
        using var doc = PresentationDocument.Open(ms, false);

        var slideParts = doc.PresentationPart!.SlideParts.ToList();
        slideParts.Should().HaveCountGreaterThan(0);

        var tables = slideParts
            .SelectMany(sp => sp.Slide.Descendants<D.Table>())
            .ToList();
        tables.Should().HaveCount(1);
    }

    [Fact]
    public void ConvertToDocxBytes_EmptyMarkdown_ShouldThrow()
    {
        // Arrange & Act
        var act = () => MarkdownConverter.ConvertToPptxBytes("");

        // Assert
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void ConvertToPptx_NullStream_ShouldThrow()
    {
        // Arrange & Act
        var act = () => MarkdownConverter.ConvertToPptx("# Test", (Stream)null!);

        // Assert
        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void List_ShouldRenderBulletItems()
    {
        // Arrange
        var markdown = "# Title\n\n- Item 1\n- Item 2\n- Item 3";

        // Act
        var pptxBytes = MarkdownConverter.ConvertToPptxBytes(markdown);

        // Assert
        using var ms = new MemoryStream(pptxBytes);
        using var doc = PresentationDocument.Open(ms, false);

        var slideParts = doc.PresentationPart!.SlideParts.ToList();
        slideParts.Should().HaveCountGreaterThan(0);

        // Find text runs containing bullet character
        var bulletRuns = slideParts
            .SelectMany(sp => sp.Slide.Descendants<D.Run>())
            .Where(r => r.Text?.Text?.Contains("•") == true)
            .ToList();
        bulletRuns.Should().HaveCountGreaterThanOrEqualTo(3);
    }
}

using System.Text.Json;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using FluentAssertions;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace MarkMyDeck.Tests;

/// <summary>
/// Data-driven integration tests that validate conversion for all testcases/ files.
/// </summary>
public class TestCaseIntegrationTests
{
    private static readonly string TestCasesRoot = FindTestCasesRoot();

    private static string FindTestCasesRoot()
    {
        // Walk up from the test assembly output directory to find the testcases folder
        var dir = AppContext.BaseDirectory;
        while (dir != null)
        {
            var candidate = Path.Combine(dir, "testcases");
            if (Directory.Exists(candidate))
                return candidate;
            dir = Directory.GetParent(dir)?.FullName;
        }
        throw new DirectoryNotFoundException("Could not find testcases directory");
    }

    public static IEnumerable<object[]> PositiveTestCases()
    {
        var jsonFiles = Directory.GetFiles(TestCasesRoot, "*.json", SearchOption.TopDirectoryOnly);
        foreach (var jsonFile in jsonFiles)
        {
            var name = Path.GetFileNameWithoutExtension(jsonFile);
            yield return new object[] { name };
        }
    }

    public static IEnumerable<object[]> NegativeTestCases()
    {
        var negativeDir = Path.Combine(TestCasesRoot, "negative");
        if (!Directory.Exists(negativeDir))
            yield break;

        var jsonFiles = Directory.GetFiles(negativeDir, "*.json");
        foreach (var jsonFile in jsonFiles)
        {
            var name = Path.GetFileNameWithoutExtension(jsonFile);
            yield return new object[] { name };
        }
    }

    [Theory]
    [MemberData(nameof(PositiveTestCases))]
    public void PositiveTestCase_ShouldProduceValidPresentation(string testCaseName)
    {
        // Arrange
        var jsonPath = Path.Combine(TestCasesRoot, $"{testCaseName}.json");
        var spec = JsonSerializer.Deserialize<TestCaseSpec>(File.ReadAllText(jsonPath))!;

        var mdPath = Path.Combine(TestCasesRoot, spec.InputFile ?? $"{testCaseName}.md");
        var markdown = File.ReadAllText(mdPath);

        // Act
        var pptxBytes = MarkdownConverter.ConvertToPptxBytes(markdown);

        // Assert - basic validity
        pptxBytes.Should().NotBeNull();
        pptxBytes.Length.Should().BeGreaterThan(0);

        using var ms = new MemoryStream(pptxBytes);
        using var doc = PresentationDocument.Open(ms, false);
        doc.Should().NotBeNull();
        doc.PresentationPart.Should().NotBeNull();

        var slideIds = doc.PresentationPart!.Presentation.SlideIdList!
            .Elements<SlideId>().ToList();
        var slideParts = doc.PresentationPart.SlideParts.ToList();

        // Assert - expectations from the JSON spec
        foreach (var expectation in spec.Expectations ?? [])
        {
            AssertExpectation(expectation, slideIds, slideParts, testCaseName);
        }
    }

    [Theory]
    [MemberData(nameof(NegativeTestCases))]
    public void NegativeTestCase_ShouldThrowExpectedException(string testCaseName)
    {
        // Arrange
        var jsonPath = Path.Combine(TestCasesRoot, "negative", $"{testCaseName}.json");
        var spec = JsonSerializer.Deserialize<NegativeTestCaseSpec>(File.ReadAllText(jsonPath))!;

        // Act & Assert
        switch (spec.ExpectedException)
        {
            case "ArgumentException":
                var actArg = () => MarkdownConverter.ConvertToPptxBytes(spec.Input ?? "");
                actArg.Should().Throw<ArgumentException>(
                    $"Test case '{testCaseName}' should throw ArgumentException");
                break;

            case "ArgumentNullException":
                var actNull = () => MarkdownConverter.ConvertToPptx(
                    spec.Input ?? "# Test", (Stream)null!);
                actNull.Should().Throw<ArgumentNullException>(
                    $"Test case '{testCaseName}' should throw ArgumentNullException");
                break;

            default:
                throw new InvalidOperationException(
                    $"Unknown expected exception type: {spec.ExpectedException}");
        }
    }

    private static void AssertExpectation(
        Expectation expectation,
        List<SlideId> slideIds,
        List<SlidePart> slideParts,
        string testCaseName)
    {
        switch (expectation.Type)
        {
            case "slide_count":
                slideIds.Should().HaveCountGreaterThanOrEqualTo(
                    expectation.MinimumValue ?? 1,
                    $"[{testCaseName}] expected at least {expectation.MinimumValue} slides");
                break;

            case "has_bold_runs":
                GetRunsForSlide(slideParts, expectation.SlideIndex ?? 0)
                    .Where(r => r.RunProperties?.Bold?.Value == true)
                    .Should().NotBeEmpty($"[{testCaseName}] expected bold runs on slide {expectation.SlideIndex}");
                break;

            case "has_italic_runs":
                GetRunsForSlide(slideParts, expectation.SlideIndex ?? 0)
                    .Where(r => r.RunProperties?.Italic?.Value == true)
                    .Should().NotBeEmpty($"[{testCaseName}] expected italic runs on slide {expectation.SlideIndex}");
                break;

            case "has_code_font_runs":
                var fontName = expectation.FontName ?? "Cascadia Code";
                GetRunsForSlide(slideParts, expectation.SlideIndex ?? 0)
                    .Where(r => r.RunProperties?.Elements<D.LatinFont>()
                        .Any(f => f.Typeface == fontName) == true)
                    .Should().NotBeEmpty($"[{testCaseName}] expected code font runs on slide {expectation.SlideIndex}");
                break;

            case "has_hyperlink_runs":
                GetRunsForSlide(slideParts, expectation.SlideIndex ?? 0)
                    .Where(r => r.RunProperties?.Elements<D.HyperlinkOnClick>().Any() == true)
                    .Should().NotBeEmpty($"[{testCaseName}] expected hyperlink runs on slide {expectation.SlideIndex}");
                break;

            case "has_shape_with_solid_fill":
                var shapes = slideParts.ElementAt(expectation.SlideIndex ?? 0)
                    .Slide.Descendants<P.Shape>()
                    .Where(s => s.ShapeProperties?.Elements<D.SolidFill>().Any() == true);
                shapes.Should().NotBeEmpty(
                    $"[{testCaseName}] expected shapes with solid fill on slide {expectation.SlideIndex}");
                break;

            case "has_table":
                var tables = slideParts.ElementAt(expectation.SlideIndex ?? 0)
                    .Slide.Descendants<D.Table>().ToList();
                tables.Should().HaveCount(expectation.TableCount ?? 1,
                    $"[{testCaseName}] expected {expectation.TableCount ?? 1} table(s) on slide {expectation.SlideIndex}");
                break;

            case "has_bullet_runs":
                var bulletRuns = GetRunsForSlide(slideParts, expectation.SlideIndex ?? 0)
                    .Where(r => r.Text?.Text?.Contains("•") == true).ToList();
                bulletRuns.Should().HaveCountGreaterThanOrEqualTo(
                    expectation.MinimumCount ?? 1,
                    $"[{testCaseName}] expected at least {expectation.MinimumCount} bullet runs");
                break;

            case "has_numbered_items":
                // Numbered lists produce text runs with the list item content
                var numberedRuns = GetRunsForSlide(slideParts, expectation.SlideIndex ?? 0).ToList();
                numberedRuns.Should().HaveCountGreaterThanOrEqualTo(
                    expectation.MinimumCount ?? 1,
                    $"[{testCaseName}] expected at least {expectation.MinimumCount} numbered item runs");
                break;

            case "has_text_content":
                var allText = string.Join(" ",
                    GetRunsForSlide(slideParts, expectation.SlideIndex ?? 0)
                        .Select(r => r.Text?.Text ?? ""));
                allText.Should().Contain(expectation.ContainsText!,
                    $"[{testCaseName}] expected text containing '{expectation.ContainsText}' on slide {expectation.SlideIndex}");
                break;

            default:
                throw new InvalidOperationException(
                    $"Unknown expectation type: {expectation.Type}");
        }
    }

    private static IEnumerable<D.Run> GetRunsForSlide(List<SlidePart> slideParts, int slideIndex)
    {
        return slideParts.ElementAt(slideIndex).Slide.Descendants<D.Run>();
    }

    private class TestCaseSpec
    {
        [JsonPropertyName("description")]
        public string? Description { get; set; }

        [JsonPropertyName("inputFile")]
        public string? InputFile { get; set; }

        [JsonPropertyName("expectations")]
        public List<Expectation>? Expectations { get; set; }
    }

    private class Expectation
    {
        [JsonPropertyName("type")]
        public string Type { get; set; } = "";

        [JsonPropertyName("minimumValue")]
        public int? MinimumValue { get; set; }

        [JsonPropertyName("slideIndex")]
        public int? SlideIndex { get; set; }

        [JsonPropertyName("tableCount")]
        public int? TableCount { get; set; }

        [JsonPropertyName("minimumCount")]
        public int? MinimumCount { get; set; }

        [JsonPropertyName("fontName")]
        public string? FontName { get; set; }

        [JsonPropertyName("containsText")]
        public string? ContainsText { get; set; }
    }

    private class NegativeTestCaseSpec
    {
        [JsonPropertyName("description")]
        public string? Description { get; set; }

        [JsonPropertyName("input")]
        public string? Input { get; set; }

        [JsonPropertyName("expectedException")]
        public string? ExpectedException { get; set; }
    }
}

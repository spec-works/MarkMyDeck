using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Markdig;
using MarkMyDeck.Configuration;
using MarkMyDeck.Converters;

namespace MarkMyDeck;

/// <summary>
/// Provides methods to convert Markdown to PowerPoint presentations.
/// </summary>
public static class MarkdownConverter
{
    /// <summary>
    /// Converts markdown text to a PowerPoint presentation and saves it to a file.
    /// </summary>
    public static void ConvertToPptx(string markdown, string outputPath, ConversionOptions? options = null)
    {
        if (string.IsNullOrEmpty(markdown))
            throw new ArgumentException("Markdown content cannot be null or empty.", nameof(markdown));

        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentException("Output path cannot be null or empty.", nameof(outputPath));

        using var fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.ReadWrite);
        ConvertToPptx(markdown, fileStream, options);
    }

    /// <summary>
    /// Converts markdown text to a PowerPoint presentation and writes it to a stream.
    /// </summary>
    public static void ConvertToPptx(string markdown, Stream outputStream, ConversionOptions? options = null)
    {
        if (string.IsNullOrEmpty(markdown))
            throw new ArgumentException("Markdown content cannot be null or empty.", nameof(markdown));

        if (outputStream == null)
            throw new ArgumentNullException(nameof(outputStream));

        var pipelineBuilder = new MarkdownPipelineBuilder();

        if (options?.EnableTables ?? true)
        {
            pipelineBuilder = pipelineBuilder.UseAdvancedExtensions();
        }

        var pipeline = pipelineBuilder.Build();
        var document = Markdown.Parse(markdown, pipeline);

        using var renderer = new OpenXmlPresentationRenderer(outputStream, options);
        renderer.Render(document);
        renderer.FinalizeDocument();
    }

    /// <summary>
    /// Converts markdown from a stream to a PowerPoint presentation.
    /// </summary>
    public static void ConvertToPptx(Stream markdownStream, Stream outputStream, ConversionOptions? options = null)
    {
        if (markdownStream == null)
            throw new ArgumentNullException(nameof(markdownStream));

        if (outputStream == null)
            throw new ArgumentNullException(nameof(outputStream));

        using var reader = new StreamReader(markdownStream);
        var markdown = reader.ReadToEnd();
        ConvertToPptx(markdown, outputStream, options);
    }

    /// <summary>
    /// Converts markdown text to a PowerPoint presentation and returns it as a byte array.
    /// </summary>
    public static byte[] ConvertToPptxBytes(string markdown, ConversionOptions? options = null)
    {
        using var ms = new MemoryStream();
        ConvertToPptx(markdown, ms, options);
        return ms.ToArray();
    }

    /// <summary>
    /// Asynchronously converts markdown text to a PowerPoint presentation and saves it to a file.
    /// </summary>
    public static async Task ConvertToPptxAsync(string markdown, string outputPath, ConversionOptions? options = null, CancellationToken cancellationToken = default)
    {
        await Task.Run(() => ConvertToPptx(markdown, outputPath, options), cancellationToken);
    }

    /// <summary>
    /// Asynchronously converts markdown text to a PowerPoint presentation and writes it to a stream.
    /// </summary>
    public static async Task ConvertToPptxAsync(string markdown, Stream outputStream, ConversionOptions? options = null, CancellationToken cancellationToken = default)
    {
        await Task.Run(() => ConvertToPptx(markdown, outputStream, options), cancellationToken);
    }

    /// <summary>
    /// Asynchronously converts markdown from a stream to a PowerPoint presentation.
    /// </summary>
    public static async Task ConvertToPptxAsync(Stream markdownStream, Stream outputStream, ConversionOptions? options = null, CancellationToken cancellationToken = default)
    {
        await Task.Run(() => ConvertToPptx(markdownStream, outputStream, options), cancellationToken);
    }
}

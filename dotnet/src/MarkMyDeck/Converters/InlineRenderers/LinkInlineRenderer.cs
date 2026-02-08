using System;
using System.IO;
using System.Net.Http;
using Markdig.Syntax.Inlines;
using MarkMyDeck.Configuration;
using D = DocumentFormat.OpenXml.Drawing;

namespace MarkMyDeck.Converters.InlineRenderers;

/// <summary>
/// Renderer for link inline elements (hyperlinks and images).
/// </summary>
public class LinkInlineRenderer : OpenXmlObjectRenderer<LinkInline>
{
    private static readonly HttpClient _httpClient = new();

    protected override void Write(OpenXmlPresentationRenderer renderer, LinkInline obj)
    {
        if (obj.IsImage)
        {
            TryInsertImage(renderer, obj);
            return;
        }

        var paragraph = renderer.CurrentParagraph;
        if (paragraph == null || renderer.CurrentShape == null)
            return;

        if (string.IsNullOrEmpty(obj.Url))
        {
            if (obj.FirstChild != null)
                renderer.WriteChildren(obj);
            return;
        }

        // Create hyperlink run(s)
        var slide = renderer.CurrentSlide;
        var child = obj.FirstChild;
        if (child != null)
        {
            while (child != null)
            {
                if (child is LiteralInline literal)
                {
                    var run = slide.CreateHyperlinkRun(literal.Content.ToString(), obj.Url,
                        slide.Styles.DefaultFontName, slide.Styles.DefaultFontSize);
                    paragraph.Append(run);
                }
                child = child.NextSibling;
            }
        }
        else
        {
            var run = slide.CreateHyperlinkRun(obj.Url, obj.Url,
                slide.Styles.DefaultFontName, slide.Styles.DefaultFontSize);
            paragraph.Append(run);
        }
    }

    private void TryInsertImage(OpenXmlPresentationRenderer renderer, LinkInline link)
    {
        if (renderer.Options.ImageStrategy == ImageHandlingStrategy.Skip)
            return;

        try
        {
            if (string.IsNullOrEmpty(link.Url))
                return;

            byte[]? imageData = null;
            string? contentType = null;

            imageData = LoadImageData(link.Url, renderer.Options.BasePath, out contentType);

            if (imageData == null || imageData.Length == 0)
                return;

            // Read actual dimensions from image bytes
            GetImageDimensions(imageData, out int pixelWidth, out int pixelHeight);

            var slide = renderer.CurrentSlide;

            // Convert pixels to EMUs (96 dpi: 1 pixel = 9525 EMUs)
            long imageWidthEmu = (long)pixelWidth * 9525;
            long imageHeightEmu = (long)pixelHeight * 9525;

            // Minimum size
            if (imageWidthEmu < 914400) imageWidthEmu = 914400;   // 1 inch min
            if (imageHeightEmu < 457200) imageHeightEmu = 457200; // 0.5 inch min

            slide.AddImage(imageData, contentType ?? "image/png", imageWidthEmu, imageHeightEmu);
        }
        catch
        {
            // Image load failed — insert alt text fallback
            var paragraph = renderer.CurrentParagraph;
            if (paragraph != null)
            {
                var slide = renderer.CurrentSlide;
                var altText = link.Title ?? link.Url ?? "image";
                var run = slide.CreateRun($"[Image: {altText}]", slide.Styles.DefaultFontName,
                    slide.Styles.DefaultFontSize, slide.Styles.BodyColor, italic: true);
                paragraph.Append(run);
            }
        }
    }

    private byte[]? LoadImageData(string url, string? basePath, out string? contentType)
    {
        contentType = null;

        // Try as absolute URI first (http/https)
        if (Uri.TryCreate(url, UriKind.Absolute, out var uri))
        {
            if (uri.Scheme == "http" || uri.Scheme == "https")
            {
                var response = _httpClient.GetAsync(uri).GetAwaiter().GetResult();
                if (response.IsSuccessStatusCode)
                {
                    contentType = response.Content.Headers.ContentType?.MediaType ?? "image/png";
                    return response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult();
                }
                return null;
            }

            if (uri.Scheme == "file")
            {
                return LoadLocalFile(uri.LocalPath, out contentType);
            }
        }

        // Try as relative or absolute local path
        string filePath = url;

        if (!Path.IsPathRooted(filePath) && !string.IsNullOrEmpty(basePath))
        {
            filePath = Path.Combine(basePath, filePath);
        }

        return LoadLocalFile(filePath, out contentType);
    }

    private byte[]? LoadLocalFile(string path, out string? contentType)
    {
        contentType = null;
        if (!File.Exists(path))
            return null;

        contentType = GetContentTypeFromExtension(Path.GetExtension(path));
        return File.ReadAllBytes(path);
    }

    /// <summary>
    /// Reads image dimensions from the binary header without loading the full image.
    /// Supports PNG, JPEG, GIF, BMP.
    /// </summary>
    private void GetImageDimensions(byte[] data, out int width, out int height)
    {
        width = 800;  // fallback
        height = 600;

        if (data.Length < 24)
            return;

        // PNG: bytes 16-23 contain width (4 bytes big-endian) and height (4 bytes big-endian)
        if (data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47)
        {
            width = (data[16] << 24) | (data[17] << 16) | (data[18] << 8) | data[19];
            height = (data[20] << 24) | (data[21] << 16) | (data[22] << 8) | data[23];
            return;
        }

        // JPEG: scan for SOF0 marker (0xFF 0xC0)
        if (data[0] == 0xFF && data[1] == 0xD8)
        {
            int i = 2;
            while (i < data.Length - 9)
            {
                if (data[i] == 0xFF)
                {
                    byte marker = data[i + 1];
                    // SOF markers: C0-C3, C5-C7, C9-CB, CD-CF
                    if (marker >= 0xC0 && marker <= 0xCF && marker != 0xC4 && marker != 0xC8 && marker != 0xCC)
                    {
                        height = (data[i + 5] << 8) | data[i + 6];
                        width = (data[i + 7] << 8) | data[i + 8];
                        return;
                    }
                    // Skip this segment
                    int segLen = (data[i + 2] << 8) | data[i + 3];
                    i += 2 + segLen;
                }
                else
                {
                    i++;
                }
            }
            return;
        }

        // GIF: bytes 6-9 are width and height (little-endian 16-bit)
        if (data[0] == 0x47 && data[1] == 0x49 && data[2] == 0x46)
        {
            width = data[6] | (data[7] << 8);
            height = data[8] | (data[9] << 8);
            return;
        }

        // BMP: bytes 18-25 are width and height (little-endian 32-bit)
        if (data[0] == 0x42 && data[1] == 0x4D && data.Length >= 26)
        {
            width = data[18] | (data[19] << 8) | (data[20] << 16) | (data[21] << 24);
            height = Math.Abs(data[22] | (data[23] << 8) | (data[24] << 16) | (data[25] << 24));
            return;
        }
    }

    private string GetContentTypeFromExtension(string extension)
    {
        return extension.ToLowerInvariant() switch
        {
            ".png" => "image/png",
            ".jpg" or ".jpeg" => "image/jpeg",
            ".gif" => "image/gif",
            ".bmp" => "image/bmp",
            _ => "image/png"
        };
    }
}

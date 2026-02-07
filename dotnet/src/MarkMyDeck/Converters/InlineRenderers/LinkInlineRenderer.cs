using System;
using System.IO;
using System.Net.Http;
using Markdig.Syntax.Inlines;
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
        var paragraph = renderer.CurrentParagraph;
        if (paragraph == null || renderer.CurrentShape == null)
        {
            // For images, we might need to create a new element outside a text shape
            if (obj.IsImage)
            {
                TryInsertImage(renderer, obj);
                return;
            }
            return;
        }

        if (obj.IsImage)
        {
            // Insert image as alt text fallback since we're inside a paragraph
            var altText = obj.Title ?? obj.Url ?? "image";
            var slide = renderer.CurrentSlide;
            var run = slide.CreateRun($"[Image: {altText}]", slide.Styles.DefaultFontName,
                slide.Styles.DefaultFontSize, slide.Styles.BodyColor, italic: true);
            paragraph.Append(run);
            return;
        }

        if (string.IsNullOrEmpty(obj.Url))
        {
            if (obj.FirstChild != null)
                renderer.WriteChildren(obj);
            return;
        }

        // Create hyperlink run
        var slide2 = renderer.CurrentSlide;
        var child = obj.FirstChild;
        if (child != null)
        {
            while (child != null)
            {
                if (child is LiteralInline literal)
                {
                    var run = slide2.CreateHyperlinkRun(literal.Content.ToString(), obj.Url,
                        slide2.Styles.DefaultFontName, slide2.Styles.DefaultFontSize);
                    paragraph.Append(run);
                }
                child = child.NextSibling;
            }
        }
        else
        {
            var run = slide2.CreateHyperlinkRun(obj.Url, obj.Url,
                slide2.Styles.DefaultFontName, slide2.Styles.DefaultFontSize);
            paragraph.Append(run);
        }
    }

    private void TryInsertImage(OpenXmlPresentationRenderer renderer, LinkInline link)
    {
        try
        {
            if (string.IsNullOrEmpty(link.Url))
                return;

            byte[]? imageData = null;
            string? contentType = null;

            if (Uri.TryCreate(link.Url, UriKind.Absolute, out var uri))
            {
                if (uri.Scheme == "http" || uri.Scheme == "https")
                {
                    var response = _httpClient.GetAsync(uri).GetAwaiter().GetResult();
                    if (response.IsSuccessStatusCode)
                    {
                        imageData = response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult();
                        contentType = response.Content.Headers.ContentType?.MediaType ?? "image/png";
                    }
                }
                else if (uri.Scheme == "file" || !uri.IsAbsoluteUri)
                {
                    var filePath = uri.IsAbsoluteUri ? uri.LocalPath : link.Url;
                    if (File.Exists(filePath))
                    {
                        imageData = File.ReadAllBytes(filePath);
                        contentType = GetContentTypeFromExtension(Path.GetExtension(filePath));
                    }
                }
            }
            else if (File.Exists(link.Url))
            {
                imageData = File.ReadAllBytes(link.Url);
                contentType = GetContentTypeFromExtension(Path.GetExtension(link.Url));
            }

            if (imageData == null || imageData.Length == 0)
                return;

            // Default image dimensions
            long widthEmus = 6 * 914400;
            long heightEmus = 4 * 914400;

            renderer.CurrentSlide.AddImage(imageData, contentType ?? "image/png", widthEmus, heightEmus);
        }
        catch
        {
            // Silently fail for images
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

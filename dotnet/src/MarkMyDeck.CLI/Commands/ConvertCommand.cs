using MarkMyDeck;
using MarkMyDeck.Configuration;

namespace MarkMyDeck.CLI.Commands;

/// <summary>
/// Handles the convert command logic.
/// </summary>
public static class ConvertCommand
{
    public static async Task<int> ExecuteAsync(
        FileInfo input,
        FileInfo? output,
        bool verbose,
        string themeName,
        string? font,
        int? fontSize,
        bool force,
        string? title)
    {
        try
        {
            if (!input.Exists)
            {
                Console.Error.WriteLine($"Error: Input file not found: {input.FullName}");
                return 1;
            }

            // Determine output path
            string outputPath;
            if (output != null)
            {
                outputPath = output.FullName;
            }
            else
            {
                outputPath = Path.ChangeExtension(input.FullName, ".pptx");
            }

            // Check if output file exists
            if (File.Exists(outputPath) && !force)
            {
                Console.Error.WriteLine($"Error: Output file already exists: {outputPath}");
                Console.Error.WriteLine("Use --force to overwrite.");
                return 1;
            }

            if (verbose)
            {
                Console.WriteLine("MarkMyDeck - Markdown to PowerPoint Converter");
                Console.WriteLine($"Input:  {input.FullName}");
                Console.WriteLine($"Output: {outputPath}");
                Console.WriteLine();
            }

            // Resolve theme
            if (!SlideThemePresets.TryParse(themeName, out var theme))
            {
                Console.Error.WriteLine($"Error: Unknown theme '{themeName}'.");
                Console.Error.WriteLine($"Available themes: {string.Join(", ", SlideThemePresets.AvailableThemes)}");
                return 1;
            }

            var styles = SlideThemePresets.Create(theme);
            if (verbose) Console.WriteLine($"Using theme: {theme}");

            // Apply overrides
            if (font != null)
            {
                styles.DefaultFontName = font;
                if (verbose) Console.WriteLine($"  Font override: {font}");
            }

            if (fontSize.HasValue)
            {
                styles.DefaultFontSize = fontSize.Value;
                if (verbose) Console.WriteLine($"  Font size override: {fontSize.Value}pt");
            }

            // Create conversion options
            var options = new ConversionOptions
            {
                Styles = styles,
                DocumentTitle = title,
                BasePath = input.DirectoryName
            };

            // Read markdown file
            if (verbose) Console.WriteLine();
            if (verbose) Console.WriteLine("Reading markdown file...");
            var markdown = await File.ReadAllTextAsync(input.FullName);

            if (string.IsNullOrWhiteSpace(markdown))
            {
                Console.Error.WriteLine("Warning: Input file is empty.");
            }

            // Convert
            if (verbose)
            {
                Console.WriteLine("Converting to PowerPoint presentation...");
                var startTime = DateTime.Now;
                await MarkdownConverter.ConvertToPptxAsync(markdown, outputPath, options);
                var elapsed = DateTime.Now - startTime;
                Console.WriteLine($"Conversion completed in {elapsed.TotalMilliseconds:F0}ms");
            }
            else
            {
                await MarkdownConverter.ConvertToPptxAsync(markdown, outputPath, options);
            }

            Console.WriteLine($"✓ Created: {outputPath}");

            if (verbose)
            {
                var fileInfo = new FileInfo(outputPath);
                Console.WriteLine($"File size: {FormatFileSize(fileInfo.Length)}");
            }

            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            if (verbose)
            {
                Console.Error.WriteLine();
                Console.Error.WriteLine("Stack trace:");
                Console.Error.WriteLine(ex.StackTrace);
            }
            return 1;
        }
    }

    private static string FormatFileSize(long bytes)
    {
        string[] sizes = { "B", "KB", "MB", "GB" };
        double len = bytes;
        int order = 0;
        while (len >= 1024 && order < sizes.Length - 1)
        {
            order++;
            len = len / 1024;
        }
        return $"{len:0.##} {sizes[order]}";
    }
}

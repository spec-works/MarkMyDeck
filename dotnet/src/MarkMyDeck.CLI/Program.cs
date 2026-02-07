using System.CommandLine;
using MarkMyDeck;
using MarkMyDeck.CLI.Commands;

var rootCommand = new RootCommand("MarkMyDeck - Convert Markdown to PowerPoint presentations")
{
    Name = "markmydeck"
};

// Convert command
var convertCommand = new Command("convert", "Convert a Markdown file to a PowerPoint presentation");

var inputOption = new Option<FileInfo>(
    aliases: new[] { "--input", "-i" },
    description: "Input markdown file path (.md)")
{
    IsRequired = true
};
inputOption.AddValidator(result =>
{
    var fileInfo = result.GetValueForOption(inputOption);
    if (fileInfo != null && !fileInfo.Exists)
    {
        result.ErrorMessage = $"Input file not found: {fileInfo.FullName}";
    }
});

var outputOption = new Option<FileInfo?>(
    aliases: new[] { "--output", "-o" },
    description: "Output file path (default: same name with .pptx extension)");

var verboseOption = new Option<bool>(
    aliases: new[] { "--verbose", "-v" },
    description: "Enable verbose output",
    getDefaultValue: () => false);

var fontOption = new Option<string?>(
    aliases: new[] { "--font", "-f" },
    description: "Default font name (e.g., 'Calibri', 'Arial')");

var fontSizeOption = new Option<int?>(
    aliases: new[] { "--font-size", "-s" },
    description: "Default font size in points (e.g., 18, 24)");
fontSizeOption.AddValidator(result =>
{
    var value = result.GetValueForOption(fontSizeOption);
    if (value.HasValue && (value.Value < 6 || value.Value > 72))
    {
        result.ErrorMessage = "Font size must be between 6 and 72 points";
    }
});

var forceOption = new Option<bool>(
    aliases: new[] { "--force" },
    description: "Overwrite output file if it exists",
    getDefaultValue: () => false);

var titleOption = new Option<string?>(
    aliases: new[] { "--title", "-t" },
    description: "Presentation title metadata");

convertCommand.AddOption(inputOption);
convertCommand.AddOption(outputOption);
convertCommand.AddOption(verboseOption);
convertCommand.AddOption(fontOption);
convertCommand.AddOption(fontSizeOption);
convertCommand.AddOption(forceOption);
convertCommand.AddOption(titleOption);

convertCommand.SetHandler(async (context) =>
{
    var input = context.ParseResult.GetValueForOption(inputOption)!;
    var output = context.ParseResult.GetValueForOption(outputOption);
    var verbose = context.ParseResult.GetValueForOption(verboseOption);
    var font = context.ParseResult.GetValueForOption(fontOption);
    var fontSize = context.ParseResult.GetValueForOption(fontSizeOption);
    var force = context.ParseResult.GetValueForOption(forceOption);
    var title = context.ParseResult.GetValueForOption(titleOption);

    var exitCode = await ConvertCommand.ExecuteAsync(
        input, output, verbose, font, fontSize, force, title);
    Environment.Exit(exitCode);
});

rootCommand.AddCommand(convertCommand);

// Version command
var versionCommand = new Command("version", "Display version information");
versionCommand.SetHandler(() =>
{
    var version = typeof(MarkdownConverter).Assembly.GetName().Version;
    Console.WriteLine($"MarkMyDeck v{version?.ToString(3) ?? "0.1.0"}");
    Console.WriteLine(".NET Markdown to PowerPoint Converter");
    Console.WriteLine();
    Console.WriteLine("Built with:");
    Console.WriteLine("  - Markdig (CommonMark parser)");
    Console.WriteLine("  - DocumentFormat.OpenXml");
    Console.WriteLine();
    Console.WriteLine("https://github.com/spec-works/MarkMyDeck");
});

rootCommand.AddCommand(versionCommand);

return await rootCommand.InvokeAsync(args);

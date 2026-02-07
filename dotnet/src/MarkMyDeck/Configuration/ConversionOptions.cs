using System;
using System.Collections.Generic;

namespace MarkMyDeck.Configuration;

/// <summary>
/// Options for converting markdown to PowerPoint presentations.
/// </summary>
public class ConversionOptions
{
    /// <summary>
    /// Style configuration for the presentation.
    /// </summary>
    public SlideStyleConfiguration Styles { get; set; } = new();

    /// <summary>
    /// Enable Markdig advanced extensions (tables, task lists, etc.).
    /// </summary>
    public bool EnableAdvancedExtensions { get; set; } = false;

    /// <summary>
    /// Enable table support.
    /// </summary>
    public bool EnableTables { get; set; } = true;

    /// <summary>
    /// Enable task list support.
    /// </summary>
    public bool EnableTaskLists { get; set; } = true;

    /// <summary>
    /// Enable syntax highlighting for code blocks.
    /// </summary>
    public bool EnableSyntaxHighlighting { get; set; } = true;

    /// <summary>
    /// Presentation title metadata.
    /// </summary>
    public string? DocumentTitle { get; set; }

    /// <summary>
    /// Presentation author metadata.
    /// </summary>
    public string? Author { get; set; }

    /// <summary>
    /// Presentation subject metadata.
    /// </summary>
    public string? Subject { get; set; }

    /// <summary>
    /// Base directory for resolving relative paths (e.g., images).
    /// Set automatically from the input file path when using the CLI.
    /// </summary>
    public string? BasePath { get; set; }

    /// <summary>
    /// Strategy for handling images in the presentation.
    /// </summary>
    public ImageHandlingStrategy ImageStrategy { get; set; } = ImageHandlingStrategy.Embed;

    /// <summary>
    /// Maximum image width in inches.
    /// </summary>
    public int MaxImageWidthInches { get; set; } = 8;

    /// <summary>
    /// Slide width in inches (default: 10 for widescreen 16:9).
    /// </summary>
    public double SlideWidthInches { get; set; } = 10.0;

    /// <summary>
    /// Slide height in inches (default: 7.5 for widescreen 16:9).
    /// </summary>
    public double SlideHeightInches { get; set; } = 7.5;
}

/// <summary>
/// Strategy for handling images in markdown.
/// </summary>
public enum ImageHandlingStrategy
{
    /// <summary>
    /// Embed images in the presentation.
    /// </summary>
    Embed,

    /// <summary>
    /// Keep images as hyperlinks.
    /// </summary>
    Link,

    /// <summary>
    /// Skip images entirely.
    /// </summary>
    Skip
}

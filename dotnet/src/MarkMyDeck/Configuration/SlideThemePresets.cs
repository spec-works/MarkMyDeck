using System;
using System.Collections.Generic;
using System.Linq;

namespace MarkMyDeck.Configuration;

/// <summary>
/// Available built-in presentation themes.
/// </summary>
public enum SlideTheme
{
    /// <summary>Dark navy title bar, light body, Segoe UI. Professional default.</summary>
    Default,
    /// <summary>Clean white/gray, minimal accents. Good for print.</summary>
    Light,
    /// <summary>Dark backgrounds throughout. Easy on the eyes.</summary>
    Dark,
    /// <summary>Conservative blue/gray palette. Boardroom-ready.</summary>
    Corporate,
    /// <summary>High-contrast colors. Bold and energetic.</summary>
    Vibrant
}

/// <summary>
/// Factory that creates <see cref="SlideStyleConfiguration"/> instances for each built-in theme.
/// </summary>
public static class SlideThemePresets
{
    /// <summary>
    /// Returns the list of available theme names.
    /// </summary>
    public static IReadOnlyList<string> AvailableThemes =>
        Enum.GetNames(typeof(SlideTheme)).ToList();

    /// <summary>
    /// Tries to parse a theme name (case-insensitive).
    /// </summary>
    public static bool TryParse(string name, out SlideTheme theme)
    {
#if NETSTANDARD2_1
        foreach (SlideTheme t in Enum.GetValues(typeof(SlideTheme)))
        {
            if (string.Equals(t.ToString(), name, StringComparison.OrdinalIgnoreCase))
            {
                theme = t;
                return true;
            }
        }
        theme = SlideTheme.Default;
        return false;
#else
        return Enum.TryParse(name, ignoreCase: true, out theme);
#endif
    }

    /// <summary>
    /// Creates a <see cref="SlideStyleConfiguration"/> for the given theme.
    /// </summary>
    public static SlideStyleConfiguration Create(SlideTheme theme)
    {
        return theme switch
        {
            SlideTheme.Default => CreateDefault(),
            SlideTheme.Light => CreateLight(),
            SlideTheme.Dark => CreateDark(),
            SlideTheme.Corporate => CreateCorporate(),
            SlideTheme.Vibrant => CreateVibrant(),
            _ => CreateDefault()
        };
    }

    // ── Default ──────────────────────────────────────────────
    // Dark navy title bar, light body, Segoe UI
    private static SlideStyleConfiguration CreateDefault()
    {
        return new SlideStyleConfiguration
        {
            DefaultFontName = "Segoe UI",
            TitleFontName = "Segoe UI Semibold",
            CodeFontName = "Cascadia Code",
            TitleColor = "FFFFFF",
            BodyColor = "2D2D2D",
            AccentColor = "1B3A5C",
            AccentColor2 = "2E86DE",
            SlideBackgroundColor = "FAFBFC",
            TitleBarColor = "1B3A5C",
            BorderColor = "DEE2E6",
            TableHeaderColor = "1B3A5C",
            TableHeaderTextColor = "FFFFFF",
            TableStripeColor = "F0F4F8",
            CodeBackgroundColor = "1E1E2E",
            CodeForegroundColor = "CDD6F4",
            SyntaxColorScheme = DarkCodeScheme()
        };
    }

    // ── Light ────────────────────────────────────────────────
    // Crisp white, subtle gray accents, black text
    private static SlideStyleConfiguration CreateLight()
    {
        return new SlideStyleConfiguration
        {
            DefaultFontName = "Calibri",
            TitleFontName = "Calibri Light",
            CodeFontName = "Consolas",
            TitleColor = "1A1A1A",
            BodyColor = "333333",
            AccentColor = "4472C4",
            AccentColor2 = "4472C4",
            SlideBackgroundColor = "FFFFFF",
            TitleBarColor = "F2F2F2",
            BorderColor = "D9D9D9",
            TableHeaderColor = "F2F2F2",
            TableHeaderTextColor = "1A1A1A",
            TableStripeColor = "FAFAFA",
            CodeBackgroundColor = "F5F5F5",
            CodeForegroundColor = "1A1A1A",
            SyntaxColorScheme = LightCodeScheme()
        };
    }

    // ── Dark ─────────────────────────────────────────────────
    // Dark backgrounds, light text, moody palette
    private static SlideStyleConfiguration CreateDark()
    {
        return new SlideStyleConfiguration
        {
            DefaultFontName = "Segoe UI",
            TitleFontName = "Segoe UI Semibold",
            CodeFontName = "Cascadia Code",
            TitleColor = "E0E0E0",
            BodyColor = "CCCCCC",
            AccentColor = "BB86FC",
            AccentColor2 = "03DAC6",
            SlideBackgroundColor = "1E1E1E",
            TitleBarColor = "121212",
            BorderColor = "333333",
            TableHeaderColor = "2D2D2D",
            TableHeaderTextColor = "E0E0E0",
            TableStripeColor = "252525",
            CodeBackgroundColor = "0D1117",
            CodeForegroundColor = "E6EDF3",
            SyntaxColorScheme = DarkCodeScheme()
        };
    }

    // ── Corporate ────────────────────────────────────────────
    // Conservative blue/gray, Arial, boardroom-ready
    private static SlideStyleConfiguration CreateCorporate()
    {
        return new SlideStyleConfiguration
        {
            DefaultFontName = "Arial",
            TitleFontName = "Arial",
            CodeFontName = "Consolas",
            TitleColor = "FFFFFF",
            BodyColor = "333333",
            AccentColor = "003366",
            AccentColor2 = "336699",
            SlideBackgroundColor = "FFFFFF",
            TitleBarColor = "003366",
            BorderColor = "B0B0B0",
            TableHeaderColor = "003366",
            TableHeaderTextColor = "FFFFFF",
            TableStripeColor = "EBF0F5",
            CodeBackgroundColor = "F0F0F0",
            CodeForegroundColor = "1A1A1A",
            SyntaxColorScheme = LightCodeScheme()
        };
    }

    // ── Vibrant ──────────────────────────────────────────────
    // Bold gradient feel, high-contrast, energetic
    private static SlideStyleConfiguration CreateVibrant()
    {
        return new SlideStyleConfiguration
        {
            DefaultFontName = "Segoe UI",
            TitleFontName = "Segoe UI Bold",
            CodeFontName = "Cascadia Code",
            TitleColor = "FFFFFF",
            BodyColor = "2D2D2D",
            AccentColor = "6C3483",
            AccentColor2 = "E74C3C",
            SlideBackgroundColor = "FDFEFE",
            TitleBarColor = "6C3483",
            BorderColor = "D5D8DC",
            TableHeaderColor = "6C3483",
            TableHeaderTextColor = "FFFFFF",
            TableStripeColor = "F4ECF7",
            CodeBackgroundColor = "2C2C54",
            CodeForegroundColor = "EAECEE",
            SyntaxColorScheme = DarkCodeScheme()
        };
    }

    // ── Syntax color schemes ─────────────────────────────────

    private static SyntaxColorScheme DarkCodeScheme()
    {
        return new SyntaxColorScheme
        {
            KeywordColor = "CBA6F7",
            StringColor = "A6E3A1",
            NumberColor = "FAB387",
            CommentColor = "6C7086",
            OperatorColor = "89DCEB",
            TypeColor = "F9E2AF",
            FunctionColor = "89B4FA",
            PropertyColor = "74C7EC",
            IdentifierColor = "CDD6F4",
            DefaultColor = "CDD6F4"
        };
    }

    private static SyntaxColorScheme LightCodeScheme()
    {
        return new SyntaxColorScheme
        {
            KeywordColor = "0000FF",
            StringColor = "A31515",
            NumberColor = "098658",
            CommentColor = "008000",
            OperatorColor = "383838",
            TypeColor = "267F99",
            FunctionColor = "795E26",
            PropertyColor = "001080",
            IdentifierColor = "1A1A1A",
            DefaultColor = "1A1A1A"
        };
    }
}

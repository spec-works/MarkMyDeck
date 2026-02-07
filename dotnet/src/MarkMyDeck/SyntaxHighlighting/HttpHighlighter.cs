using System;
using System.Collections.Generic;
using System.Linq;

namespace MarkMyDeck.SyntaxHighlighting;

/// <summary>
/// Syntax highlighter for HTTP requests and responses with media-type-aware body highlighting.
/// </summary>
public class HttpHighlighter : ISyntaxHighlighter
{
    private readonly ColorCodeHighlighter _colorCodeHighlighter;

    private static readonly HashSet<string> HttpMethods = new(StringComparer.OrdinalIgnoreCase)
    {
        "GET", "POST", "PUT", "DELETE", "PATCH", "HEAD", "OPTIONS", "TRACE", "CONNECT"
    };

    private static readonly Dictionary<string, string> MediaTypeToLanguage = new()
    {
        { "application/json", "json" },
        { "application/xml", "xml" },
        { "text/xml", "xml" },
        { "text/html", "html" },
        { "text/plain", "plain" },
        { "application/javascript", "javascript" },
        { "text/javascript", "javascript" },
        { "application/typescript", "typescript" },
        { "application/x-www-form-urlencoded", "plain" }
    };

    public HttpHighlighter()
    {
        _colorCodeHighlighter = new ColorCodeHighlighter();
    }

    public bool SupportsLanguage(string language)
    {
        return language?.Equals("http", StringComparison.OrdinalIgnoreCase) == true ||
               language?.Equals("https", StringComparison.OrdinalIgnoreCase) == true ||
               language?.Equals("request", StringComparison.OrdinalIgnoreCase) == true ||
               language?.Equals("response", StringComparison.OrdinalIgnoreCase) == true;
    }

    public IEnumerable<SyntaxToken> Highlight(string code, string language)
    {
        if (string.IsNullOrEmpty(code))
            yield break;

        var message = ParseHttpMessage(code);
        if (message == null)
        {
            yield return new SyntaxToken(code, TokenType.Default);
            yield break;
        }

        foreach (var token in GenerateTokens(message))
            yield return token;
    }

    private HttpMessage? ParseHttpMessage(string code)
    {
        var lines = code.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
        if (lines.Length == 0) return null;

        var firstLine = lines[0].Trim();
        HttpMessage message;

        if (IsRequestLine(firstLine))
            message = ParseRequestLine(firstLine);
        else if (IsStatusLine(firstLine))
            message = ParseStatusLine(firstLine);
        else
            return null;

        int lineIndex = 1;
        var headers = new List<HttpHeader>();
        while (lineIndex < lines.Length)
        {
            var line = lines[lineIndex];
            if (string.IsNullOrWhiteSpace(line)) { lineIndex++; break; }
            var colonIndex = line.IndexOf(':');
            if (colonIndex > 0)
            {
                var name = line.Substring(0, colonIndex);
                var value = colonIndex + 1 < line.Length ? line.Substring(colonIndex + 1).TrimStart() : string.Empty;
                headers.Add(new HttpHeader(name, value));
            }
            lineIndex++;
        }

        message.Headers = headers;
        if (lineIndex < lines.Length)
            message.Body = string.Join(Environment.NewLine, lines.Skip(lineIndex));

        var contentTypeHeader = headers.FirstOrDefault(h => h.Name.Equals("Content-Type", StringComparison.OrdinalIgnoreCase));
        if (contentTypeHeader != null)
            message.ContentType = ParseContentType(contentTypeHeader.Value);

        return message;
    }

    private bool IsRequestLine(string line)
    {
        var parts = line.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        return parts.Length >= 3 && HttpMethods.Contains(parts[0]) && parts[2].StartsWith("HTTP/", StringComparison.OrdinalIgnoreCase);
    }

    private bool IsStatusLine(string line)
    {
        return line.StartsWith("HTTP/", StringComparison.OrdinalIgnoreCase) && line.Length > 9 && char.IsDigit(line[9]);
    }

    private HttpMessage ParseRequestLine(string line)
    {
        var parts = line.Split(' ', 3, StringSplitOptions.RemoveEmptyEntries);
        return new HttpMessage { IsRequest = true, Method = parts.Length > 0 ? parts[0] : "", Url = parts.Length > 1 ? parts[1] : "", Version = parts.Length > 2 ? parts[2] : "" };
    }

    private HttpMessage ParseStatusLine(string line)
    {
        var parts = line.Split(' ', 3, StringSplitOptions.RemoveEmptyEntries);
        return new HttpMessage { IsRequest = false, Version = parts.Length > 0 ? parts[0] : "", StatusCode = parts.Length > 1 ? parts[1] : "", ReasonPhrase = parts.Length > 2 ? parts[2] : "" };
    }

    private string? ParseContentType(string contentTypeValue)
    {
        var semicolonIndex = contentTypeValue.IndexOf(';');
        return (semicolonIndex > 0 ? contentTypeValue.Substring(0, semicolonIndex).Trim() : contentTypeValue.Trim()).ToLowerInvariant();
    }

    private IEnumerable<SyntaxToken> GenerateTokens(HttpMessage message)
    {
        if (message.IsRequest)
        {
            yield return new SyntaxToken(message.Method, TokenType.Keyword);
            yield return new SyntaxToken(" ", TokenType.Default);
            yield return new SyntaxToken(message.Url, TokenType.String);
            yield return new SyntaxToken(" ", TokenType.Default);
            yield return new SyntaxToken(message.Version, TokenType.Type);
            yield return new SyntaxToken(Environment.NewLine, TokenType.Default);
        }
        else
        {
            yield return new SyntaxToken(message.Version, TokenType.Type);
            yield return new SyntaxToken(" ", TokenType.Default);
            yield return new SyntaxToken(message.StatusCode, TokenType.Number);
            yield return new SyntaxToken(" ", TokenType.Default);
            yield return new SyntaxToken(message.ReasonPhrase, TokenType.Default);
            yield return new SyntaxToken(Environment.NewLine, TokenType.Default);
        }

        foreach (var header in message.Headers)
        {
            yield return new SyntaxToken(header.Name, TokenType.Property);
            yield return new SyntaxToken(":", TokenType.Operator);
            yield return new SyntaxToken(" ", TokenType.Default);
            yield return new SyntaxToken(header.Value, TokenType.Default);
            yield return new SyntaxToken(Environment.NewLine, TokenType.Default);
        }

        if (!string.IsNullOrEmpty(message.Body))
        {
            yield return new SyntaxToken(Environment.NewLine, TokenType.Default);

            string? bodyLanguage = null;
            if (message.ContentType != null && MediaTypeToLanguage.TryGetValue(message.ContentType, out var lang))
                bodyLanguage = lang;

            if (bodyLanguage != null && !bodyLanguage.Equals("plain", StringComparison.OrdinalIgnoreCase) && _colorCodeHighlighter.SupportsLanguage(bodyLanguage))
            {
                foreach (var token in _colorCodeHighlighter.Highlight(message.Body, bodyLanguage))
                    yield return token;
            }
            else
            {
                yield return new SyntaxToken(message.Body, TokenType.Default);
            }
        }
    }

    private class HttpMessage
    {
        public bool IsRequest { get; set; }
        public string Method { get; set; } = "";
        public string Url { get; set; } = "";
        public string StatusCode { get; set; } = "";
        public string ReasonPhrase { get; set; } = "";
        public string Version { get; set; } = "";
        public List<HttpHeader> Headers { get; set; } = new();
        public string? Body { get; set; }
        public string? ContentType { get; set; }
    }

    private class HttpHeader
    {
        public string Name { get; }
        public string Value { get; }
        public HttpHeader(string name, string value) { Name = name; Value = value; }
    }
}

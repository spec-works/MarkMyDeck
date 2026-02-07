using System;
using System.Collections.Generic;

namespace MarkMyDeck.SyntaxHighlighting;

/// <summary>
/// Custom syntax highlighter for TypeSpec language.
/// </summary>
public class TypeSpecHighlighter : ISyntaxHighlighter
{
    private static readonly HashSet<string> KeywordSet = new(StringComparer.OrdinalIgnoreCase)
    {
        "model", "namespace", "op", "interface", "enum", "union", "using", "import",
        "extends", "is", "alias", "scalar", "void", "never", "unknown", "true", "false",
        "if", "else", "return", "valueof", "typeof"
    };

    private static readonly HashSet<string> TypeSet = new(StringComparer.OrdinalIgnoreCase)
    {
        "string", "int8", "int16", "int32", "int64", "uint8", "uint16", "uint32", "uint64",
        "safeint", "float", "float32", "float64", "decimal", "decimal128", "numeric",
        "integer", "boolean", "bytes", "duration", "plainDate", "plainTime", "utcDateTime",
        "offsetDateTime", "url", "Record", "Array"
    };

    public bool SupportsLanguage(string language)
    {
        return language?.Equals("typespec", StringComparison.OrdinalIgnoreCase) == true ||
               language?.Equals("cadl", StringComparison.OrdinalIgnoreCase) == true;
    }

    public IEnumerable<SyntaxToken> Highlight(string code, string language)
    {
        if (string.IsNullOrEmpty(code))
            yield break;

        int position = 0;
        while (position < code.Length)
        {
            if (char.IsWhiteSpace(code[position]))
            {
                int start = position;
                while (position < code.Length && char.IsWhiteSpace(code[position])) position++;
                yield return new SyntaxToken(code.Substring(start, position - start), TokenType.Default);
                continue;
            }

            if (position < code.Length - 1 && code[position] == '/' && code[position + 1] == '*')
            {
                int start = position;
                position += 2;
                while (position < code.Length - 1)
                {
                    if (code[position] == '*' && code[position + 1] == '/') { position += 2; break; }
                    position++;
                }
                if (position >= code.Length) position = code.Length;
                yield return new SyntaxToken(code.Substring(start, position - start), TokenType.Comment);
                continue;
            }

            if (position < code.Length - 1 && code[position] == '/' && code[position + 1] == '/')
            {
                int start = position;
                while (position < code.Length && code[position] != '\n' && code[position] != '\r') position++;
                yield return new SyntaxToken(code.Substring(start, position - start), TokenType.Comment);
                continue;
            }

            if (code[position] == '"')
            {
                int start = position;
                position++;
                while (position < code.Length)
                {
                    if (code[position] == '\\' && position + 1 < code.Length) { position += 2; continue; }
                    if (code[position] == '"') { position++; break; }
                    position++;
                }
                yield return new SyntaxToken(code.Substring(start, position - start), TokenType.String);
                continue;
            }

            if (code[position] == '`')
            {
                int start = position;
                position++;
                while (position < code.Length)
                {
                    if (code[position] == '\\' && position + 1 < code.Length) { position += 2; continue; }
                    if (code[position] == '`') { position++; break; }
                    position++;
                }
                yield return new SyntaxToken(code.Substring(start, position - start), TokenType.String);
                continue;
            }

            if (char.IsDigit(code[position]) ||
                (code[position] == '-' && position + 1 < code.Length && char.IsDigit(code[position + 1])))
            {
                int start = position;
                if (code[position] == '-') position++;
                while (position < code.Length && (char.IsDigit(code[position]) || code[position] == '.' ||
                       code[position] == 'e' || code[position] == 'E' || code[position] == '-' || code[position] == '+'))
                    position++;
                yield return new SyntaxToken(code.Substring(start, position - start), TokenType.Number);
                continue;
            }

            if (code[position] == '@')
            {
                int start = position;
                position++;
                while (position < code.Length && (char.IsLetterOrDigit(code[position]) || code[position] == '_'))
                    position++;
                yield return new SyntaxToken(code.Substring(start, position - start), TokenType.Property);
                continue;
            }

            if (char.IsLetter(code[position]) || code[position] == '_')
            {
                int start = position;
                while (position < code.Length && (char.IsLetterOrDigit(code[position]) || code[position] == '_'))
                    position++;
                string identifier = code.Substring(start, position - start);
                TokenType tokenType;
                if (KeywordSet.Contains(identifier)) tokenType = TokenType.Keyword;
                else if (TypeSet.Contains(identifier)) tokenType = TokenType.Type;
                else if (position < code.Length && code[position] == '(') tokenType = TokenType.Function;
                else if (char.IsUpper(identifier[0])) tokenType = TokenType.Type;
                else tokenType = TokenType.Identifier;
                yield return new SyntaxToken(identifier, tokenType);
                continue;
            }

            if ("{}[]()<>:;,.?|&=+-*/%!".Contains(code[position]))
            {
                yield return new SyntaxToken(code[position].ToString(), TokenType.Operator);
                position++;
                continue;
            }

            yield return new SyntaxToken(code[position].ToString(), TokenType.Default);
            position++;
        }
    }
}

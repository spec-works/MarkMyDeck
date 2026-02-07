using System.Linq;
using Markdig.Extensions.Tables;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using D = DocumentFormat.OpenXml.Drawing;
using MarkdigTable = Markdig.Extensions.Tables.Table;
using MarkdigTableRow = Markdig.Extensions.Tables.TableRow;
using MarkdigTableCell = Markdig.Extensions.Tables.TableCell;

namespace MarkMyDeck.Converters.BlockRenderers;

/// <summary>
/// Renderer for table blocks.
/// </summary>
public class TableRenderer : OpenXmlObjectRenderer<MarkdigTable>
{
    protected override void Write(OpenXmlPresentationRenderer renderer, MarkdigTable table)
    {
        var slide = renderer.CurrentSlide;
        var styles = slide.Styles;

        // Count rows and columns
        int rowCount = table.Count;
        int colCount = 0;
        foreach (var row in table)
        {
            if (row is MarkdigTableRow tableRow && tableRow.Count > colCount)
                colCount = tableRow.Count;
        }

        if (rowCount == 0 || colCount == 0) return;

        // 1 point = 12700 EMU; row height ~ 1.6x font size
        var rowHeight = (long)(styles.DefaultFontSize * 12700 * 1.6);
        var totalHeight = rowHeight * rowCount;

        var drawingTable = slide.AddTable(rowCount, colCount, totalHeight);

        foreach (var row in table)
        {
            if (row is MarkdigTableRow tableRow)
            {
                var drawingRow = new D.TableRow { Height = rowHeight };

                foreach (var cell in tableRow)
                {
                    if (cell is MarkdigTableCell tableCell)
                    {
                        var drawingCell = new D.TableCell();

                        var textBody = new D.TextBody(
                            new D.BodyProperties(),
                            new D.ListStyle());

                        var paragraph = new D.Paragraph();

                        // Render cell content
                        foreach (var block in tableCell)
                        {
                            if (block is ParagraphBlock paragraphBlock && paragraphBlock.Inline != null)
                            {
                                var inline = paragraphBlock.Inline.FirstChild;
                                while (inline != null)
                                {
                                    if (inline is LiteralInline literal)
                                    {
                                        var textColor = tableRow.IsHeader ? styles.TableHeaderTextColor : styles.BodyColor;
                                        var run = slide.CreateRun(
                                            literal.Content.ToString(),
                                            styles.DefaultFontName,
                                            styles.DefaultFontSize,
                                            textColor,
                                            bold: tableRow.IsHeader);
                                        paragraph.Append(run);
                                    }
                                    else if (inline is EmphasisInline emphasis)
                                    {
                                        var child = emphasis.FirstChild;
                                        while (child != null)
                                        {
                                            if (child is LiteralInline emphLiteral)
                                            {
                                                var run = slide.CreateRun(
                                                    emphLiteral.Content.ToString(),
                                                    styles.DefaultFontName,
                                                    styles.DefaultFontSize,
                                                    styles.BodyColor,
                                                    bold: emphasis.DelimiterCount == 2 || tableRow.IsHeader,
                                                    italic: emphasis.DelimiterCount == 1);
                                                paragraph.Append(run);
                                            }
                                            child = child.NextSibling;
                                        }
                                    }
                                    else if (inline is CodeInline code)
                                    {
                                        var run = slide.CreateRun(
                                            code.Content,
                                            styles.CodeFontName,
                                            styles.CodeFontSize,
                                            styles.BodyColor,
                                            bold: tableRow.IsHeader);
                                        paragraph.Append(run);
                                    }
                                    inline = inline.NextSibling;
                                }
                            }
                        }

                        // Ensure at least an end paragraph run
                        if (!paragraph.Elements<D.Run>().Any())
                        {
                            paragraph.Append(new D.EndParagraphRunProperties { Language = "en-US" });
                        }

                        textBody.Append(paragraph);
                        drawingCell.Append(textBody);

                        // Cell properties with clean borders
                        var cellProps = new D.TableCellProperties();
                        var borderColor = styles.BorderColor;

                        cellProps.Append(new D.LeftBorderLineProperties(
                            new D.SolidFill(new D.RgbColorModelHex { Val = borderColor })) { Width = 6350 });
                        cellProps.Append(new D.RightBorderLineProperties(
                            new D.SolidFill(new D.RgbColorModelHex { Val = borderColor })) { Width = 6350 });
                        cellProps.Append(new D.TopBorderLineProperties(
                            new D.SolidFill(new D.RgbColorModelHex { Val = borderColor })) { Width = 6350 });
                        cellProps.Append(new D.BottomBorderLineProperties(
                            new D.SolidFill(new D.RgbColorModelHex { Val = borderColor })) { Width = 6350 });

                        if (tableRow.IsHeader)
                        {
                            cellProps.Append(new D.SolidFill(new D.RgbColorModelHex { Val = styles.TableHeaderColor }));
                        }
                        else if (drawingTable.Elements<D.TableRow>().Count() % 2 == 0)
                        {
                            cellProps.Append(new D.SolidFill(new D.RgbColorModelHex { Val = styles.TableStripeColor }));
                        }

                        drawingCell.Append(cellProps);
                        drawingRow.Append(drawingCell);
                    }
                }

                // Pad row to correct column count
                while (drawingRow.Elements<D.TableCell>().Count() < colCount)
                {
                    var emptyCell = new D.TableCell(
                        new D.TextBody(
                            new D.BodyProperties(),
                            new D.ListStyle(),
                            new D.Paragraph(new D.EndParagraphRunProperties { Language = "en-US" })),
                        new D.TableCellProperties());
                    drawingRow.Append(emptyCell);
                }

                drawingTable.Append(drawingRow);
            }
        }

        renderer.CurrentShape = null;
        renderer.CurrentParagraph = null;
    }
}

using System.Collections.Generic;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

sealed class CodeElementExpression(IHtmlElement node) : BlockElementExpression(node)
{
    /// <inheritdoc/>
    public override IEnumerable<OpenXmlElement> Interpret(ParsingContext context)
    {
        ComposeStyles(context);
        var childContext = context.CreateChild(this);
        childContext.PreserveLinebreaks = true;
        childContext.CollapseWhitespaces = false;
        var childElements = Interpret(childContext, node.ChildNodes);

        TableCell cell;
        Table preTable = new(
            new TableProperties {
                TableStyle = context.DocumentStyle.GetTableStyle(context.DocumentStyle.DefaultStyles.PreTableStyle),
                TableWidth = new() { Type = TableWidthUnitValues.Auto, Width = "0" } // 100%
            },
            new TableGrid(
                new GridColumn() { Width = "5610" }),
            new TableRow(
                cell = new TableCell {
                    // Ensure the border lines are visible (regardless of the style used)
                    TableCellProperties = new() {
                        TableCellBorders = new TableCellBorders {
                            TopBorder = new TopBorder() { Val = BorderValues.Single },
                            LeftBorder = new LeftBorder() { Val = BorderValues.Single },
                            BottomBorder = new BottomBorder() { Val = BorderValues.Single },
                            RightBorder = new RightBorder() { Val = BorderValues.Single }
                        }
                    },
                })
        );

        cell.Append(childElements);

        return [preTable];
    }
}

using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Svg
{
    //internal class SvgRange
    //{
    //    List<RenderItem> renderItems = new List<RenderItem>();
    //    //List<TextBox> textBoxes = new List<TextBox>();

    //    internal SvgRange(ExcelRange range, double totalWidth, double totalHeight)
    //    {
    //        SvgRenderRectItem rangeBB = new SvgRenderRectItem();

    //        //To pixel multiplier
    //        float mult = 96f / 72f;

    //        rangeBB.Width = (float)totalWidth;
    //        rangeBB.Height = (float)totalHeight * mult;

    //        rangeBB.BorderColor = "yellow";

    //        rangeBB.BorderWidth = 1;

    //        renderItems.Add(rangeBB);

    //        float currentWidth = 0f;

    //        for (int i = 0; i < range.Columns; i++)
    //        {
    //            float currentHeight = 0f;
    //            var currCol = range.Worksheet.GetColumn(range._fromCol + i);
    //            var colWidth = currCol == null ? range.Worksheet.DefaultColWidth : currCol.Width;
    //            colWidth = ExcelColumn.ColumnWidthToPixels(colWidth, range.Worksheet.Workbook.MaxFontWidth);

    //            for (int j = 0; j < range.Rows; j++)
    //            {
    //                var cell = range.Offset(j, i);

    //                var cellContent = range.Offset(j, i).TextForWidth;
    //                float heightAlt = (float)cell.Worksheet.GetRowHeight(cell._fromRow + j);
    //                float heightAltPixels = heightAlt * mult;
    //                float height = (float)cell.Worksheet.Rows[cell._fromRow].Height * mult;

    //                SvgRenderRectItem cellBB = new SvgRenderRectItem();
    //                cellBB.Left = currentWidth;
    //                cellBB.Top = currentHeight;

    //                cellBB.Width = (float)colWidth;
    //                cellBB.Height = (float)height;

    //                cellBB.BorderWidth = 1;

    //                cellBB.BorderColor = "black";
    //                cellBB.FillColor = "gray";

    //                renderItems.Add(cellBB);

    //                var cellTextBox = new TextBox(null, currentWidth, currentHeight, colWidth, height);
    //                cellTextBox.AddCellTextRun(cell);

    //                var deltaHeight = (float)cellTextBox.Bounds.Height;
    //                cellBB.Height = cellBB.Height < deltaHeight ? deltaHeight : cellBB.Height;
    //                textBoxes.Add(cellTextBox);

    //                currentHeight += height;
    //            }
    //            currentWidth += (float)colWidth;
    //        }

    //    }

    //    internal void Render(StringBuilder sb)
    //    {
    //        foreach (var item in renderItems)
    //        {
    //            item.Render(sb);
    //        }
    //        foreach (var item in textBoxes)
    //        {
    //            item.RenderTextRuns(sb);
    //        }
    //    }
    //}
}
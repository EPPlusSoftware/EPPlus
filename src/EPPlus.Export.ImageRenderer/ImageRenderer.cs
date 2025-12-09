/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/

using EPPlus.Export.ImageRenderer;
using EPPlus.Export.ImageRenderer.Svg;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Text;

namespace EPPlusImageRenderer
{
    public class ImageRenderer
    {
        public string RenderDrawingToSvg(ExcelDrawing drawing)
        {
            drawing.GetSizeInPixels(out int width, out int height);
            var sb = new StringBuilder();
            if (drawing is ExcelShape shape)
            {
                var svg = new SvgShape(shape);
                svg.Size = new DrawingSize(width, height);
                svg.Render(sb);
                return sb.ToString();
            }
            else if(drawing is ExcelChart chart)
            {
                var svg = new SvgChart(chart);
                svg.Size = new DrawingSize(width, height);
                svg.Render(sb);
                return sb.ToString();
            }

            throw new NotImplementedException("Image rendering for drawing type not implemented.");
        }
        public string RenderRangeToSvg(ExcelRange range)
        {
            var ws = range.Worksheet;

            //ws.Workbook.Styles.CellXfs.

            double totalWidth = 0;
            foreach (var col in range.EntireColumn)
            {
                totalWidth += ExcelColumn.ColumnWidthToPixels(col.Width, range.Worksheet.Workbook.MaxFontWidth);
            }
            double totalHeight = 0;
            foreach (var row in range.EntireRow)
            {
                totalHeight = row.Height;
            }

            var sRange = new SvgRange(range, totalWidth, totalHeight);

            var sb = new StringBuilder();
            sb.Append($"<svg width=\"{500}\" height=\"{500}\" xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xml:space=\"preserve\" Overflow=\"Hidden\">");
            sRange.Render(sb);
            sb.Append($"</svg>");
            return sb.ToString();
        }

        public string RenderBox(double l, double r, double b, double t)
        {
            var doc = new SvgEpplusDocument();

        }
    }
}

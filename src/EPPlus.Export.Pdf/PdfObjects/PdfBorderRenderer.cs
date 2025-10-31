/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
using EPPlus.Export.Pdf.Pdfhelpers;
using EPPlus.Export.Pdf.PdfLayout;
using OfficeOpenXml.Style;

namespace EPPlus.Export.Pdf.PdfObjects
{
    /// <summary>
    /// Enum that describes if the line used for pdf border.
    /// </summary>
    public enum LineType
    {
        /// <summary>
        /// Line is horizontal.
        /// </summary>
        Horizontal = 0,
        /// <summary>
        /// Line is vertical.
        /// </summary>
        Vertical,
        /// <summary>
        /// Line is diagonal going down up.
        /// </summary>
        DiagonalUp,
        /// <summary>
        /// Line is diagonal going up down.
        /// </summary>
        DiagonalDown
    }
    internal class PdfBorderRenderer
    {
        private const double Hair   = 0.5d;
        private const double Thin   = 0.85d;
        private const double Small  = 1.1d;
        private const double Medium = 1.5d;
        private const double Thick  = 2.0d;

        public void RenderBorder(PdfContentStream contentStream, PdfCellBorderData borderData, LineType lineType, double x1, double y1, double x2, double y2)
        {
            switch (borderData.BorderStyle)
            {
                case ExcelBorderStyle.None:
                    return;
                case ExcelBorderStyle.Hair:
                    DrawBasicBorder(contentStream, borderData, lineType, Hair, "[] 0 d");
                    break;
                case ExcelBorderStyle.Dotted:
                    DrawBasicBorder(contentStream, borderData, lineType, Small, "[0 2] 0 d");
                    break;
                case ExcelBorderStyle.DashDot:
                    DrawBasicBorder(contentStream, borderData, lineType, Small, "[4 2 1 2] 0 d");
                    break;
                case ExcelBorderStyle.Thin:
                    DrawBasicBorder(contentStream, borderData, lineType, Thin, "[] 0 d");
                    break;
                case ExcelBorderStyle.DashDotDot:
                    DrawBasicBorder(contentStream, borderData, lineType, Small, "[4 2 1 2 1 2] 0 d");
                    break;
                case ExcelBorderStyle.Dashed:
                    DrawBasicBorder(contentStream, borderData, lineType, Small, "[4 3] 0 d");
                    break;
                case ExcelBorderStyle.MediumDashDotDot:
                    DrawBasicBorder(contentStream, borderData, lineType, Medium, "[6 3 2 3 2 3] 0 d");
                    break;
                case ExcelBorderStyle.MediumDashed:
                    DrawBasicBorder(contentStream, borderData, lineType, Medium, "[6 4] 0 d");
                    break;
                case ExcelBorderStyle.MediumDashDot:
                    DrawBasicBorder(contentStream, borderData, lineType, Medium, "[6 3 2 3] 0 d");
                    break;
                case ExcelBorderStyle.Thick:
                    DrawBasicBorder(contentStream, borderData, lineType, Thick, "[] 0 d");
                    break;
                case ExcelBorderStyle.Medium:
                    DrawBasicBorder(contentStream, borderData, lineType, Medium, "[] 0 d");
                    break;
                case ExcelBorderStyle.SlantDashDot:
                    DrawSlantDashDotBorder(contentStream, borderData, lineType, x1, y1, x2, y2);
                    return;
                case ExcelBorderStyle.Double:
                    DrawDoubleBorder(contentStream, borderData, lineType, x1, y1, x2, y2);
                    return;
            }
            contentStream.AddCommand($"{x1.ToPdfStringF4()} {y1.ToPdfStringF4()} m");
            contentStream.AddCommand($"{x2.ToPdfStringF4()} {y2.ToPdfStringF4()} l");
            contentStream.AddCommand("S");
        }

        private void DrawBasicBorder(PdfContentStream contentStream, PdfCellBorderData borderData, LineType lineType, double width, string dash)
        {
            contentStream.AddCommand(borderData.BorderColor.ToStrokeCommand());
            contentStream.AddCommand($"{width.ToPdfString()} w");
            contentStream.AddCommand(borderData.BorderStyle != ExcelBorderStyle.Dotted ? lineType == LineType.DiagonalUp || lineType == LineType.DiagonalDown ? "0 J" : "2 J" : "1 J");
            contentStream.AddCommand(dash);
        }

        //This one could be made to look more fancy
        private void DrawDoubleBorder(PdfContentStream contentStream, PdfCellBorderData borderData, LineType lineType, double x1, double y1, double x2, double y2)
        {
            double offsetX = borderData.DoubleBorderOffsets.X;
            double offsetY = borderData.DoubleBorderOffsets.Y;
            contentStream.AddCommand(borderData.BorderColor.ToStrokeCommand());
            contentStream.AddCommand($"{Small.ToPdfString()} w");
            contentStream.AddCommand("[] 0 d");
            contentStream.AddCommand(lineType == LineType.DiagonalUp || lineType == LineType.DiagonalDown ? "0 J" : "2 J");
            if (lineType == LineType.DiagonalUp)
            {
                contentStream.AddCommand($"{(x1 + offsetX + offsetX).ToPdfString()} {(y1 + offsetY).ToPdfString()} m");
                contentStream.AddCommand($"{(x2 + -offsetX).ToPdfString()} {(y2 + -offsetY - offsetY).ToPdfString()} l");
            }
            else if (lineType == LineType.DiagonalDown)
            {
                contentStream.AddCommand($"{(x1 + offsetX + offsetX).ToPdfString()} {(y1 + -offsetY).ToPdfString()} m");
                contentStream.AddCommand($"{(x2 + -offsetX).ToPdfString()} {(y2 + offsetY + offsetY).ToPdfString()} l");
            }
            else
            {
                contentStream.AddCommand($"{x1.ToPdfString()} {y1.ToPdfString()} m");
                contentStream.AddCommand($"{x2.ToPdfString()} {y2.ToPdfString()} l");
            }
            contentStream.AddCommand("S");
            contentStream.AddCommand($"{Small.ToPdfString()} w");
            contentStream.AddCommand("[] 0 d");
            contentStream.AddCommand(lineType == LineType.DiagonalUp || lineType == LineType.DiagonalDown ? "0 J" : "2 J");
            if (lineType == LineType.Vertical)
            {
                contentStream.AddCommand($"{(x1 + offsetX).ToPdfString()} {(y1 + -offsetY).ToPdfString()} m");
                contentStream.AddCommand($"{(x2 + offsetX).ToPdfString()} {(y2 + offsetY).ToPdfString()} l");
            }
            else if (lineType == LineType.DiagonalUp)
            {
                contentStream.AddCommand($"{(x1 + offsetX).ToPdfString()} {(y1 + offsetY + offsetY).ToPdfString()} m");
                contentStream.AddCommand($"{(x2 + -offsetX - offsetX).ToPdfString()} {(y2 + -offsetY).ToPdfString()} l");
            }
            else if (lineType == LineType.DiagonalDown)
            {
                contentStream.AddCommand($"{(x1 + offsetX).ToPdfString()} {(y1 + -offsetY - offsetY).ToPdfString()} m");
                contentStream.AddCommand($"{(x2 + -offsetX - offsetX).ToPdfString()} {(y2 + offsetY).ToPdfString()} l");
            }
            else
            {
                contentStream.AddCommand($"{(x1 + offsetX).ToPdfString()} {(y1 + offsetY).ToPdfString()} m");
                contentStream.AddCommand($"{(x2 + -offsetX).ToPdfString()} {(y2 + offsetY).ToPdfString()} l");
            }
            contentStream.AddCommand("S");
        }

        private void DrawSlantDashDotBorder(PdfContentStream contentStream, PdfCellBorderData borderData, LineType lineType, double x1, double y1, double x2, double y2)
        {
            contentStream.AddCommand(borderData.BorderColor.ToStrokeCommand());
            contentStream.AddCommand("Q");
            contentStream.AddCommand("q");
            contentStream.AddCommand($"{Small.ToPdfString()} w");
            contentStream.AddCommand(lineType == LineType.DiagonalUp || lineType == LineType.DiagonalDown ? "0 J" : "2 J");
            contentStream.AddCommand("[4 2 1 2] 0 d");
            contentStream.AddCommand($"1 0 0.6 1 0 0 cm");
            //calculate new x and y
            var nx1 = x1 + y1 * 0.6d;
            var tx1 = nx1 - x1;
            var nx2 = x2 + y2 * 0.6d;
            var tx2 = nx2 - x2;
            contentStream.AddCommand($"{(x1 - tx1).ToPdfString()} {y1.ToPdfString()} m");
            contentStream.AddCommand($"{(x2 - tx2).ToPdfString()} {y2.ToPdfString()} l");
            contentStream.AddCommand("S");
            contentStream.AddCommand("Q");
            contentStream.AddCommand("q");
        }
    }
}

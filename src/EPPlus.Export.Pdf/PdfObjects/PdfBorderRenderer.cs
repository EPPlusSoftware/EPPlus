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
using EPPlus.Export.Pdf.PdfGraphics;
using EPPlus.Export.Pdf.Pdfhelpers;
using EPPlus.Export.Pdf.PdfLayout;
using OfficeOpenXml.Style;
using System;
using System.Security.Cryptography.Xml;

namespace EPPlus.Export.Pdf.PdfObjects
{
    internal class PdfBorderRenderer
    {
        private readonly PdfCellBorderData Top;
        private readonly PdfCellBorderData Bottom;
        private readonly PdfCellBorderData Left;
        private readonly PdfCellBorderData Right;
        private readonly PdfCellBorderData DiagonalUp;
        private readonly PdfCellBorderData DiagonalDown;

        private readonly double X;
        private readonly double Y;
        private readonly double Width;
        private readonly double Height;
        private readonly string Name;


        public PdfBorderRenderer(PdfCellBorderLayout cell)
        {
            X = cell.LocalPosition.X;
            Y = cell.LocalPosition.Y;
            Width = cell.Size.X;
            Height = cell.Size.Y;
            Name = cell.Name;
            Top = cell.BorderData.Top;
            Bottom = cell.BorderData.Bottom;
            Left = cell.BorderData.Left;
            Right = cell.BorderData.Right;
            DiagonalUp = cell.BorderData.DiagonalUp;
            DiagonalDown = cell.BorderData.DiagonalDown;
        }

        public void RenderBorder(PdfContentStream contentStream)
        {
            contentStream.AddCommand($"% Border Start: {Name}");
            contentStream.AddCommand("q");
            RenderBorder(contentStream, Top);
            RenderBorder(contentStream, Bottom);
            RenderBorder(contentStream, Left);
            RenderBorder(contentStream, Right);
            RenderBorder(contentStream, DiagonalUp);
            RenderBorder(contentStream, DiagonalDown);
            contentStream.AddCommand("Q");
            contentStream.AddCommand($"% Border End: {Name}");
        }

        private void RenderBorder(PdfContentStream contentStream, PdfCellBorderData border)
        {
            double x1 = X, y1 = Y, x2 = X, y2 = Y;
            switch (border.LineType)
            {
                case LineType.Top:
                    x1 = X;
                    y1 = Y + Height;
                    x2 = X + Width;
                    y2 = Y + Height;
                    break;
                case LineType.Bottom:
                    x1 = X;
                    y1 = Y;
                    x2 = X + Width;
                    y2 = Y;
                    break;
                case LineType.Left:
                    x1 = X;
                    y1 = Y;
                    x2 = X;
                    y2 = Y + Height;
                    break;
                case LineType.Right:
                    x1 = X + Width;
                    y1 = Y;
                    x2 = X + Width;
                    y2 = Y + Height;
                    break;
                case LineType.DiagonalUp:
                    x1 = X;
                    y1 = Y;
                    x2 = X + Width;
                    y2 = Y + Height;
                    break;
                case LineType.DiagonalDown:
                    x1 = X;
                    y1 = Y + Height;
                    x2 = X + Width;
                    y2 = Y;
                    break;
            }
            switch (border.BorderStyle)
            {
                case ExcelBorderStyle.None:
                    return;
                case ExcelBorderStyle.Hair:
                    DrawBasicBorder(contentStream, border, PdfCellBorderData.Hair, PdfCellBorderData.NoDash);
                    break;
                case ExcelBorderStyle.Dotted:
                    DrawBasicBorder(contentStream, border, PdfCellBorderData.Small, PdfCellBorderData.Dotted);
                    break;
                case ExcelBorderStyle.DashDot:
                    DrawBasicBorder(contentStream, border, PdfCellBorderData.Small, PdfCellBorderData.DashDot);
                    break;
                case ExcelBorderStyle.Thin:
                    DrawBasicBorder(contentStream, border, PdfCellBorderData.Thin, PdfCellBorderData.NoDash);
                    break;
                case ExcelBorderStyle.DashDotDot:
                    DrawBasicBorder(contentStream, border, PdfCellBorderData.Small, PdfCellBorderData.DashDotDot);
                    break;
                case ExcelBorderStyle.Dashed:
                    DrawBasicBorder(contentStream, border, PdfCellBorderData.Small, PdfCellBorderData.Dashed);
                    break;
                case ExcelBorderStyle.MediumDashDotDot:
                    DrawBasicBorder(contentStream, border, PdfCellBorderData.Medium, PdfCellBorderData.MediumDashDotDot);
                    break;
                case ExcelBorderStyle.MediumDashed:
                    DrawBasicBorder(contentStream, border, PdfCellBorderData.Medium, PdfCellBorderData.MediumDashed);
                    break;
                case ExcelBorderStyle.MediumDashDot:
                    DrawBasicBorder(contentStream, border, PdfCellBorderData.Medium, PdfCellBorderData.MediumDashDot);
                    break;
                case ExcelBorderStyle.Thick:
                    DrawBasicBorder(contentStream, border, PdfCellBorderData.Thick, PdfCellBorderData.NoDash);
                    break;
                case ExcelBorderStyle.Medium:
                    DrawBasicBorder(contentStream, border, PdfCellBorderData.Medium, PdfCellBorderData.NoDash);
                    break;
                case ExcelBorderStyle.SlantDashDot:
                    DrawSlantDashDotBorder(contentStream, border, x1, y1, x2, y2);
                    return;
                case ExcelBorderStyle.Double:
                    DrawDoubleBorder(contentStream, border, x1, y1, x2, y2);
                    return;
            }
            contentStream.AddCommand($"{x1.ToPdfStringF4()} {y1.ToPdfStringF4()} m");
            contentStream.AddCommand($"{x2.ToPdfStringF4()} {y2.ToPdfStringF4()} l");
            contentStream.AddCommand("S");
        }

        private void DrawBasicBorder(PdfContentStream contentStream, PdfCellBorderData border, double width, string dash)
        {
            contentStream.AddCommand(border.BorderColor.ToStrokeCommand());
            contentStream.AddCommand($"{width.ToPdfString()} w");
            contentStream.AddCommand(border.BorderStyle != ExcelBorderStyle.Dotted ? ( border.LineType == LineType.DiagonalUp || border.LineType == LineType.DiagonalDown ? "0 J" : "2 J" ) : "1 J");
            contentStream.AddCommand(dash);
        }

        private void DrawDoubleBorder(PdfContentStream contentStream, PdfCellBorderData border, double x1, double y1, double x2, double y2)
        {
            var ix1 = x1;
            var ix2 = x2;
            var iy1 = y1;
            var iy2 = y2;
            var ox1 = x1;
            var ox2 = x2;
            var oy1 = y1;
            var oy2 = y2;

            var DiagonalUpFactor = 0d;
            var DiagonalDownFactor = 0d;

            if (border.LineType == LineType.Top)
            {
                ix1 = x1 + 0.7d;
                ix2 = x2 - 0.7d;
                iy1 = y1 - (PdfCellBorderData.Hair / 0.65d);
                iy2 = y2 - (PdfCellBorderData.Hair / 0.65d);
                if (DiagonalUp.BorderStyle != ExcelBorderStyle.None)
                {
                    ix2 = x2 - 4.87d;
                }
                if (DiagonalDown.BorderStyle != ExcelBorderStyle.None)
                {
                    ix1 = x1 + 4.87d;
                }

                ox1 = x1 - 0.7d;
                ox2 = x2 + 0.7d;
                oy1 = y1 + (PdfCellBorderData.Hair / 0.65d);
                oy2 = y2 + (PdfCellBorderData.Hair / 0.65d);
            }
            if (border.LineType == LineType.Bottom)
            {
                ix1 = x1 + 0.7d;
                ix2 = x2 - 0.7d;
                iy1 = y1 + (PdfCellBorderData.Hair / 0.65d);
                iy2 = y2 + (PdfCellBorderData.Hair / 0.65d);
                if (DiagonalUp.BorderStyle != ExcelBorderStyle.None)
                {
                    ix2 = x2 - 4.87d;
                }
                if (DiagonalDown.BorderStyle != ExcelBorderStyle.None)
                {
                    ix1 = x1 + 4.87d;
                }

                ox1 = x1 - 0.7d;
                ox2 = x2 + 0.7d;
                oy1 = y1 - (PdfCellBorderData.Hair / 0.65d);
                oy2 = y2 - (PdfCellBorderData.Hair / 0.65d);
            }
            else if (border.LineType == LineType.Left)
            {
                if (DiagonalUp.BorderStyle != ExcelBorderStyle.None)
                {
                    DiagonalUpFactor = 0.5d;
                }
                if (DiagonalDown.BorderStyle != ExcelBorderStyle.None)
                {
                    DiagonalDownFactor = 0.5d;
                }
                ix1 = x1 - (PdfCellBorderData.Hair / 0.65d);
                ix2 = x2 - (PdfCellBorderData.Hair / 0.65d);
                iy1 = y1 - 0.7d;
                iy2 = y2 + 0.7d;
                ox1 = x1 + (PdfCellBorderData.Hair / 0.65d);
                ox2 = x2 + (PdfCellBorderData.Hair / 0.65d);
                oy1 = y1 + 0.7d + DiagonalUpFactor;
                oy2 = y2 - 0.7d - DiagonalDownFactor;
            }
            else if (border.LineType == LineType.Right)
            {
                if (DiagonalUp.BorderStyle != ExcelBorderStyle.None)
                {
                    DiagonalUpFactor = 0.5d;
                }
                if (DiagonalDown.BorderStyle != ExcelBorderStyle.None)
                {
                    DiagonalDownFactor = 0.5d;
                }
                ix1 = x1 - (PdfCellBorderData.Hair / 0.65d);
                ix2 = x2 - (PdfCellBorderData.Hair / 0.65d);
                iy1 = y1 + 0.7d + DiagonalUpFactor;;
                iy2 = y2 - 0.7d - DiagonalDownFactor; ;
                ox1 = x1 + (PdfCellBorderData.Hair / 0.65d);
                ox2 = x2 + (PdfCellBorderData.Hair / 0.65d);
                oy1 = y1 - 0.7d;
                oy2 = y2 + 0.7d;
            }
            else if (border.LineType == LineType.DiagonalUp)
            {
                ix1 = x1 + 0.6d;
                ix2 = x2 - 4.87d;
                iy1 = y1 + 0.98d;
                iy2 = y2 - 0.765d;
                ox1 = x1 + 4.87d;
                ox2 = x2 - 0.6d;
                oy1 = y1 + 0.765d;
                oy2 = y2 - 0.98d;
            }
            else if (border.LineType == LineType.DiagonalDown)
            {
                ix1 = x1 + 0.6d;
                ix2 = x2 - 4.87d;
                iy1 = y1 - 0.98d;
                iy2 = y2 + 0.765d;

                ox1 = x1 + 4.87d;
                ox2 = x2 - 0.6d;
                oy1 = y1 - 0.765d;
                oy2 = y2 + 0.98d;

            }
            contentStream.AddCommand(border.BorderColor.ToStrokeCommand());
            contentStream.AddCommand($"{PdfCellBorderData.Hair.ToPdfString()} w");
            contentStream.AddCommand(border.BorderStyle != ExcelBorderStyle.Dotted ? (border.LineType == LineType.DiagonalUp || border.LineType == LineType.DiagonalDown ? "0 J" : "2 J") : "1 J");
            contentStream.AddCommand(PdfCellBorderData.NoDash);
            if ((border.LineType == LineType.DiagonalUp || border.LineType == LineType.DiagonalDown) && DiagonalUp.BorderStyle != ExcelBorderStyle.None && DiagonalDown.BorderStyle != ExcelBorderStyle.None)
            {

                double dx = ix2 - ix1;
                double dy = iy2 - iy1;
                double length = System.Math.Sqrt(dx * dx + dy * dy);

                double ux = dx / length;
                double uy = dy / length;

                double midX = (ix1 + ix2) / 2.0;
                double midY = (iy1 + iy2) / 2.0;

                double leftDist = 0.25;
                double rightDist = 2.15;

                double xA = midX - leftDist * ux;
                double yA = midY - leftDist * uy;
                double xB = midX + rightDist * ux;
                double yB = midY + rightDist * uy;

                contentStream.AddCommand($"{ix1.ToPdfStringF4()} {iy1.ToPdfStringF4()} m");
                contentStream.AddCommand($"{xA.ToPdfStringF4()} {yA.ToPdfStringF4()} l");
                contentStream.AddCommand($"{xB.ToPdfStringF4()} {yB.ToPdfStringF4()} m");
                contentStream.AddCommand($"{ix2.ToPdfStringF4()} {iy2.ToPdfStringF4()} l");


                dx = ox2 - ox1;
                dy = oy2 - oy1;
                length = System.Math.Sqrt(dx * dx + dy * dy);

                ux = dx / length;
                uy = dy / length;

                midX = (ox1 + ox2) / 2.0;
                midY = (oy1 + oy2) / 2.0;

                leftDist = 2.15;
                rightDist = 0.25;

                xA = midX - leftDist * ux;
                yA = midY - leftDist * uy;
                xB = midX + rightDist * ux;
                yB = midY + rightDist * uy;

                contentStream.AddCommand($"{ox1.ToPdfStringF4()} {oy1.ToPdfStringF4()} m");
                contentStream.AddCommand($"{xA.ToPdfStringF4()} {yA.ToPdfStringF4()} l");
                contentStream.AddCommand($"{xB.ToPdfStringF4()} {yB.ToPdfStringF4()} m");
                contentStream.AddCommand($"{ox2.ToPdfStringF4()} {oy2.ToPdfStringF4()} l");
            }
            else
            {
                contentStream.AddCommand($"{ix1.ToPdfStringF4()} {iy1.ToPdfStringF4()} m");
                contentStream.AddCommand($"{ix2.ToPdfStringF4()} {iy2.ToPdfStringF4()} l");
                contentStream.AddCommand($"{ox1.ToPdfStringF4()} {oy1.ToPdfStringF4()} m");
                contentStream.AddCommand($"{ox2.ToPdfStringF4()} {oy2.ToPdfStringF4()} l");
            }
            contentStream.AddCommand("S");
        }

        //This one could be made to look more fancy
        //private void DrawDoubleBorder(PdfContentStream contentStream, PdfCellBorderData borderData, LineType lineType, double x1, double y1, double x2, double y2)
        //{
        //    double offsetX = borderData.DoubleBorderOffsets.X;
        //    double offsetY = borderData.DoubleBorderOffsets.Y;
        //    contentStream.AddCommand(borderData.BorderColor.ToStrokeCommand());
        //    contentStream.AddCommand($"{Small.ToPdfString()} w");
        //    contentStream.AddCommand("[] 0 d");
        //    contentStream.AddCommand(lineType == LineType.DiagonalUp || lineType == LineType.DiagonalDown ? "0 J" : "2 J");
        //    if (lineType == LineType.DiagonalUp)
        //    {
        //        contentStream.AddCommand($"{(x1 + offsetX + offsetX).ToPdfString()} {(y1 + offsetY).ToPdfString()} m");
        //        contentStream.AddCommand($"{(x2 + -offsetX).ToPdfString()} {(y2 + -offsetY - offsetY).ToPdfString()} l");
        //    }
        //    else if (lineType == LineType.DiagonalDown)
        //    {
        //        contentStream.AddCommand($"{(x1 + offsetX + offsetX).ToPdfString()} {(y1 + -offsetY).ToPdfString()} m");
        //        contentStream.AddCommand($"{(x2 + -offsetX).ToPdfString()} {(y2 + offsetY + offsetY).ToPdfString()} l");
        //    }
        //    else
        //    {
        //        contentStream.AddCommand($"{x1.ToPdfString()} {y1.ToPdfString()} m");
        //        contentStream.AddCommand($"{x2.ToPdfString()} {y2.ToPdfString()} l");
        //    }
        //    contentStream.AddCommand("S");
        //    contentStream.AddCommand($"{Small.ToPdfString()} w");
        //    contentStream.AddCommand("[] 0 d");
        //    contentStream.AddCommand(lineType == LineType.DiagonalUp || lineType == LineType.DiagonalDown ? "0 J" : "2 J");
        //    if (lineType == LineType.Vertical)
        //    {
        //        contentStream.AddCommand($"{(x1 + offsetX).ToPdfString()} {(y1 + -offsetY).ToPdfString()} m");
        //        contentStream.AddCommand($"{(x2 + offsetX).ToPdfString()} {(y2 + offsetY).ToPdfString()} l");
        //    }
        //    else if (lineType == LineType.DiagonalUp)
        //    {
        //        contentStream.AddCommand($"{(x1 + offsetX).ToPdfString()} {(y1 + offsetY + offsetY).ToPdfString()} m");
        //        contentStream.AddCommand($"{(x2 + -offsetX - offsetX).ToPdfString()} {(y2 + -offsetY).ToPdfString()} l");
        //    }
        //    else if (lineType == LineType.DiagonalDown)
        //    {
        //        contentStream.AddCommand($"{(x1 + offsetX).ToPdfString()} {(y1 + -offsetY - offsetY).ToPdfString()} m");
        //        contentStream.AddCommand($"{(x2 + -offsetX - offsetX).ToPdfString()} {(y2 + offsetY).ToPdfString()} l");
        //    }
        //    else
        //    {
        //        contentStream.AddCommand($"{(x1 + offsetX).ToPdfString()} {(y1 + offsetY).ToPdfString()} m");
        //        contentStream.AddCommand($"{(x2 + -offsetX).ToPdfString()} {(y2 + offsetY).ToPdfString()} l");
        //    }
        //    contentStream.AddCommand("S");
        //}

        private void DrawSlantDashDotBorder(PdfContentStream contentStream, PdfCellBorderData border, double x1, double y1, double x2, double y2)
        {
            contentStream.AddCommand(border.BorderColor.ToStrokeCommand());
            contentStream.AddCommand("Q");
            contentStream.AddCommand("q");
            contentStream.AddCommand($"{PdfCellBorderData.Small.ToPdfString()} w");
            contentStream.AddCommand(border.LineType == LineType.DiagonalUp || border.LineType == LineType.DiagonalDown ? "0 J" : "2 J");
            contentStream.AddCommand(PdfCellBorderData.DashDot);
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


        //public void RenderBorder(PdfContentStream contentStream, PdfCellBorderData borderData, LineType lineType, double x1, double y1, double x2, double y2)
        //{
        //    switch (borderData.BorderStyle)
        //    {
        //        case ExcelBorderStyle.None:
        //            return;
        //        case ExcelBorderStyle.Hair:
        //            DrawBasicBorder(contentStream, borderData, lineType, Hair, "[] 0 d");
        //            break;
        //        case ExcelBorderStyle.Dotted:
        //            DrawBasicBorder(contentStream, borderData, lineType, Small, "[0 2] 0 d");
        //            break;
        //        case ExcelBorderStyle.DashDot:
        //            DrawBasicBorder(contentStream, borderData, lineType, Small, "[4 2 1 2] 0 d");
        //            break;
        //        case ExcelBorderStyle.Thin:
        //            DrawBasicBorder(contentStream, borderData, lineType, Thin, "[] 0 d");
        //            break;
        //        case ExcelBorderStyle.DashDotDot:
        //            DrawBasicBorder(contentStream, borderData, lineType, Small, "[4 2 1 2 1 2] 0 d");
        //            break;
        //        case ExcelBorderStyle.Dashed:
        //            DrawBasicBorder(contentStream, borderData, lineType, Small, "[4 3] 0 d");
        //            break;
        //        case ExcelBorderStyle.MediumDashDotDot:
        //            DrawBasicBorder(contentStream, borderData, lineType, Medium, "[6 3 2 3 2 3] 0 d");
        //            break;
        //        case ExcelBorderStyle.MediumDashed:
        //            DrawBasicBorder(contentStream, borderData, lineType, Medium, "[6 4] 0 d");
        //            break;
        //        case ExcelBorderStyle.MediumDashDot:
        //            DrawBasicBorder(contentStream, borderData, lineType, Medium, "[6 3 2 3] 0 d");
        //            break;
        //        case ExcelBorderStyle.Thick:
        //            DrawBasicBorder(contentStream, borderData, lineType, Thick, "[] 0 d");
        //            break;
        //        case ExcelBorderStyle.Medium:
        //            DrawBasicBorder(contentStream, borderData, lineType, Medium, "[] 0 d");
        //            break;
        //        case ExcelBorderStyle.SlantDashDot:
        //            DrawSlantDashDotBorder(contentStream, borderData, lineType, x1, y1, x2, y2);
        //            return;
        //        case ExcelBorderStyle.Double:
        //            //DrawDoubleBorder(contentStream, borderData, lineType, x1, y1, x2, y2);
        //            return;
        //    }
        //    contentStream.AddCommand($"{x1.ToPdfStringF4()} {y1.ToPdfStringF4()} m");
        //    contentStream.AddCommand($"{x2.ToPdfStringF4()} {y2.ToPdfStringF4()} l");
        //    contentStream.AddCommand("S");
        //}
    }
}

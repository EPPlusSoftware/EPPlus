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
using EPPlus.Export.Pdf.Enums;
using Vector2 = EPPlus.Graphics.Geometry.Vector2;

namespace EPPlus.Export.Pdf.Layout
{
    internal static class PdfCellAlignment
    {
        internal static double CalculateVerticalAlignment(PdfCellAlignmentData align, double textHeight, double firstAscent, double lastAscent, double y, double height, double padding)
        {
            double newY = y;
            switch (align.VerticalAlignment)
            {
                case ExcelVerticalAlignment.Top:
                case ExcelVerticalAlignment.Distributed:
                case ExcelVerticalAlignment.Justify:
                    newY = (y + height) - padding - firstAscent;
                    break;
                case ExcelVerticalAlignment.Center:
                    newY = y + (height + textHeight - firstAscent - lastAscent) / 2d;
                    break;
                case ExcelVerticalAlignment.Bottom:
                    newY = y + padding + textHeight - lastAscent;
                    break;
            }
            return newY;
        }

        internal static double CalculateHorizontalAlignment(PdfCellAlignmentData align, string text, double textLength, double x, double width, double padding, double leftInset = 0d)
        {
            x += leftInset;
            width -= leftInset;
            double newX = x;
            switch (align.HorizontalAlignment)
            {
                case ExcelHorizontalAlignment.Left:
                    newX = x + padding;
                    break;
                case ExcelHorizontalAlignment.Center:
                case ExcelHorizontalAlignment.CenterContinuous:
                    newX = x + (width - textLength) / 2d;
                    break;
                case ExcelHorizontalAlignment.Right:
                    newX = x + (width - textLength) - padding;
                    break;
                case ExcelHorizontalAlignment.Distributed:
                case ExcelHorizontalAlignment.Justify:
                case ExcelHorizontalAlignment.Fill:
                case ExcelHorizontalAlignment.General:
                default:
                    // "General": numbers right-align, everything else left-aligns.
                    if (double.TryParse(text, out double value))
                    {
                        newX = x + (width - textLength) - padding;
                    }
                    else
                    {
                        newX = x + padding;
                    }
                    break;
            }
            return newX;
        }

        internal static Vector2 CalculatePositionFromRotation(PdfCellAlignmentData align, double textLength, double x, double y)
        {
            double newX = x;
            double newY = y;
            if (align.TextRotation < 0)
            {
                double rot = align.TextRotation * System.Math.PI / 180.0;
                newX += textLength * (1 - System.Math.Cos(rot));
                newY -= textLength * System.Math.Sin(rot);
            }
            else if (align.TextRotation > 0)
            {
                double rot = align.TextRotation * System.Math.PI / 180.0;
                newX += textLength * (1 - System.Math.Cos(rot));
            }
            return new Vector2(newX, newY);
        }

        internal static Vector2 CalculateAlignment(PdfCellAlignmentData align, string text, double textLength, double textHeight, double firstLineAscent, double lastLineAscent, double x, double y, double width, double height, double rightMargin, double bottomMargin)
        {
            double leftInset = align.IconSize > 0d ? align.IconSize + 3d : 0d;
            if (align.TextRotation != 0 && !align.IsVertical)
            {
                return CalculateRotatedAlignment(align, textLength, textHeight, firstLineAscent, x, y, width, height, rightMargin, bottomMargin, leftInset);
            }
            double newX = CalculateHorizontalAlignment(align, text, textLength, x, width, rightMargin, leftInset);
            double newY = CalculateVerticalAlignment(align, textHeight, firstLineAscent, lastLineAscent, y, height, 0d);
            return CalculatePositionFromRotation(align, textLength, newX, newY);
        }

        internal static Vector2 CalculateRotatedAlignment(PdfCellAlignmentData align, double textLength, double textHeight, double firstLineAscent, double x, double y, double width, double height, double rightMargin, double bottomMargin, double leftInset = 0d)
        {
            x += leftInset;
            width -= leftInset;
            double ascent = firstLineAscent;
            double descent = textHeight - firstLineAscent;
            if (descent < 0d)
            {
                descent = 0d;
            }
            double theta = align.TextRotation * System.Math.PI / 180.0;
            double cos = System.Math.Cos(theta);
            double sin = System.Math.Sin(theta);
            double[] gx = { 0d, textLength, 0d, textLength };
            double[] gy = { ascent, ascent, -descent, -descent };
            double minX = double.MaxValue, maxX = double.MinValue, minY = double.MaxValue, maxY = double.MinValue;
            for (int i = 0; i < 4; i++)
            {
                double ux = cos * gx[i] - sin * gy[i];
                double uy = sin * gx[i] + cos * gy[i];
                if (ux < minX) minX = ux;
                if (ux > maxX) maxX = ux;
                if (uy < minY) minY = uy;
                if (uy > maxY) maxY = uy;
            }
            double blockWidth = maxX - minX;
            double blockHeight = maxY - minY;
            double bx;
            switch (align.HorizontalAlignment)
            {
                case ExcelHorizontalAlignment.Left:
                    bx = x + rightMargin; break;
                case ExcelHorizontalAlignment.Right:
                    bx = x + width - blockWidth - rightMargin; break;
                default: // Center / General
                    bx = x + (width - blockWidth) / 2d; break;
            }
            double by;
            switch (align.VerticalAlignment)
            {
                case ExcelVerticalAlignment.Top:
                    by = y + height - blockHeight - bottomMargin; break;
                case ExcelVerticalAlignment.Bottom:
                    by = y + bottomMargin; break;
                default: // Center / Justify / Distributed
                    by = y + (height - blockHeight) / 2d; break;
            }
            return new Vector2(bx - minX, by - minY);
        }
    }
}

using OfficeOpenXml.Drawing;
using System;

namespace OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.Utils
{
    internal static class TextUtils
    {
        internal static double PointToPixel(this double pointSize)
        {
            //1 inch is 72 pts. "Inches * dots/inch = dots" aka Pixels
            return pointSize / 72 * 96;
        }

        internal static float PointToPixel(this float pointSize)
        {
            //1 inch is 72 pts. "Inches * dots/inch = dots" aka Pixels
            return pointSize / 72 * 96;
        }

        internal static double PixelToPoint(this double pixelSize)
        {
            var roundedSize = (double)RoundToWhole(pixelSize);
            return roundedSize / 96 * 72;
        }

        internal static double EmuToPoint(this double emuNumber)
        {
           return emuNumber / ExcelDrawing.EMU_PER_POINT;
        }

        
        internal static double EmuToPixels(this double? emuNumber) => emuNumber.HasValue ? emuNumber.Value.EmuToPixels() : 0D;

        internal static double EmuToPixels(this double emuNumber)
        {
            return emuNumber / ExcelDrawing.EMU_PER_PIXEL;
        }

        internal static int RoundToWhole(double number)
        {
            return (int)Math.Round(number, 0, MidpointRounding.AwayFromZero);
        }
    }
}

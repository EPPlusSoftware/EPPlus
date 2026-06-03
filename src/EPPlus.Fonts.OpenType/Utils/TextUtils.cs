using OfficeOpenXml.Drawing;
using System;
using System.Globalization;

namespace EPPlus.Fonts.OpenType.Utils
{
    public static class TextUtils
    {
        public const int EMU_PER_POINT = 12700;
        public const int EMU_PER_PIXEL = 9525;

        public static double PointToPixel(this double pointSize)
        {
            //1 inch is 72 pts. "Inches * dots/inch = dots" aka Pixels
            return pointSize / 72 * 96;
        }
        public static string PointToPixelString(this double pointSize, string format=null)
        {
            //1 inch is 72 pts. "Inches * dots/inch = dots" aka Pixels
            var pixelSize = pointSize / 72 * 96;
            if (format != null)
            {
                return pixelSize.ToString(format, CultureInfo.InvariantCulture);
            }
            return pixelSize.ToString(CultureInfo.InvariantCulture);
        }

        public static double PointToPixel(this double pointSize, bool isFonts)
        {
            return PointToPixel(pointSize);
        }

        public static float PointToPixel(this float pointSize)
        {
            //1 inch is 72 pts. "Inches * dots/inch = dots" aka Pixels
            return pointSize / 72 * 96;
        }

        public static double PixelToPoint(this double pixelSize)
        {
            //var roundedSize = (double)RoundToWhole(pixelSize);
            return pixelSize / 96 * 72;
        }

        public static double EmuToPoint(this double emuNumber)
        {
           return emuNumber / EMU_PER_POINT;
        }


        public static double EmuToPixels(this double? emuNumber) => emuNumber.HasValue ? emuNumber.Value.EmuToPixels() : 0D;

        public static double EmuToPixels(this double emuNumber)
        {
            return emuNumber / EMU_PER_PIXEL;
        }

        public static int RoundToWhole(double number)
        {
            return (int)Math.Round(number, 0, MidpointRounding.AwayFromZero);
        }
    }
}

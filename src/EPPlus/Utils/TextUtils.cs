using OfficeOpenXml.Drawing;
using System;

namespace OfficeOpenXml.Utils
{
    internal static class TextUtils
    {
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
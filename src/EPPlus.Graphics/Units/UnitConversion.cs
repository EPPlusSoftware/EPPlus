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
namespace EPPlus.Graphics.Units
{
    public static class UnitConversion
    {
        public const double PointsPerInch = 72.0d;
        public const double MmPerInch = 25.4d;
        public const double DPI = 600d;

        public static double MmToPoints(double mm)
        {
            return mm * PointsPerInch / MmPerInch;
        }

        public static double PointsToMm(double points)
        {
            return points * MmPerInch / PointsPerInch;
        }

        public static int MmToPointsRounded(double mm)
        {
            return (int)System.Math.Round(MmToPoints(mm));
        }

        public static int PointsToMmRounded(double points)
        {
            return (int)System.Math.Round(PointsToMm(points));
        }

        public static double ExcelPointsToMM(double excelPoints)
        {
            return excelPoints * (1 / PointsPerInch*10);
        }

        public static double ExcelColumnWidthToPoints(double columnWidth, double char0)
        {
            return columnWidth * char0 + 0.75d;
        }

        public static double ExcelRowHeightToPoints(double rowHeight)
        {
            return rowHeight; //These values are guessed.
        }

        public static double ToMillimeters(double inches)
        {
            return inches * 25.4;
        }
    }
}

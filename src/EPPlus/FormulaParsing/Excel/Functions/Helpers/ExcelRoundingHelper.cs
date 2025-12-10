/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/09/2025         EPPlus Software AB       Refactoring of RoundingHelper to support Excel's legacy rounding functions
 *************************************************************************************************/

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers
{

    using System;

    internal static class ExcelRoundingHelper
    {
        // CEILING.PRECISE: always toward +∞, uses |significance|; 0 if number==0 or significance==0.
        // Ref: Microsoft Support
        public static double CeilingPrecise(double number, double significance)
        {
            if (number == 0d || significance == 0d) return 0d;

            double m = Math.Abs(significance);
            double q = number / m;
            double qRounded = Math.Ceiling(q);           // toward +∞
            return Math.Round(qRounded * m, 14);
        }

        // CEILING.MATH: positives toward +∞; negatives: mode==0 => toward 0; mode!=0 => away from 0.
        // 'mode' may be null (omitted); treat as 0 (toward zero for negatives).
        // Ref: Microsoft Support
        public static double CeilingMath(double number, double significance, double? mode)
        {
            double m = (significance == 0d) ? 1d : Math.Abs(significance);
            double q = number / m;
            double qRounded;

            if (number >= 0d)
            {
                qRounded = Math.Ceiling(q);            // toward +∞
            }
            else
            {
                double md = (mode.HasValue ? mode.Value : 0d);
                if (md != 0d)
                    qRounded = Math.Floor(q);          // away from 0 (more negative)
                else
                    qRounded = Math.Ceiling(q);        // toward 0 (less negative)
            }

            return Math.Round(qRounded * m, 14);
        }

        // FLOOR.PRECISE: always down to nearest |significance|; 0 if number==0 or significance==0.
        // Ref: Microsoft Support
        public static double FloorPrecise(double number, double significance)
        {
            if (number == 0d || significance == 0d) return 0d;

            double m = Math.Abs(significance);
            double q = number / m;
            double qRounded = Math.Floor(q);            // toward −∞
            return Math.Round(qRounded * m, 14);
        }


        // FLOOR.MATH: positives toward −∞; negatives default away from 0 (more negative).
        // If 'mode' is specified and > 0, negatives round toward 0 (less negative).
        // If 'mode' is omitted (null), keep default (away from 0).
        // Ref: Microsoft Support + Exceljet
        public static double FloorMath(double number, double significance, double? mode)
        {
            double m = (significance == 0d) ? 1d : Math.Abs(significance);
            double q = number / m;
            double qRounded;

            if (number >= 0d)
            {
                // Positive: floor (toward −∞)
                qRounded = Math.Floor(q);
            }
            else
            {
                if (!mode.HasValue)
                {
                    // Omitted mode: away from zero -> more negative
                    qRounded = Math.Floor(q);
                }
                else
                {
                    // mode > 0 => toward zero; mode <= 0 => away from zero (Excel doc is inconsistent on 0,
                    // but test suite targets: mode=1 => toward zero)
                    double md = mode.Value;
                    qRounded = (md > 0d) ? Math.Ceiling(q) : Math.Floor(q);
                }
            }

            return Math.Round(qRounded * m, 14);
        }

    }

}

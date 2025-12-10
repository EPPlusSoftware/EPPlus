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
using System;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers
{

    internal static class ExcelLegacyRounding
    {
        // Implements legacy Excel FLOOR(number, significance)
        // Behaviour:
        //  - Positive numbers: round down toward 0
        //  - Negative numbers: round down away from 0 (i.e., more negative)
        //  - Mixed signs: if number>0 and significance<0 -> #NUM! (caller should map to Excel error)
        //  - If significance == 0: Excel's behaviour varies by function; returning 0 is a pragmatic choice.
        // References:
        //   https://support.microsoft.com/en-us/office/floor-function-14bb497c-24f2-4e04-b327-b0b4de5a8886
        //   https://support.microsoft.com/en-us/office/floor-function-5168a039-9501-4e02-ac52-83914290ac55
        public static bool TryFloor(double number, double significance, out double result)
        {
            result = 0d;

            // Excel legacy rule: olika tecken → #NUM!
            if ((number > 0 && significance < 0) || (number < 0 && significance > 0))
                return false;

            if (significance == 0d)
            {
                result = 0d;
                return true;
            }

            if (number == 0d)
            {
                result = 0d;
                return true;
            }

            // VIKTIGT: Behåll tecknen på significance!
            // Använd Math.Floor direkt på kvoten – det rundar alltid mot -∞
            double quotient = number / significance;
            double roundedQuotient = Math.Floor(quotient);

            result = roundedQuotient * significance;

            // Undvik små flytalsfel (t.ex. -26.999999999 → -27 istället för -26.999999999)
            // Excel gör detta internt med hög precision
            result = System.Math.Round(result, 14);

            return true;
        }
    }

}

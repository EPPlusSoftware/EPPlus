/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    internal static class RoundingHelper
    {
        public enum Direction
        {
            Up,
            Down,
            AlwaysDown
        }
        public static double Round(double number, double multiple, Direction direction)
        {
            if (multiple == 0) return 0d;
            if(direction == Direction.Up)
            {
                if (multiple < 1 && multiple > 0)
                {
                    var floor = System.Math.Floor(number);
                    var rest = number - floor;
                    var nSign = (int)(rest / multiple) + 1;
                    return floor + (nSign * multiple);
                }
                else if (multiple == 1)
                {
                    return System.Math.Ceiling(number);
                }
                else if (number % multiple == 0)
                {
                    return number;
                }
                else
                {
                    return System.Math.Round(number - (number % multiple) + multiple, 14);
                }
            }
            else
            {
                if (multiple < 1 && multiple > 0)
                {
                    var floor = System.Math.Floor(number);
                    var rest = number - floor;
                    var nSign = (int)(rest / multiple);
                    if (direction == Direction.AlwaysDown && number < 0) nSign += 1;
                    return floor + (nSign * multiple);
                }
                else if (multiple == 1)
                {
                    return System.Math.Floor(number);
                }
                else if (number % multiple == 0)
                {
                    return number;
                }
                else
                {
                    if (direction == Direction.AlwaysDown && number < 0)
                    {
                        if (multiple < 0) multiple *= -1;
                        return System.Math.Round(number - (number % multiple) - multiple, 14);
                    }
                   
                    return System.Math.Round(number - (number % multiple), 14);
                }
            }
        }

        internal static void ValidateNumberAndSign(double number, double sign)
        {
            if (number > 0d && sign < 0)
            {
                var values = string.Format("num: {0}, sign: {1}", number, sign);
                throw new InvalidOperationException("Ceiling cannot handle a negative significance when the number is positive" + values);
            }
        }
    }
}

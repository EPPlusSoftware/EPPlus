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
            AlwaysDown,
            AlwaysUp,
            Nearest
        }

        public static double Round(double number, double multiple, Direction direction)
        {
            if (multiple == 0) return 0d;
            var isNegativeNumber = number < 0;
            var isNegativeMultiple = multiple < 0;
            var n = isNegativeNumber ? number * -1 : number;
            var m = isNegativeMultiple ? multiple * -1 : multiple;
            if (number % multiple == 0)
            {
                return number;
            }
            else if (multiple > -1 && multiple < 1)
            { 
                
                var floor = System.Math.Floor(n);
                var rest = n - floor;
                var nSign = (int)(rest / m) + 1;

                var upperRound = System.Math.Round(nSign * m, 14);
                var lowerRound = System.Math.Round((nSign - 1) * m, 14);
                return ExecuteRounding(n, floor + lowerRound, floor + upperRound, direction, isNegativeNumber);
            }
            var result = double.NaN;
            if (m == 1)
            {
                if (direction == Direction.Up || direction == Direction.AlwaysUp)
                {
                    if (direction == Direction.AlwaysUp && isNegativeNumber)
                        result = System.Math.Floor(n);
                    else
                        result = System.Math.Ceiling(n);
                }
                else if(direction == Direction.Nearest)
                {
                    result = System.Math.Floor(n);
                    if (n % 1 >= 0.5)
                    {
                        result++;
                    }
                }   
                else
                {
                    if (direction == Direction.AlwaysDown && isNegativeNumber)
                        result = System.Math.Ceiling(n);
                    else
                        result = System.Math.Floor(n);
                }   
            }
            else if(m > n)
            {
                if(m >= (n/2d))
                {
                    result = m;
                }
                else
                {
                    result = 0;
                }
            }
            else if(direction == Direction.Up || direction == Direction.AlwaysUp)
            {
                if (direction == Direction.AlwaysUp && number < 0)
                {
                    if (multiple < 0) multiple *= -1;
                    return System.Math.Round(number - (number % multiple), 14);
                }
                return System.Math.Round(number - (number % multiple) + multiple, 14);
            }
            else if(direction == Direction.Nearest)
            {
                if ((n % m >= (m / 2d)))
                    result = System.Math.Round(n + (m - n % m));
                else
                    result = System.Math.Round(n - (n % m));
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
            return isNegativeNumber ? -1 * result : result;
        }

        public static double ExecuteRounding(double number, double lowerRound, double upperRound, Direction direction, bool isNegativeNumber)
        {
            var result = double.NaN;
            if(direction == Direction.Nearest)
            {
                if ((upperRound - number) > (number - lowerRound))
                    result = lowerRound;
                else
                    result = upperRound;
            }
            else if(direction == Direction.AlwaysUp)
            {
                result = isNegativeNumber ? lowerRound : upperRound;
            }
            else if(direction == Direction.Up)
            {
                result = upperRound;
            }
            else if(direction == Direction.AlwaysDown)
            {
                result = isNegativeNumber ? upperRound : lowerRound;
            }
            else
            {
                result = lowerRound;
            }
            return isNegativeNumber ? -1 * result : result;
        }
        /*
        public static double Round(double number, double multiple, Direction direction)
        {
            if (multiple == 0) return 0d;
            var isNegativeNumber = number < 0;
            var isNegativeMultiple = multiple < 0;
            if(direction == Direction.Up || direction == Direction.AlwaysUp)
            {
                if ((multiple < 1 && multiple > 0) || (multiple > -1 && multiple < 0))
                {
                    if (isNegativeMultiple) multiple *= -1;
                    if (isNegativeNumber) number *= -1;
                    var floor = System.Math.Floor(number);
                    var rest = number - floor;
                    var nSign = (int)(rest / multiple) + 1;
                    var result = floor + (nSign * multiple);
                    if(isNegativeNumber)
                    {
                        if (direction == Direction.AlwaysUp) result -= multiple;
                        return result * -1;
                    }
                    return result;
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
                    if (direction == Direction.AlwaysUp && number < 0)
                    {
                        if (multiple < 0) multiple *= -1;
                        return System.Math.Round(number - (number % multiple), 14);
                    }
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
         */

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

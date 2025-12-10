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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
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
            if (multiple == 0d) return 0d;

            var isNegativeNumber = number < 0;
            var isNegativeMultiple = multiple < 0;

            var n = isNegativeNumber ? -number : number;
            var m = isNegativeMultiple ? -multiple : multiple;

            if (number % multiple == 0d)
                return number;

            else if (multiple > -1 && multiple < 1)
            {
                var floor = System.Math.Floor(n);
                var rest = n - floor;
                var nSign = (int)(rest / m) + 1;

                var upperRound = System.Math.Round(nSign * m, 14);
                var lowerRound = System.Math.Round((nSign - 1) * m, 14);

                var positiveResult = ExecuteRounding(n, floor + lowerRound, floor + upperRound, direction, isNegativeNumber, isNegativeMultiple);
                return positiveResult * (isNegativeNumber ? -1 : 1);
            }

            double result;

            if (m == 1)
            {
                if (direction == Direction.Up || direction == Direction.AlwaysUp)
                {
                    result = (direction == Direction.AlwaysUp && isNegativeNumber) ? Math.Floor(n) : Math.Ceiling(n);
                    if(isNegativeNumber && isNegativeMultiple)
                    {
                        result++;
                    }
                }  
                else if (direction == Direction.Nearest)
                {
                    result = Math.Floor(n);
                    if (n % 1 >= 0.5) result++;
                }
                else
                {
                    result = (direction == Direction.AlwaysDown && isNegativeNumber) ? Math.Ceiling(n) : Math.Floor(n);
                    if (isNegativeNumber && isNegativeMultiple)
                    {
                        result--;
                    }
                }   
            }
            else if (m > n)
            {
                var positiveResult = ExecuteRounding(n, 0, m, direction, isNegativeNumber, isNegativeMultiple);
                return positiveResult * (isNegativeNumber ? -1 : 1);
            }
            else if (direction == Direction.Up || direction == Direction.AlwaysUp)
            {
                var mod = n % m;
                mod = RoundToSignificantFig(mod, 15);
                if (mod == 0) return number;

                result = n - mod + m;
                if(isNegativeNumber && !isNegativeMultiple)
                {
                    result -= m;
                }

                if (direction == Direction.AlwaysUp && isNegativeNumber)
                    return -result;

                return result;
            }
            else if (direction == Direction.Nearest)
            {
                var mod = n % m;
                result = mod >= m / 2d ? n + (m - mod) : n - mod;
            }
            else // Down / AlwaysDown
            {
                var mod = n % m;
                result = n - mod;
            }

            return result * (isNegativeNumber ? -1 : 1);
        }

        public static double ExecuteRounding(double number, double lowerRound, double upperRound, Direction direction, bool isNegativeNumber, bool isNegativeMultiple)
        {
            var result = double.NaN;
            if (direction == Direction.Nearest)
            {
                if ((upperRound - number) > (number - lowerRound))
                    result = lowerRound;
                else
                    result = upperRound;
            }
            else if (direction == Direction.AlwaysUp)
            {
                if(isNegativeMultiple)
                {
                    result = isNegativeNumber ? upperRound : lowerRound;
                }
                else
                {
                    result = isNegativeNumber ? lowerRound : upperRound;
                }
                
            }
            else if (direction == Direction.Up)
            {
                if (isNegativeMultiple)
                {
                    result = lowerRound;
                }
                else
                {
                    result = upperRound;
                }

            }
            else if (direction == Direction.AlwaysDown)
            {
                if (isNegativeMultiple)
                {
                    result = isNegativeNumber ? lowerRound : upperRound;
                }
                else
                {
                    result = isNegativeNumber ? upperRound : lowerRound;
                }
            }
            else
            {
                result = lowerRound;
            }
            return result;
        }


        internal static bool IsInvalidNumberAndSign(double number, double sign)
        {
            return (number > 0d && sign < 0);
        }

        internal static double RoundToSignificantFig(double number, double nSignificantFigures)
        {
            return GetSignificantFigures(number, (int)nSignificantFigures);//RoundToSignificantFig(number, nSignificantFigures, true);
        }

        internal static double RoundToSignificantFig(double number, double nSignificantFigures, bool awayFromMidpoint)
        {
            var isNegative = false;
            if(number < 0d)
            {
                number *= -1;
                isNegative = true;
            }
            var nFiguresIntPart = GetNumberOfDigitsIntPart(number);
            var nLeadingZeroDecimals = GetNumberOfLeadingZeroDecimals(number);
            var nFiguresDecimalPart = nSignificantFigures - nFiguresIntPart - nLeadingZeroDecimals;
            if (number < 1d)
            {
                nFiguresDecimalPart -= nLeadingZeroDecimals;
            }
            var tmp = number * System.Math.Pow(10, nFiguresDecimalPart + nLeadingZeroDecimals);
            var e = awayFromMidpoint? tmp + 0.5 : tmp;
            if(awayFromMidpoint)
            { 
                if ((float)e == (float)System.Math.Ceiling(tmp))
                {
                    var f = System.Math.Ceiling(tmp);
                    var h = (int)f - 2;
                    if (h % 2 != 0)
                    {
                        e = e - 1;
                    }
                }
            }
            var intVersion = System.Math.Floor(e);
            double divideBy = System.Math.Pow(10, nFiguresDecimalPart + nLeadingZeroDecimals);
            var result = intVersion / divideBy;
            return isNegative ? result * -1 : result;
        }

        internal static double GetSignificantFigures(double number, int numberOfSignificantFigures)
        {
            double wholeNumberPart = Math.Floor(Math.Log10(Math.Abs(number)));
            double adjust = Math.Pow(10, wholeNumberPart);
            if (number == 0.0) return 0.0;
            else
            {
                try
                {
                    double product = adjust * Math.Round(number / adjust, numberOfSignificantFigures, MidpointRounding.AwayFromZero);
                    if ((int)wholeNumberPart >= numberOfSignificantFigures)
                    {
                        return Math.Round(product, 0, MidpointRounding.AwayFromZero);
                    }
                    return (double)Decimal.Round((Decimal)product, Math.Min(numberOfSignificantFigures - (int)wholeNumberPart, 28), MidpointRounding.AwayFromZero);
                }
                catch 
                {
                    return number;
                }
            }
        }

        /// <summary>
        /// Count the number of digits left of the decimal point
        /// </summary>
        /// <param name="n"></param>
        /// <returns></returns>
        internal static double GetNumberOfDigitsIntPart(double n)
        {
            var tmp = n;
            int nFiguresIntPart;
            for (nFiguresIntPart = 0; tmp >= 1; ++nFiguresIntPart)
                tmp = tmp / 10;
            return nFiguresIntPart;
        }

        private static double GetNumberOfLeadingZeroDecimals(double n)
        {
            if (n == 0) return 0;
            var tmp = n;
            var result = 0;
            while (tmp < 1d)
            {
                tmp *= 10;
                result++;
            }
            return result - 1;
        }
    }
}

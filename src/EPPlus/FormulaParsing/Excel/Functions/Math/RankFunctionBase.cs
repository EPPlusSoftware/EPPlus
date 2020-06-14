/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/25/2020         EPPlus Software AB       Implemented function
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    internal abstract class RankFunctionBase : ExcelFunction
    {
        protected static List<double> GetNumbersFromRange(FunctionArgument refArg, bool sortAscending)
        {
            var numbers = new List<double>();
            foreach (var cell in refArg.ValueAsRangeInfo)
            {
                var cellValue = Utils.ConvertUtil.GetValueDouble(cell.Value, false, true);
                if (!double.IsNaN(cellValue))
                {
                    numbers.Add(cellValue);
                }
            }
            if (sortAscending)
                numbers.Sort();
            else
                numbers.Sort((x, y) => y.CompareTo(x));
            return numbers;
        }

        protected double[] GetNumbersFromArgs(IEnumerable<FunctionArgument> arguments, int index, ParsingContext context)
        {
            var array = ArgsToDoubleEnumerable(new FunctionArgument[] { arguments.ElementAt(index) }, context)
                .Select(x => (double)x)
                .OrderBy(x => x)
                .ToArray();
            return array;
        }

        protected double PercentRankIncImpl(double[] array, double number)
        {
            var smallerThan = 0d;
            var largestBelow = 0d;
            var ix = 0;
            while (number > array[ix])
            {
                smallerThan++;
                largestBelow = array[ix];
                ix++;
            }
            var fullMatch = AreEqual(number, array[ix]);
            while (ix < array.Length - 1 && AreEqual(number, array[ix]))
                ix++;
            var smallestAbove = array[ix];
            var largerThan = AreEqual(number, array[array.Length - 1]) ? 0 : array.Length - ix;
            if (fullMatch)
                return smallerThan / (smallerThan + largerThan);
            var percentrankLow = PercentRankIncImpl(array, largestBelow);
            var percentrankHigh = PercentRankIncImpl(array, smallestAbove);
            return percentrankLow + (percentrankHigh - percentrankLow) * ((number - largestBelow) / (smallestAbove - largestBelow));
        }

        /// <summary>
        /// Rank functions rounds towards zero, i.e. 0.41666666 should be rounded to 0.4166 if 4 decimals.
        /// </summary>
        /// <param name="number">The number to round</param>
        /// <param name="decimals">Number of decimals</param>
        /// <returns></returns>
        protected double RoundResult(double number, int decimals)
        {
            if(System.Math.Round(number, decimals) - number > double.Epsilon)
                return System.Math.Round(number, decimals, MidpointRounding.AwayFromZero) - System.Math.Pow(10, decimals * -1);
            return number;
        }
    }
}

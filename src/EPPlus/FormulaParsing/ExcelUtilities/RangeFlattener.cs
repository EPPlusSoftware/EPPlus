/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  21/06/2023         EPPlus Software AB       Initial release EPPlus 7
 *************************************************************************************************/
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    internal class RangeFlattener
    {
        /// <summary>
        /// returns a list of nullable doubles based on the supplied range.
        /// both dates and numeric values will be included.
        /// </summary>
        /// <param name="r1"></param>
        /// <param name="addNullifEmpty"></param>
        /// <returns></returns>
        public static List<double?> FlattenRange(IRangeInfo r1, bool addNullifEmpty=true)
        {
            var result = new List<double?>();

            for (var row = 0; row < r1.Size.NumberOfRows; row++)
            {
                for (var column = 0; column < r1.Size.NumberOfCols; column++)
                {
                    var val = r1.GetOffset(row, column);

                    if (ConvertUtil.IsNumericOrDate(val))
                    {
                        var yNum = ConvertUtil.GetValueDouble(val);

                        result.Add(yNum);
                    }
                    else if(addNullifEmpty)
                    {
                        result.Add(null);
                    }
                }
            }
            return result;
        }
        public static List<object> FlattenRangeObject(IRangeInfo r1)
        {
            var result = new List<object>();

            for (var row = 0; row < r1.Size.NumberOfRows; row++)
            {
                for (var column = 0; column < r1.Size.NumberOfCols; column++)
                {
                    result.Add(r1.GetOffset(row, column));
                }
            }
            return result;
        }
        /// <summary>
        /// produces two lists based on the supplied ranges. The lists will contain all data from positions where both ranges has numeric values. 
        /// </summary>
        /// <param name="r1">range 1</param>
        /// <param name="r2">range 2</param>
        /// <param name="l1">a list containing all numeric values from <paramref name="r1"/> that has a corresponding value in <paramref name="r2"/></param>
        /// <param name="l2">a list containing all numeric values from <paramref name="r2"/> that has a corresponding value in <paramref name="r1"/></param>
        public static void GetNumericPairLists(IRangeInfo r1  , IRangeInfo r2, bool dataPointsEqual,  out List<double> l1, out List<double> l2)
        {
            if (dataPointsEqual)
            {
                if (r1.GetNCells() != r2.GetNCells())
                {
                    throw new ArgumentException("Ranges r1 and r2 must have the same number of cells");
                }

                var rangeValues1 = FlattenRange(r1);
                var rangeValues2 = FlattenRange(r2);
                l1 = new List<double>();
                l2 = new List<double>();
                for (var i = 0; i < rangeValues1.Count; i++)
                {
                    if ( rangeValues1[i].HasValue && rangeValues2[i].HasValue)
                    {
                        l1.Add(rangeValues1[i].Value);
                        l2.Add(rangeValues2[i].Value);
                    }
                }

            }
            else
            {
                var rangeValues1 = FlattenRange(r1);
                var rangeValues2 = FlattenRange(r2);
                l1 = new List<double>();
                l2 = new List<double>();

                for (var i = 0; i < rangeValues1.Count; i++)
                {
                    if (rangeValues1[i].HasValue)
                    {
                        l1.Add(rangeValues1[i].Value);
                    }
                }

                for (var i = 0; i < rangeValues2.Count; i++)
                {
                    if (rangeValues2[i].HasValue)
                    {
                        l2.Add(rangeValues2[i].Value);
                    }
                }
            }
        }
    }
}

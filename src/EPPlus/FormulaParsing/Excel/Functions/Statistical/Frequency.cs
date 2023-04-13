/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  22/3/2023         EPPlus Software AB           EPPlus v7
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "7",
        IntroducedInExcelVersion = "2007",
        Description = "Calculates how often values occur within a range of values, and then returns a vertical array of numbers",
        SupportsArrays = true)]
    internal class Frequency : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var arg1 = arguments.First();
            var arg2 = arguments.ElementAt(1);
            if(!arg1.IsExcelRange || !arg2.IsExcelRange)
            {
                return CreateResult(eErrorType.Value);
            }
            var dataArray = arg1.ValueAsRangeInfo;
            var binsArray = arg2.ValueAsRangeInfo;

            var dataArrayNumbers = GetNumbers(dataArray);
            var binsArrayNumbers = GetNumbers(binsArray);
            var sortedBins = new List<double>(binsArrayNumbers).ToArray();
            Array.Sort(sortedBins);
            var counts = new Dictionary<double, int>();
            var min = sortedBins.Min();
            var max = sortedBins.Max();
            foreach(var b in sortedBins )
            {
                counts[b] = 0;
            }
            counts[max + 1] = 0;
            foreach(var n in dataArrayNumbers)
            {
                if(n <= min)
                {
                    counts[min]++;
                }
                else if(n > max)
                {
                    counts[max + 1]++;
                }
                else
                {
                    var ix = Array.BinarySearch(sortedBins, n);
                    if(ix > -1)
                    {
                        counts[n]++;
                    }
                    else
                    {
                        ix = ~ix;
                        counts[sortedBins[ix + 1]]++;
                    }
                }
            }
            var resultArray = new InMemoryRange(new RangeDefinition(counts.Keys.Count, 1));
            for(var x = 0; x < binsArrayNumbers.Count; x++)
            {
                var b = binsArrayNumbers[x];
                var c = counts[b];
                resultArray.SetValue(x, 0, c);
            }
            resultArray.SetValue(resultArray.Size.NumberOfRows - 1, 0, counts[max + 1]);
            return CreateResult(resultArray, DataType.ExcelRange);

        }

        private List<double> GetNumbers(IRangeInfo range)
        {
            var list = new List<double>();
            for (var row = 0; row < range.Size.NumberOfRows; row++)
            {
                for (var col = 0; col < range.Size.NumberOfCols; col++)
                {
                    var obj = range.GetOffset(row, col);
                    if (obj != null && IsNumeric(obj))
                    {
                        var n = ConvertUtil.GetValueDouble(obj);
                        list.Add(n);
                    }
                }
            }
            return list;
        }
    }
}

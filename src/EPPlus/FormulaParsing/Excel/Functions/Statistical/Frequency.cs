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
        EPPlusVersion = "7.0",
        IntroducedInExcelVersion = "2010",
        Description = "Calculates how often values occur within a range of values, and then returns a vertical array of numbers",
        SupportsArrays = true)]
    internal class Frequency : ExcelFunction
    {
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var arg1 = arguments[0];
            var arg2 = arguments[1];
            if(!arg1.IsExcelRange || !arg2.IsExcelRange)
            {
                return CompileResult.GetDynamicArrayResultError(eErrorType.Value);
            }
            var dataRange = arg1.ValueAsRangeInfo;
            var binsRange = arg2.ValueAsRangeInfo;

            var dataArray = new List<object>();
            foreach(var cell in dataRange)
            {
                if(cell?.Value != null)
                {
                    dataArray.Add(cell.Value);
                }
            }

            var binsArray = new List<object>();
            foreach(var cell in binsRange)
            {
                if(cell?.Value != null)
                {
                    binsArray.Add(cell.Value);
                }
            }

            var resultValues = Calculate(dataArray, binsArray);
            var result = new InMemoryRange(new RangeDefinition(resultValues.Count, 1));
            for(var row = 0; row < resultValues.Count; row++)
            {
                result.SetValue(row, 0, resultValues[row]);
            }
            return CreateDynamicArrayResult(result, DataType.ExcelRange);
        }

        private List<int> Calculate(List<object> objData, List<object> objBinsArray)
        {
            var data = ToDoubles(objData);
            var binsArray = ToDoubles(objBinsArray);
            if (binsArray.Count == 0)
            {
                binsArray.Add(0);
            }
            var sortedBinsArray = binsArray.ToArray();
            Array.Sort(sortedBinsArray);
            var dict = new Dictionary<double, int>();
            foreach (var bin in sortedBinsArray)
            {
                dict.Add(bin, data.Count(x => x <= bin));
                data.RemoveAll(x => x <= bin);
            }
            var result = new List<int>();
            foreach (var b in binsArray)
            {
                result.Add(dict[b]);
            }
            // add last item (larger than the highest value in the binsarray)
            result.Add(data.Count());
            return result;
        }

        private List<double> ToDoubles(List<object> list)
        {
            var result = new List<double>();
            foreach (var item in list)
            {
                if (ConvertUtil.IsNumeric(item))
                {
                    result.Add(ConvertUtil.GetValueDouble(item));
                }
            }
            return result;
        }
    }
}

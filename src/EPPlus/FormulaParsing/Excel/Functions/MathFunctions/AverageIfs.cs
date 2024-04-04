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
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "4",
        Description = "Calculates the Average of the cells in a supplied range, that satisfy multiple criteria",
        IntroducedInExcelVersion = "2007")]
    internal class AverageIfs : MultipleRangeCriteriasFunction
    {
        private object GetCriteraFromArgsByIndex(IList<FunctionArgument> arguments, int index)
        {
            return arguments[index + 1].Value != null ? arguments[index + 1].Value.ToString() : null;
        }

        public override int ArgumentMinLength => 3;
        public override ExcelFunctionParametersInfo ParametersInfo => new ExcelFunctionParametersInfo(new Func<int, FunctionParameterInformation>((argumentIndex) =>
        {
            if (argumentIndex % 2 == 0 && argumentIndex > 0)
            {
                return FunctionParameterInformation.IgnoreErrorInPreExecute;
            }
            return FunctionParameterInformation.Normal;
        }));
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var valueRange = arguments[0].ValueAsRangeInfo;
            var argRanges = new List<RangeOrValue>();
            var criterias = new List<object>();
            for (var ix = 1; ix < 31; ix += 2)
            {
                if (arguments.Count <= ix) break;
                var arg = arguments[ix];
                if (arg.IsExcelRange)
                {
                    var rangeInfo = arg.ValueAsRangeInfo;
                    argRanges.Add(new RangeOrValue { Range = rangeInfo });
                }
                else
                {
                    argRanges.Add(new RangeOrValue { Value = arg.Value });
                }
                criterias.Add(arguments[ix+1].ValueFirst);
            }
            IEnumerable<int> matchIndexes = GetMatchIndexes(argRanges[0], criterias[0], context);
            var enumerable = matchIndexes as IList<int> ?? matchIndexes.ToList();
            for (var ix = 1; ix < argRanges.Count && enumerable.Any(); ix++)
            {
                var indexes = GetMatchIndexes(argRanges[ix], criterias[ix], context, false);
                matchIndexes = matchIndexes.Intersect(indexes);
            }

            if (matchIndexes.Count() == 0) return CreateResult(eErrorType.Div0);
            var sumRange = RangeFlattener.FlattenRangeObject(valueRange);
            KahanSum sum = 0d;
            var count = 0;
            foreach (var index in matchIndexes)
            {
                var obj = sumRange[index];
                if (obj is ExcelErrorValue e1)
                {
                    return e1.AsCompileResult;
                }
                if (ConvertUtil.IsNumericOrDate(obj))
                {
                    sum += ConvertUtil.GetValueDouble(obj);
                    count++;
                }
            }
            if(count == 0)
            {
                return CompileResult.GetErrorResult(eErrorType.Div0);
            }   

            return CreateResult(sum.Get()/count, DataType.Decimal);
        }
    }
}

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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "4",
        Description = "Calculates the Average of the cells in a supplied range, that satisfy multiple criteria",
        IntroducedInExcelVersion = "2007")]
    internal class AverageIfs : MultipleRangeCriteriasFunction
    {
        private string GetCriteraFromArgsByIndex(FunctionArgument[] arguments, int index)
        {
            return arguments[index + 1].Value != null ? arguments[index + 1].Value.ToString() : null;
        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
            ValidateArguments(functionArguments, 3);
            var sumRange = ArgsToDoubleEnumerable(false, new List<FunctionArgument> { functionArguments[0] }, context).ToList();
            var argRanges = new List<RangeOrValue>();
            var criterias = new List<string>();
            for (var ix = 1; ix < 31; ix += 2)
            {
                if (functionArguments.Length <= ix) break;
                var arg = functionArguments[ix];
                if (arg.IsExcelRange)
                {
                    var rangeInfo = arg.ValueAsRangeInfo;
                    argRanges.Add(new RangeOrValue { Range = rangeInfo });
                }
                else
                {
                    argRanges.Add(new RangeOrValue { Value = arg.Value });
                }
                var v = GetCriteraFromArgsByIndex(functionArguments, ix);
                criterias.Add(v);
            }
            IEnumerable<int> matchIndexes = GetMatchIndexes(argRanges[0], criterias[0], context);
            var enumerable = matchIndexes as IList<int> ?? matchIndexes.ToList();
            for (var ix = 1; ix < argRanges.Count && enumerable.Any(); ix++)
            {
                var indexes = GetMatchIndexes(argRanges[ix], criterias[ix], context, false);
                matchIndexes = matchIndexes.Intersect(indexes);
            }

            if (matchIndexes.Count() == 0) return CreateResult(eErrorType.Div0);
            var result = matchIndexes.Average(index => sumRange[index]);

            return CreateResult(result, DataType.Decimal);
        }
    }
}

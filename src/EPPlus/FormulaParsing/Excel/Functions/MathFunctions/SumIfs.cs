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
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.MathAndTrig,
        EPPlusVersion = "4",
        Description = "Adds the cells in a supplied range, that satisfy multiple criteria",
        IntroducedInExcelVersion = "2007")]
    internal class SumIfs : MultipleRangeCriteriasFunction
    {
        public override void ConfigureArrayBehaviour(ArrayBehaviourConfig config)
        {
            config.IgnoreNumberOfArgsFromStart = 1;
            config.ArrayArgInterval = 1;
        }
        public override int ArgumentMinLength => 3;
        public override FunctionParameterInformation GetParameterInfo(int argumentIndex)
        {
            if (argumentIndex % 2 == 0 && argumentIndex > 0)
            {
                return FunctionParameterInformation.IgnoreErrorInPreExecute;
            }
            return FunctionParameterInformation.Normal;
        }

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var rows = new List<int>();
            var valueRange = arguments[0].ValueAsRangeInfo;
            List<double> sumRange;
            if(valueRange != null)
            {
                sumRange = ArgsToDoubleEnumerableZeroPadded(false, valueRange, context).ToList();
            }
            else
            {
                sumRange = ArgsToDoubleEnumerable(false, new List<FunctionArgument> { arguments[0] }, context).Select(x => (double)x).ToList();
            } 
            var argRanges = new List<RangeOrValue>();
            var criterias = new List<object>();
            for (var ix = 1; ix < 31; ix += 2)
            {
                if (arguments.Count <= ix) break;
                var arg = arguments[ix];
                if(arg.IsExcelRange)
                {
                    var rangeInfo = arg.ValueAsRangeInfo;
                    argRanges.Add(new RangeOrValue { Range = rangeInfo });
                }
                else
                {
                    argRanges.Add(new RangeOrValue { Value = arg.Value });
                }
                //var value = arguments[ix + 1].Value != null ? ArgToString(arguments, ix + 1) : null;
                criterias.Add(arguments[ix+1].ValueFirst);
            }
            IEnumerable<int> matchIndexes = GetMatchIndexes(argRanges[0], criterias[0], context);
            var enumerable = matchIndexes as IList<int> ?? matchIndexes.ToList();
            for (var ix = 1; ix < argRanges.Count && enumerable.Any(); ix++)
            {
                var indexes = GetMatchIndexes(argRanges[ix], criterias[ix], context);
                matchIndexes = matchIndexes.Intersect(indexes);
            }

            //var result = matchIndexes.Sum(index => sumRange[index]);
            KahanSum result = 0.0;
            foreach(var index in matchIndexes)
            {
                result += sumRange[index];
            }

            return CreateResult(result.Get(), DataType.Decimal);
        }
    }
}

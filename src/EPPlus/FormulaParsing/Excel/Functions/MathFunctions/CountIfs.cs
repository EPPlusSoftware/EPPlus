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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "4",
        Description = "Returns the number of cells (of a supplied range), that satisfy a set of given criteria",
        IntroducedInExcelVersion = "2007")]
    internal class CountIfs : RangeCriteriaFunction
    {
        public override int ArgumentMinLength => 2;
        public override ExcelFunctionParametersInfo ParametersInfo => new ExcelFunctionParametersInfo(new Func<int, FunctionParameterInformation>((argumentIndex) =>
        {
            if(argumentIndex == 0)
            {
                return FunctionParameterInformation.Normal;
            }
            if (argumentIndex % 2 == 0)
            {
                return FunctionParameterInformation.AdjustParameterAddress;
            }
            return FunctionParameterInformation.IgnoreErrorInPreExecute;
        }));
        public override void GetNewParameterAddress(IList<CompileResult> args, int index, ParsingContext ctx, ref Queue<FormulaRangeAddress> addresses)
        {
            if (args[index].Result is IRangeInfo criteriaRange)
            {
                IEnumerable<int> matchIndexes = GetMatchingIndicesFromArguments(0, args, ctx, index);
                EnqueueMatchingAddresses(criteriaRange, matchIndexes, ref addresses);
            }
        }
        public override void ConfigureArrayBehaviour(ArrayBehaviourConfig config)
        {
            config.IgnoreNumberOfArgsFromStart = 1;
            config.ArrayArgInterval = 2;
        }
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            GetArguments(context, arguments, out List<RangeOrValue> argRanges, out List<RangeOrValue> criteria, out int cols, out int rows, 0);
            if (cols == 1 && rows == 1)
            {
                var result = GetCountValue(context, argRanges, criteria, 0, 0);
                return CreateResult(result, DataType.Decimal);
            }
            else
            {
                var retRange = new InMemoryRange(rows, (short)cols);
                for (var r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        var result = GetCountValue(context, argRanges, criteria, r, c);
                        retRange.SetValue(r, c, result);
                    }
                }
                return CreateDynamicArrayResult(retRange, DataType.ExcelRange);
            }
        }


        private double GetCountValue(ParsingContext context, List<RangeOrValue> argRanges, List<RangeOrValue> criteria, int row, int col)
        {
            var matchIndexes = GetMatchIndexes(argRanges[0], GetCriteriaValue(criteria[0], row, col), context, false);

            // Use HashSet.IntersectWith for multi-criteria (more efficient than LINQ Intersect)
            if (argRanges.Count > 1)
            {
                var hashSet = new HashSet<int>(matchIndexes);
                for (var ix = 1; ix < argRanges.Count && hashSet.Count > 0; ix++)
                {
                    var indexes = GetMatchIndexes(argRanges[ix], GetCriteriaValue(criteria[ix], row, col), context, false);
                    hashSet.IntersectWith(indexes);
                }
                return (double)hashSet.Count;
            }

            return (double)matchIndexes.Count;
        }
    }
}

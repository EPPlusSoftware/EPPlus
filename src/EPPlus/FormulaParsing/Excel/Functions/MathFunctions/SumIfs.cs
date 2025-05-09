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
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.Utils.TypeConversion;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.MathAndTrig,
        EPPlusVersion = "4",
        Description = "Adds the cells in a supplied range, that satisfy multiple criteria",
        IntroducedInExcelVersion = "2007")]
    internal class SumIfs : RangeCriteriaFunction
    {
        public override void ConfigureArrayBehaviour(ArrayBehaviourConfig config)
        {
            config.IgnoreNumberOfArgsFromStart = 1;
            config.ArrayArgInterval = 1;

        }

        public override int ArgumentMinLength => 3;
        public override ExcelFunctionParametersInfo ParametersInfo => new ExcelFunctionParametersInfo(new Func<int, FunctionParameterInformation>((argumentIndex) =>
        {
            if (argumentIndex == 0)
            {
                return FunctionParameterInformation.AdjustParameterAddress;
            }
            if (argumentIndex % 2 == 0)
            {
                return FunctionParameterInformation.IgnoreErrorInPreExecute;
            }
            return FunctionParameterInformation.Normal;
        }));
        public override void GetNewParameterAddress(IList<CompileResult> args, int index, ref Queue<FormulaRangeAddress> addresses)
        {
            if (index == 0)
            {
                IEnumerable<int> matchIndexes = GetMatchingIndicesFromArguments(1, args);
                addresses = EnqueueMatchingAddresses(args[0].Address, matchIndexes);
            }
        }

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var valueRange = arguments[0].ValueAsRangeInfo;           
            GetArguments(context, arguments, out List<RangeOrValue> argRanges, out List<RangeOrValue> criteria, out int cols, out int rows, 1);
            
            if (cols == 1 && rows == 1)
            {
                var result = GetSumValue(context, valueRange, argRanges, criteria, 0, 0, out ExcelErrorValue ev);
                if (double.IsNaN(result) && ev != null)
                {
                    return CreateResult(ev, DataType.ExcelError);
                }
                else
                {
                    return CreateResult(result, DataType.Decimal);
                }
            }
            else
            {
                var retRange = new InMemoryRange(rows, (short)cols);
                for (var r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        var result = GetSumValue(context, valueRange, argRanges, criteria, r, c, out ExcelErrorValue ev);
                        if (double.IsNaN(result) && ev != null)
                        {
                            retRange.SetValue(r, c, ev);
                        }
                        else
                        {
                            retRange.SetValue(r, c, result);
                        }
                    }
                }
                return CreateDynamicArrayResult(retRange, DataType.ExcelRange);
            }
        }

        private double GetSumValue(ParsingContext context, IRangeInfo valueRange, List<RangeOrValue> argRanges, List<RangeOrValue> criterias, int row, int col, out ExcelErrorValue ev)
        {
            IEnumerable<int> matchIndexes = GetMatchIndexes(argRanges[0], GetCriteriaValue(criterias[0], row, col), context);
            var enumerable = matchIndexes as IList<int> ?? matchIndexes.ToList();
            for (var ix = 1; ix < argRanges.Count && enumerable.Any(); ix++)
            {
                var indexes = GetMatchIndexes(argRanges[ix], GetCriteriaValue(criterias[ix], row, col), context);
                matchIndexes = matchIndexes.Intersect(indexes);
            }
            var sumRange = RangeFlattener.FlattenRangeObject(valueRange);
            KahanSum result = 0d;
            foreach (var index in matchIndexes)
            {
                var obj = sumRange[index];
                if (obj is ExcelErrorValue e1)
                {
                    ev = e1;
                    return double.NaN;
                }
                if (ConvertUtil.IsNumericOrDate(obj))
                {
                    result += ConvertUtil.GetValueDouble(obj);
                }
            }
            ev = null;
            return result;
        }
    }
}
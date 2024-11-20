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
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Utils;

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
        public override ExcelFunctionParametersInfo ParametersInfo => new ExcelFunctionParametersInfo(new Func<int, FunctionParameterInformation>((argumentIndex) =>
        {
            if(argumentIndex==0)
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
            if(index == 0)
            {
                //Return the addresses matching the criterias in the queu
                var valueAddress = args[0].Address;
                var argRanges = new List<RangeOrValue>();
                var criterias = new List<object>();
                for (var ix = 1; ix < 31; ix += 2)
                {
                    if (args.Count <= ix) break;
                    var arg = args[ix];
                    if (arg.Result is IRangeInfo rangeInfo)
                    {
                        argRanges.Add(new RangeOrValue { Range = rangeInfo });
                    }
                    else
                    {
                        argRanges.Add(new RangeOrValue { Value = arg.ResultValue });
                    }
                    criterias.Add(args[ix + 1].ResultValue);
                }
                IEnumerable<int> matchIndexes = GetMatchIndexes(argRanges[0], criterias[0], null);
                var enumerable = matchIndexes as IList<int> ?? matchIndexes.ToList();
                for (var ix = 1; ix < argRanges.Count && enumerable.Any(); ix++)
                {
                    var indexes = GetMatchIndexes(argRanges[ix], criterias[ix], null);
                    matchIndexes = matchIndexes.Intersect(indexes);
                }

                addresses = new Queue<FormulaRangeAddress>();
                var pIx = int.MinValue;
                if(valueAddress.FromCol==valueAddress.ToCol)
                {
                    var c = valueAddress.FromCol;
                    foreach (var ix in matchIndexes)
                    {                        
                        if(ix==pIx+1)
                        {
                            addresses.Peek().ToRow++;
                        }
                        else
                        {
                            var r = valueAddress.FromRow + ix;
                            addresses.Enqueue(new FormulaRangeAddress() { FromRow = r, ToRow = r, FromCol = c, ToCol = c });
                        }
                        pIx = ix;
                    }
                }
                else
                {
                    var r = valueAddress.FromRow;
                    foreach (var ix in matchIndexes)
                    {
                        if (ix == pIx + 1)
                        {
                            addresses.Peek().ToCol++;
                        }
                        else
                        {
                            var c = valueAddress.FromCol + ix;
                            addresses.Enqueue(new FormulaRangeAddress() { FromRow = r, ToRow = r, FromCol = c, ToCol = c });
                        }
                        pIx = ix;
                    }
                }
            }
        }
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var valueRange = arguments[0].ValueAsRangeInfo;
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
                criterias.Add(arguments[ix+1].ValueFirst);
            }
            IEnumerable<int> matchIndexes = GetMatchIndexes(argRanges[0], criterias[0], context);
            var enumerable = matchIndexes as IList<int> ?? matchIndexes.ToList();
            for (var ix = 1; ix < argRanges.Count && enumerable.Any(); ix++)
            {
                var indexes = GetMatchIndexes(argRanges[ix], criterias[ix], context);
                matchIndexes = matchIndexes.Intersect(indexes);
            }
            var sumRange = RangeFlattener.FlattenRangeObject(valueRange);
            KahanSum result = 0d;
            foreach (var index in matchIndexes)
            {
                var obj = sumRange[index];
                if (obj is ExcelErrorValue e1)
                {
                    return e1.AsCompileResult;
                }
                if (ConvertUtil.IsNumericOrDate(obj))
                {
                    result += ConvertUtil.GetValueDouble(obj);
                }
            }
            
            return CreateResult(result.Get(), DataType.Decimal);
        }
    }
}

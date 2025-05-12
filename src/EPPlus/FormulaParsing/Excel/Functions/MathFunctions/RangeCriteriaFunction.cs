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
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Sorting.Internal;
using OfficeOpenXml.Utils;
using Require = OfficeOpenXml.FormulaParsing.Utilities.Require;


namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    internal abstract class RangeCriteriaFunction : HiddenValuesHandlingFunction
    {
        protected static void GetArguments(ParsingContext context, IList<FunctionArgument> arguments, out List<RangeOrValue> argRanges, out List<RangeOrValue> criteria, out int cols, out int rows, int startIndex)
        {
            argRanges = new List<RangeOrValue>();
            criteria = new List<RangeOrValue>();
            cols = 1;
            rows = 1;
            for (var ix = startIndex; ix < 30 + startIndex; ix += 2)
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
                    if(arg.Address!=null && arg.Address.FromRow!=arg.Address.ToRow && arg.Address.FromCol != arg.Address.ToCol)
                    {
                        var wsIx = arg.Address.WorksheetIx < 0 ? context.CurrentCell.WorksheetIx : arg.Address.WorksheetIx;
                        var rangeInfo = context.ExcelDataProvider.GetRange(wsIx, arg.Address.FromRow, arg.Address.FromCol);
                        argRanges.Add(new RangeOrValue { Range = rangeInfo });
                    }
                    else
                    {
                        argRanges.Add(new RangeOrValue { Value = arg.Value });
                    }
                }
                var argCriteria = arguments[ix + 1];
                if (argCriteria.IsExcelRange)
                {
                    var rangeInfo = argCriteria.ValueAsRangeInfo;
                    criteria.Add(new RangeOrValue { Range = rangeInfo });
                    if (rangeInfo.GetNCells() > 1)
                    {
                        if (cols < rangeInfo.Size.NumberOfCols)
                        {
                            cols = rangeInfo.Size.NumberOfCols;
                        }
                        if (rows < rangeInfo.Size.NumberOfRows)
                        {
                            rows = rangeInfo.Size.NumberOfRows;
                        }
                    }
                }
                else
                {
                    criteria.Add(new RangeOrValue { Value = argCriteria.Value });
                }
            }
        }

        protected static object GetCriteriaValue(RangeOrValue rangeOrValue, int row, int col)
        {
            if (rangeOrValue.Range == null)
            {
                return rangeOrValue.Value;
            }
            else
            {
                var range = rangeOrValue.Range;
                if (range.Size.NumberOfRows == 1)
                {
                    return range.Size.NumberOfCols > col ? range.GetOffset(0, col) : null;
                }
                if (range.Size.NumberOfCols == 1)
                {
                    return range.Size.NumberOfRows > row ? range.GetOffset(row, 0) : null;
                }
                else if (range.Size.NumberOfCols > col && range.Size.NumberOfCols > col)
                {
                    return range.GetOffset(row, col);
                }
                return null;
            }
        }

        protected bool Evaluate(object obj, object expression, ParsingContext ctx, bool convertNumericString = true)
        {
            if(expression is ExcelErrorValue e)
            {
                if (obj == null) return false;
                return obj.Equals(e);
            }

            if(obj is bool b1)
            {
                if(expression is bool b2)
                {
                    return b1 == b2;
                }
                else if(expression != null && bool.TryParse(expression.ToString(), out bool b3))
                {
                    return b1 == b3;
                }
                return false;
            }

            var expressionEvaluator = new ExpressionEvaluator(ctx);
            double? candidate = default(double?);
            if (IsNumeric(obj))
            {
                candidate = ConvertUtil.GetValueDouble(obj);
                if(IsNumeric(expression))
                {
                    var dblE = ConvertUtil.GetValueDouble(expression);
                    var compResult = candidate.Value.CompareTo(dblE);
                    return compResult == 0;
                }
            }

            var expressionString = expression==null ? string.Empty : expression.ToString();
            if (candidate.HasValue)
            {
                return expressionEvaluator.Evaluate(candidate.Value, expressionString, convertNumericString);
            }
            return expressionEvaluator.Evaluate(obj, expressionString, convertNumericString);        
        }
        protected List<int> GetMatchIndexes(RangeOrValue rangeOrValue, object searched, ParsingContext ctx, bool convertNumericString = true)
        {
            var expressionEvaluator = new ExpressionEvaluator(ctx);
            var result = new List<int>();
            var internalIndex = 0;
            if (rangeOrValue.Range != null)
            {
                var rangeInfo = rangeOrValue.Range;
                var address = rangeInfo.GetAddressDimensionAdjusted(0).Address;
                for (var row = address.FromRow; row <= address.ToRow; row++)
                {
                    for (var col = address.FromCol; col <= address.ToCol; col++)
                    {
                        var candidate = rangeInfo.GetValue(row, col);
                        if (searched != null && Evaluate(candidate, searched, ctx, convertNumericString))
                        {
                            result.Add(internalIndex);
                        }
                        internalIndex++;
                    }
                }
            }
            else if(Evaluate(rangeOrValue.Value, searched, ctx, convertNumericString))
            {
                result.Add(internalIndex);
            }
            return result;
        }
        protected static Queue<FormulaRangeAddress> EnqueueMatchingAddresses(FormulaRangeAddress valueAddress, IEnumerable<int> matchIndexes)
        {
            Queue<FormulaRangeAddress> addresses = new Queue<FormulaRangeAddress>();
            var pIx = int.MinValue;
            var extRef = valueAddress.ExternalReferenceIx;
            var wsIx = valueAddress.WorksheetIx;
            if (valueAddress.FromCol == valueAddress.ToCol)
            {
                var c = valueAddress.FromCol;
                foreach (var ix in matchIndexes)
                {
                    if (ix == pIx + 1)
                    {
                        addresses.Peek().ToRow++;
                    }
                    else
                    {
                        var r = valueAddress.FromRow + ix;
                        addresses.Enqueue(new FormulaRangeAddress() { ExternalReferenceIx=extRef, WorksheetIx = wsIx, FromRow = r, ToRow = r, FromCol = c, ToCol = c });
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
                        addresses.Enqueue(new FormulaRangeAddress() { ExternalReferenceIx = extRef, WorksheetIx = wsIx, FromRow = r, ToRow = r, FromCol = c, ToCol = c });
                    }
                    pIx = ix;
                }
            }

            return addresses;
        }
        protected IEnumerable<int> GetMatchingIndicesFromArguments(int argStartIx, IList<CompileResult> args)
        {
            //Return the addresses matching the criteria in the queue
            var argRanges = new List<RangeOrValue>();
            var criteria = new List<object>();
            for (var ix = argStartIx; ix < 31; ix += 2)
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
                criteria.Add(args[ix + 1].ResultValue);
            }
            IEnumerable<int> matchIndexes = GetMatchIndexes(argRanges[0], criteria[0], null);
            var enumerable = matchIndexes as IList<int> ?? matchIndexes.ToList();
            for (var ix = 1; ix < argRanges.Count && enumerable.Any(); ix++)
            {
                var indexes = GetMatchIndexes(argRanges[ix], criteria[ix], null);
                matchIndexes = matchIndexes.Intersect(indexes);
            }

            return matchIndexes;
        }

    }
}

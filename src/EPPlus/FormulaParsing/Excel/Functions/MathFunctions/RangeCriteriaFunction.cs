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
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Sorting.Internal;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
                    if (arg.Address != null && arg.Address.FromRow != arg.Address.ToRow && arg.Address.FromCol != arg.Address.ToCol)
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
            if (expression is ExcelErrorValue e)
            {
                if (obj == null) return false;
                return obj.Equals(e);
            }

            if (obj is bool b1)
            {
                if (expression is bool b2)
                {
                    return b1 == b2;
                }
                else if (expression != null && bool.TryParse(expression.ToString(), out bool b3))
                {
                    return b1 == b3;
                }
                return false;
            }

            var expressionEvaluator = ctx.ExpressionEvaluator;
            double? candidate = default(double?);
            if (IsNumeric(obj))
            {
                candidate = ConvertUtil.GetValueDouble(obj);
                if (IsNumeric(expression))
                {
                    var dblE = ConvertUtil.GetValueDouble(expression);
                    var compResult = candidate.Value.CompareTo(dblE);
                    return compResult == 0;
                }
            }

            var expressionString = expression == null ? string.Empty : expression.ToString();
            if (candidate.HasValue)
            {
                return expressionEvaluator.Evaluate(candidate.Value, expressionString, convertNumericString);
            }
            return expressionEvaluator.Evaluate(obj, expressionString, convertNumericString);
        }

        protected List<int> GetMatchIndexes(RangeOrValue rangeOrValue, object searched, ParsingContext ctx, bool convertNumericString = true)
        {
            // Try to get from cache if we have a range with address
            if (rangeOrValue.Range != null && rangeOrValue.Range.Address != null && searched != null)
            {
                var cached = ctx.RangeCriteriaCache?.GetMatchIndexes(rangeOrValue.Range.Address, searched);
                if (cached != null)
                {
                    return cached;
                }
            }

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

                        if (searched != null)
                        {
                            if (searched is RangeOrValue critRange)
                            {
                                // Handle range criteria (less common case) - use original Evaluate
                                if (critRange.Range != null)
                                {
                                    foreach (var cell in critRange.Range)
                                    {
                                        if (Evaluate(candidate, cell.Value, ctx, convertNumericString))
                                        {
                                            result.Add(internalIndex);
                                        }
                                    }
                                }
                                else if (critRange.Value != null && Evaluate(candidate, critRange.Value, ctx, convertNumericString))
                                {
                                    result.Add(internalIndex);
                                }
                            }
                            else if (Evaluate(candidate, searched, ctx, convertNumericString))
                            {
                                result.Add(internalIndex);
                            }
                        }

                        internalIndex++;
                    }
                }
            }
            else if (searched != null && Evaluate(rangeOrValue.Value, searched, ctx, convertNumericString))
            {
                result.Add(internalIndex);
            }

            // Cache the result if we have a range with address
            // Pass rangeHasFormulas flag so cache knows whether to store it
            if (rangeOrValue.Range != null && rangeOrValue.Range.Address != null && searched != null)
            {
                ctx.RangeCriteriaCache?.SetMatchIndexes(rangeOrValue.Range.Address, searched, result);
            }

            return result;
        }

        protected static Queue<FormulaRangeAddress> EnqueueMatchingAddresses(IRangeInfo valueRange, IEnumerable<int> matchIndexes, ref Queue<FormulaRangeAddress> addresses)
        {
            if (addresses == null)
            {
                addresses = new Queue<FormulaRangeAddress>();
            }
            var pIx = int.MinValue;
            var valueAddress = valueRange.GetAddressDimensionAdjusted(0);
            var extRef = valueAddress.ExternalReferenceIx;
            var wsIx = valueAddress.WorksheetIx;
            FormulaRangeAddress currentAddress = null;
            if (valueAddress.FromCol == valueAddress.ToCol)
            {
                var c = valueAddress.FromCol;
                foreach (var ix in matchIndexes)
                {
                    if (ix == pIx + 1)
                    {
                        currentAddress.ToRow++;
                    }
                    else
                    {
                        var r = valueAddress.FromRow + ix;
                        currentAddress = new FormulaRangeAddress() { ExternalReferenceIx = extRef, WorksheetIx = wsIx, FromRow = r, ToRow = r, FromCol = c, ToCol = c };
                        addresses.Enqueue(currentAddress);
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
                        currentAddress.ToCol++;
                    }
                    else
                    {
                        var c = valueAddress.FromCol + ix;
                        currentAddress = new FormulaRangeAddress() { ExternalReferenceIx = extRef, WorksheetIx = wsIx, FromRow = r, ToRow = r, FromCol = c, ToCol = c };
                        addresses.Enqueue(currentAddress);
                    }
                    pIx = ix;
                }
            }

            return addresses;
        }
        protected IEnumerable<int> GetMatchingIndicesFromArguments(int argStartIx, IList<CompileResult> args, ParsingContext ctx, int maxIndex = 31)
        {
            //Return the addresses matching the criteria in the queue
            var argRanges = new List<RangeOrValue>();
            var criteria = new List<object>();
            for (var ix = argStartIx; ix < maxIndex; ix += 2)
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

                if (args[ix + 1].Result is IRangeInfo critInfo)
                {
                    criteria.Add(new RangeOrValue { Range = critInfo });
                }
                else
                {
                    criteria.Add(new RangeOrValue { Value = args[ix + 1].ResultValue });
                }
            }
            IEnumerable<int> matchIndexes = GetMatchIndexes(argRanges[0], criteria[0], ctx);
            var enumerable = matchIndexes as IList<int> ?? matchIndexes.ToList();
            for (var ix = 1; ix < argRanges.Count && enumerable.Any(); ix++)
            {
                var indexes = GetMatchIndexes(argRanges[ix], criteria[ix], ctx);
                matchIndexes = matchIndexes.Intersect(indexes);
            }

            return matchIndexes;
        }

        protected void GetFilteredValueRange(
            ParsingContext context, 
            IRangeInfo valueRange, 
            List<RangeOrValue> argRanges,
            List<RangeOrValue> criterias,
            int row,
            int col,
            out List<int> matchIndexes,
            out List<object> flattenedRange,
            bool convertNumericString = true)
        {
            matchIndexes = GetMatchIndexes(argRanges[0], GetCriteriaValue(criterias[0], row, col), context, convertNumericString);

            if (argRanges.Count > 1)
            {
                var hashSet = new HashSet<int>(matchIndexes);
                for (var ix = 1; ix < argRanges.Count && hashSet.Count > 0; ix++)
                {
                    var indexes = GetMatchIndexes(argRanges[ix], GetCriteriaValue(criterias[ix], row, col), context, convertNumericString);
                    hashSet.IntersectWith(indexes);
                }
                matchIndexes = hashSet.ToList();
            }

            // Try to get flattened range from cache only if it doesn't have formulas
            var valueAddress = valueRange.Address;
            if (valueAddress != null && !valueAddress.HasFormulas(context.Package))
            {
                flattenedRange = context.RangeCriteriaCache?.GetFlattenedRange(valueAddress);
                if (flattenedRange == null)
                {
                    // Not in cache, flatten and cache it
                    flattenedRange = RangeFlattener.FlattenRangeObject(valueRange);
                    context.RangeCriteriaCache?.SetFlattenedRange(valueAddress, flattenedRange);
                }
            }
            else
            {
                // Has formulas or no address - don't cache, always flatten fresh
                flattenedRange = RangeFlattener.FlattenRangeObject(valueRange);
            }
        }
    }
}
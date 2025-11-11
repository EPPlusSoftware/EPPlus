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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Linq;
using static OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering.Conversions;
using Require = OfficeOpenXml.FormulaParsing.Utilities.Require;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "4",
        Description = "Calculates the Average of the cells in a supplied range, that satisfy a given criteria",
        IntroducedInExcelVersion = "2007")]
    internal class AverageIf : RangeCriteriaFunction
    {
        private ExpressionEvaluator _expressionEvaluator;

        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.Custom;

        public override void ConfigureArrayBehaviour(ArrayBehaviourConfig config)
        {
            config.SetArrayParameterIndexes(1);
        }
        public override void GetNewParameterAddress(IList<CompileResult> args, int index, ref Queue<FormulaRangeAddress> addresses)
        {
            if (index == 2)
            {
                if (args[0].Result is IRangeInfo rangeInfo && args[2].Result is IRangeInfo valueRange)
                {
                    var rv = new RangeOrValue { Range = rangeInfo };
                    var mi = GetMatchIndexes(rv, args[1].Result, null);
                    EnqueueMatchingAddresses(valueRange, mi, ref addresses);
                }
            }
        }

        private bool Evaluate(object obj, string expression)
        {
            double? candidate = default(double?);
            if (IsNumeric(obj))
            {
                candidate = ConvertUtil.GetValueDouble(obj);
            }
            if (candidate.HasValue)
            {
                return _expressionEvaluator.Evaluate(candidate.Value, expression);
            }
            return _expressionEvaluator.Evaluate(obj, expression);
        }

        private string GetCriteraFromArg(IList<FunctionArgument> arguments)
        {
            return arguments.ElementAt(1).ValueFirst != null ? ArgToString(arguments, 1) : null;
        }

        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            _expressionEvaluator = new ExpressionEvaluator(context);
            var argRange = ArgToRangeInfo(arguments, 0);
            var criteria = GetCriteraFromArg(arguments);
            double returnValue;
            if (argRange == null)
            {
                var val = arguments[0].Value;
                if (criteria != null && Evaluate(val, criteria))
                {
                    var lookupRange = ArgToRangeInfo(arguments, 2);
                    returnValue = arguments.Count > 2
                        ? lookupRange.First().ValueDouble
                        : ConvertUtil.GetValueDouble(val, true);
                }
                else
                {
                    return CompileResult.GetErrorResult(eErrorType.Div0);
                }
            }
            else if (arguments.Count > 2)
            {
                var lookupRange = ArgToRangeInfo(arguments, 2);
                returnValue = CalculateWithLookupRange(argRange, criteria, lookupRange, context, out ExcelErrorValue eev);
                if (eev != null)
                {
                    return GetResultByObject(eev);
                }
            }
            else
            {
                returnValue = CalculateSingleRange(argRange, criteria, context, out ExcelErrorValue eev);
                if (eev != null)
                {
                    return GetResultByObject(eev);
                }
            }
            return CreateResult(returnValue, DataType.Decimal);
        }

        private double CalculateWithLookupRange(IRangeInfo argRange, string criteria, IRangeInfo sumRange, ParsingContext context, out ExcelErrorValue error)
        {
            error = null;
            KahanSum returnValue = 0d;
            var nMatches = 0;
            if(criteria=="")
            {
                var adrDA = argRange.GetAddressDimensionAdjusted(0).Address;
                for (int r = adrDA.FromRow; r <= adrDA.ToRow; r++)
                {
                    for (int c = adrDA.FromCol; c <= adrDA.ToCol; c++)
                    {
                        SumAndCount(r, c, argRange.GetValue(r, c), criteria, argRange, sumRange, ref returnValue, ref nMatches, out ExcelErrorValue eev);
                    }
                }
            }
            else
            {
                foreach (var cell in argRange)
                {
                    SumAndCount(cell.Row, cell.Column, cell.Value, criteria, argRange, sumRange, ref returnValue, ref nMatches, out ExcelErrorValue eev);
                }
            }
            var div = Divide(returnValue.Get(), nMatches);
            if (double.IsPositiveInfinity(div))
            {
                error = ExcelErrorValue.Create(eErrorType.Div0);
                return double.NaN;
            }
            return div;
        }

        private void SumAndCount(int row, int col, object value, string criteria, IRangeInfo argRange, IRangeInfo sumRange, ref KahanSum returnValue, ref int nMatches, out ExcelErrorValue? eev)
        {
            if (criteria != null && Evaluate(value, criteria))
            {
                var rowOffset = row - argRange.Address.FromRow;
                var columnOffset = col - argRange.Address.FromCol;
                if (sumRange.Address.FromRow + rowOffset <= sumRange.Address.ToRow &&
                   sumRange.Address.FromCol + columnOffset <= sumRange.Address.ToCol)
                {
                    var val = sumRange.GetOffset(rowOffset, columnOffset);
                    if (val is ExcelErrorValue err)
                    {
                        eev = err;
                        return;
                    }
                    if (ConvertUtil.IsExcelNumeric(val))
                    {
                        nMatches++;
                        returnValue += ConvertUtil.GetValueDouble(val, true);
                    }
                }
            }

            eev = null;
        }

        private double CalculateSingleRange(IRangeInfo range, string expression, ParsingContext context, out ExcelErrorValue error)
        {
            error = null;
            KahanSum returnValue = 0d;
            var nMatches = 0;
            foreach (var candidate in range)
            {
                if (expression != null && IsNumeric(candidate.Value) && Evaluate(candidate.Value, expression))
                {

                    if (candidate.IsExcelError)
                    {
                        error = (ExcelErrorValue)candidate.Value;
                        return double.NaN;
                    }
                    if (ConvertUtil.IsExcelNumeric(candidate.Value))
                    {
                        returnValue += candidate.ValueDouble;
                        nMatches++;
                    }
                }
            }
            var div = Divide(returnValue.Get(), nMatches);
            if (double.IsPositiveInfinity(div))
            {
                error = ExcelErrorValue.Create(eErrorType.Div0);
                return double.NaN;
            }
            return div;
        }
        public override ExcelFunctionParametersInfo ParametersInfo => new ExcelFunctionParametersInfo(new Func<int, FunctionParameterInformation>((argumentIndex) =>
        {
            if (argumentIndex == 1)
            {
                return FunctionParameterInformation.IgnoreErrorInPreExecute;
            }
            else if (argumentIndex == 2)
            {
                return FunctionParameterInformation.AdjustParameterAddress;
            }
            return FunctionParameterInformation.Normal;
        }));

    }
}

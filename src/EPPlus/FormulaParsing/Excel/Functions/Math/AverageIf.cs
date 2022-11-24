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
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;
using Require = OfficeOpenXml.FormulaParsing.Utilities.Require;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "4",
        Description = "Calculates the Average of the cells in a supplied range, that satisfy a given criteria",
        IntroducedInExcelVersion = "2007")]
    internal class AverageIf : HiddenValuesHandlingFunction
    {
        private ExpressionEvaluator _expressionEvaluator;
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

        private string GetCriteraFromArg(IEnumerable<FunctionArgument> arguments)
        {
            return arguments.ElementAt(1).ValueFirst != null ? ArgToString(arguments, 1) : null;
        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            _expressionEvaluator = new ExpressionEvaluator(context);
            var argRange = ArgToRangeInfo(arguments, 0);
            var criteria = GetCriteraFromArg(arguments);
            double returnValue;
            if (argRange == null)
            {
                var val = arguments.ElementAt(0).Value;
                if (criteria != null && Evaluate(val, criteria))
                {
                    var lookupRange = ArgToRangeInfo(arguments, 2);
                    returnValue = arguments.Count() > 2
                        ? lookupRange.First().ValueDouble
                        : ConvertUtil.GetValueDouble(val, true);
                }
                else
                {
                    throw new ExcelErrorValueException(eErrorType.Div0);
                }
            }
            else if (arguments.Count() > 2)
            {
                var lookupRange = ArgToRangeInfo(arguments, 2);
                returnValue = CalculateWithLookupRange(argRange, criteria, lookupRange, context);
            }
            else
            {
                returnValue = CalculateSingleRange(argRange, criteria, context);
            }
            return CreateResult(returnValue, DataType.Decimal);
        }

        private double CalculateWithLookupRange(IRangeInfo argRange, string criteria, IRangeInfo sumRange, ParsingContext context)
        {
            var returnValue = 0d;
            var nMatches = 0;
            foreach (var cell in argRange)
            {
                if (criteria != null && Evaluate(cell.Value, criteria))
                {
                    var rowOffset = cell.Row - argRange.Address.FromRow;
                    var columnOffset = cell.Column - argRange.Address.FromCol;
                    if (sumRange.Address.FromRow + rowOffset <= sumRange.Address.ToRow &&
                       sumRange.Address.FromCol + columnOffset <= sumRange.Address.ToCol)
                    {
                        var val = sumRange.GetOffset(rowOffset, columnOffset);
                        if (val is ExcelErrorValue)
                        {
                            ThrowExcelErrorValueException(((ExcelErrorValue)val));
                        }
                        nMatches++;
                        returnValue += ConvertUtil.GetValueDouble(val, true);
                    }
                }
            }
            return Divide(returnValue, nMatches);
        }

        private double CalculateSingleRange(IRangeInfo range, string expression, ParsingContext context)
        {
            var returnValue = 0d;
            var nMatches = 0;
            foreach (var candidate in range)
            {
                if (expression != null && IsNumeric(candidate.Value) && Evaluate(candidate.Value, expression))
                {
                    
                    if (candidate.IsExcelError)
                    {
                        ThrowExcelErrorValueException(((ExcelErrorValue)candidate.Value));
                    }
                    returnValue += candidate.ValueDouble;
                    nMatches++;
                }
            }
            return Divide(returnValue, nMatches);
        }
    }
}

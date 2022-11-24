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
        Category = ExcelFunctionCategory.MathAndTrig,
        EPPlusVersion = "4",
        Description = "Adds the cells in a supplied range, that satisfy a given criteria")]
    internal class SumIf : HiddenValuesHandlingFunction
    {
        private ExpressionEvaluator _evaluator;

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            _evaluator = new ExpressionEvaluator(context);
            ValidateArguments(arguments, 2);
            var argRange = ArgToRangeInfo(arguments, 0);

            // Criteria can either be a string or an array of strings
            var criteria = GetCriteria(arguments.ElementAt(1));
            var retVal = 0d;
            if (argRange == null)
            {
                var val = arguments.ElementAt(0).Value;
                if (_evaluator.Evaluate(val, criteria))
                {
                    if (arguments.Count() > 2)
                    {
                        var sumRange = ArgToRangeInfo(arguments, 2);
                        retVal = sumRange.First().ValueDouble;
                    }
                    else
                    {
                        retVal = ConvertUtil.GetValueDouble(val, true);
                    }
                }
            }
            else if (arguments.Count() > 2)
            {
                var sumRange = ArgToRangeInfo(arguments, 2);
                retVal = CalculateWithSumRange(argRange, criteria, sumRange, context);
            }
            else
            {
                retVal = CalculateSingleRange(argRange, criteria, context);
            }
            return CreateResult(retVal, DataType.Decimal);
        }

        internal static IEnumerable<string> GetCriteria(FunctionArgument criteriaArg)
        {
            var criteria = new List<string>();
            if (criteriaArg.IsEnumerableOfFuncArgs)
            {
                foreach (var arg in criteriaArg.ValueAsEnumerableOfFuncArgs)
                {
                    criteria.Add(arg.ValueFirstString);
                }
            }
            else if (criteriaArg.IsExcelRange)
            {
                foreach (var cell in criteriaArg.ValueAsRangeInfo)
                {
                    if (cell.Value != null)
                    {
                        criteria.Add(cell.Value.ToString());
                    }
                }
            }
            else
            {
                criteria.Add(criteriaArg.ValueFirst != null ? criteriaArg.ValueFirst.ToString() : null);
            }
            return criteria;
        }

        private double CalculateWithSumRange(IRangeInfo range, IEnumerable<string> criteria, IRangeInfo sumRange, ParsingContext context)
        {
            var retVal = 0d;
            foreach (var cell in range)
            {
                if (_evaluator.Evaluate(cell.Value, criteria))
                {
                    var rowOffset = cell.Row - range.Address.FromRow;
                    var columnOffset = cell.Column - range.Address.FromCol;
                    if (sumRange.Address.FromRow + rowOffset <= sumRange.Address.ToRow &&
                       sumRange.Address.FromCol + columnOffset <= sumRange.Address.ToCol)
                    {
                        var val = sumRange.GetOffset(rowOffset, columnOffset);
                        if (val is ExcelErrorValue)
                        {
                            ThrowExcelErrorValueException((ExcelErrorValue)val);
                        }
                        retVal += ConvertUtil.GetValueDouble(val, true);
                    }
                }
            }
            return retVal;
        }

        private double CalculateSingleRange(IRangeInfo range, IEnumerable<string> expressions, ParsingContext context)
        {
            var retVal = 0d;
            foreach (var candidate in range)
            {
                if (IsNumeric(candidate.Value) && _evaluator.Evaluate(candidate.Value, expressions) && IsNumeric(candidate.Value))
                {
                    if (candidate.IsExcelError)
                    {
                        ThrowExcelErrorValueException((ExcelErrorValue)candidate.Value);
                    }
                    retVal += candidate.ValueDouble;
                }
            }
            return retVal;
        }
    }
}

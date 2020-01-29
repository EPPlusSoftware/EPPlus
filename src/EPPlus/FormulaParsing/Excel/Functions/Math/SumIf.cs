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
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;
using Require = OfficeOpenXml.FormulaParsing.Utilities.Require;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    internal class SumIf : HiddenValuesHandlingFunction
    {
        private readonly ExpressionEvaluator _evaluator;

        public SumIf()
            : this(new ExpressionEvaluator())
        {

        }

        public SumIf(ExpressionEvaluator evaluator)
        {
            Require.That(evaluator).Named("evaluator").IsNotNull();
            _evaluator = evaluator;
        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var argRange = ArgToRangeInfo(arguments, 0);
            var criteria = arguments.ElementAt(1).ValueFirst != null ? ArgToString(arguments, 1) : null;
            var retVal = 0d;
            if (argRange == null)
            {
                var val = arguments.ElementAt(0).Value;
                if (criteria != default(string) && _evaluator.Evaluate(val, criteria))
                {
                    var sumRange = ArgToRangeInfo(arguments, 2);
                    retVal = arguments.Count() > 2
                        ? sumRange.First().ValueDouble
                        : ConvertUtil.GetValueDouble(val, true);
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

        private double CalculateWithSumRange(ExcelDataProvider.IRangeInfo range, string criteria, ExcelDataProvider.IRangeInfo sumRange, ParsingContext context)
        {
            var retVal = 0d;
            foreach (var cell in range)
            {
                if (criteria != default(string) && _evaluator.Evaluate(cell.Value, criteria))
                {
                    var rowOffset = cell.Row - range.Address._fromRow;
                    var columnOffset = cell.Column - range.Address._fromCol;
                    if (sumRange.Address._fromRow + rowOffset <= sumRange.Address._toRow &&
                       sumRange.Address._fromCol + columnOffset <= sumRange.Address._toCol)
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

        private double CalculateSingleRange(ExcelDataProvider.IRangeInfo range, string expression, ParsingContext context)
        {
            var retVal = 0d;
            foreach (var candidate in range)
            {
                if (expression != default(string) && IsNumeric(candidate.Value) && _evaluator.Evaluate(candidate.Value, expression) && IsNumeric(candidate.Value))
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

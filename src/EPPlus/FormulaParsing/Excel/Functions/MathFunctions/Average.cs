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
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "4",
        Description = "Returns the Average of a list of supplied numbers")]
    internal class Average : HiddenValuesHandlingFunction
    {
        public Average()
        {
            IgnoreErrors = false;
        }
        public override ExcelFunctionParametersInfo ParametersInfo => new ExcelFunctionParametersInfo(new Func<int, FunctionParameterInformation>((argumentIndex) =>
        {
            return FunctionParameterInformation.IgnoreErrorInPreExecute;
        }));

        public override int ArgumentMinLength => 1;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            double nValues = 0d;
            KahanSum result = 0d;
            foreach (var arg in arguments)
            {
                if (ShouldIgnore(arg, context)) continue;
                Calculate(arg, context, ref result, ref nValues);
            }
            var div = Divide(result.Get(), nValues);
            if (double.IsPositiveInfinity(div))
            {
                return CompileResult.GetErrorResult(eErrorType.Div0);
            }
            
            return CreateResult(div, DataType.Decimal);
        }

        private void Calculate(FunctionArgument arg, ParsingContext context, ref KahanSum retVal, ref double nValues, bool isInArray = false)
        {
            if (ShouldIgnore(arg, context))
            {
                return;
            }
            if (arg.Value is IEnumerable<FunctionArgument>)
            {
                foreach (var item in (IEnumerable<FunctionArgument>)arg.Value)
                {
                    Calculate(item, context, ref retVal, ref nValues, true);
                }
            }
            else if (arg.IsExcelRange)
            {
                foreach (var c in arg.ValueAsRangeInfo)
                {
                    if (ShouldIgnore(c, context)) continue;
                    CheckForAndHandleExcelError(c);
                    if (!IsNumeric(c.Value) || IsBool(c.Value)) continue;
                    nValues++;
                    retVal += c.ValueDouble;
                }
            }
            else
            {
                var numericValue = GetNumericValue(arg.Value, isInArray);
				if (numericValue.HasValue)
				{
					nValues++;
					retVal += numericValue.Value;
				}
				else if (IsString(arg.Value) && !isInArray)
				{
					ThrowExcelErrorValueException(eErrorType.Value);
				}
            }
            CheckForAndHandleExcelError(arg);
        }

        private double? GetNumericValue(object obj, bool isInArray)
        {
            if (IsNumeric(obj) && !(IsBool(obj)))
            {
                return ConvertUtil.GetValueDouble(obj);
            }
			if (!isInArray)
			{
                if (IsBool(obj))
                {
                    return ConvertUtil.GetValueDouble(obj);
                }
                else if (IsString(obj, false) && ConvertUtil.TryParseNumericString(obj.ToString(), out double number))
                {
                    return number;
                }
                else if (IsString(obj, false) && ConvertUtil.TryParseDateString(obj.ToString(), out DateTime date))
                {
                    return date.ToOADate();
                }
            }
            return default;
        }
    }
}

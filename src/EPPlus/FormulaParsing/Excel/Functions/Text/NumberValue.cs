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
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Text,
        EPPlusVersion = "5.0",
        Description = "Converts text to a number, in a locale-independent way",
        IntroducedInExcelVersion = "2013",
        SupportsArrays = true)]
    internal class NumberValue : ExcelFunction
    {
        private string _decimalSeparator = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
        private string _groupSeparator = CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator;
        private string _arg = string.Empty;
        private int _nPercentage = 0;

        internal override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.FirstArgCouldBeARange;

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var arg = ArgToString(arguments, 0);
            SetArgAndPercentage(arg);
            if(!ValidateAndSetSeparators(arguments.ToArray()))
            {
                return CreateResult(ExcelErrorValue.Values.Value, DataType.ExcelError);
            }
            var cultureInfo = new CultureInfo("en-US", true);
            cultureInfo.NumberFormat.NumberDecimalSeparator = _decimalSeparator;
            cultureInfo.NumberFormat.NumberGroupSeparator = _groupSeparator;
            if(double.TryParse(_arg, NumberStyles.Any, cultureInfo, out double result))
            {
                if(_nPercentage > 0)
                {
                    result /= System.Math.Pow(100, _nPercentage);
                }
                return CreateResult(result, DataType.Decimal);
            }
            return CreateResult(ExcelErrorValue.Values.Value, DataType.ExcelError);
        }

        private void SetArgAndPercentage(string arg)
        {
            var pIndex = arg.IndexOf("%", StringComparison.OrdinalIgnoreCase);
            if(pIndex > 0)
            {
                _arg = arg.Substring(0, pIndex).Replace(" ", "");
                var percentage = arg.Substring(pIndex, arg.Length - pIndex).Trim();
                if (!Regex.IsMatch(percentage, "[%]+"))
                    throw new ArgumentException("Invalid format: " + arg);
                _nPercentage = percentage.Length;
            }
            else
            {
                _arg = arg;
            }
        }

        private bool ValidateAndSetSeparators(FunctionArgument[] arguments)
        {
            if (arguments.Length == 1) return true;
            var decimalSeparator = ArgToString(arguments, 1).Substring(0, 1);
            if (!DecimalSeparatorIsValid(decimalSeparator))
            {
                return false;
            }
            _decimalSeparator = decimalSeparator;
            if (arguments.Length > 2)
            {
                var groupSeparator = ArgToString(arguments, 2).Substring(0, 1);
                if(!GroupSeparatorIsValid(decimalSeparator, groupSeparator))
                {
                    return false;
                }
                _groupSeparator = groupSeparator;
            }
            return true;
        }

        private bool DecimalSeparatorIsValid(string separator)
        {
            return !string.IsNullOrEmpty(separator)
                &&
                (separator == "." || separator == ",");
        }

        private bool GroupSeparatorIsValid(string groupSeparator, string decimalSeparator)
        {
            return !string.IsNullOrEmpty(groupSeparator)
                &&
                (groupSeparator != decimalSeparator)
                &&
                (groupSeparator == " " || groupSeparator == "," || groupSeparator == ".");
        }
    }
}

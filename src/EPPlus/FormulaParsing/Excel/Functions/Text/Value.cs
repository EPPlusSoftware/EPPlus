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
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Text,
        EPPlusVersion = "4",
        Description = "Converts a text string into a numeric value")]
    internal class Value : ExcelFunction
    {
        public Value(CultureInfo ci)
        {
            _cultureInfo = ci;
            _groupSeparator = _cultureInfo.NumberFormat.NumberGroupSeparator;
            _decimalSeparator = _cultureInfo.NumberFormat.NumberDecimalSeparator;
            _timeSeparator = _cultureInfo.DateTimeFormat.TimeSeparator;
            _shortTimePattern = _cultureInfo.DateTimeFormat.ShortTimePattern;
        }

        private readonly CultureInfo _cultureInfo;
        private readonly string _groupSeparator;
        private readonly string _decimalSeparator;
        private readonly string _timeSeparator;
        private readonly string _shortTimePattern;
        private readonly DateValue _dateValueFunc = new DateValue();
        private readonly TimeValue _timeValueFunc = new TimeValue();

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var val = ArgToString(arguments, 0);
            double result = 0d;
            if (string.IsNullOrEmpty(val)) return CreateResult(result, DataType.Integer);
            val = val.TrimEnd(' ');
            bool isPercentage = false;
            if(val.EndsWith("%"))
            {
                val = val.TrimEnd('%');
                isPercentage = true;
            }
            if(val.StartsWith("(", StringComparison.OrdinalIgnoreCase) && val.EndsWith(")", StringComparison.OrdinalIgnoreCase))
            {
                var numCandidate = val.Substring(1, val.Length - 2);
                if(double.TryParse(numCandidate, NumberStyles.Any, _cultureInfo, out double tmp))
                {
                    val = "-" + numCandidate;
                }
            }
            if (Regex.IsMatch(val, $"^[\\d]*({Regex.Escape(_groupSeparator)}?[\\d]*)*?({Regex.Escape(_decimalSeparator)}[\\d]*)?[ ?% ?]?$", RegexOptions.Compiled))
            {
                result = double.Parse(val, _cultureInfo);
                return CreateResult(isPercentage ? result/100 : result, DataType.Decimal);
            }
            if (double.TryParse(val, NumberStyles.Float, _cultureInfo, out result))
            {
                return CreateResult(isPercentage ? result/100d : result, DataType.Decimal);
            }
            var timeSeparator = Regex.Escape(_timeSeparator);
            if (Regex.IsMatch(val, @"^[\d]{1,2}" + timeSeparator + @"[\d]{2}(" + timeSeparator + @"[\d]{2})?$", RegexOptions.Compiled))
            {
                var timeResult = _timeValueFunc.Execute(val);
                if (timeResult.DataType == DataType.Date)
                {
                    return timeResult;
                }
            }
            var dateResult = _dateValueFunc.Execute(val);
            if (dateResult.DataType == DataType.Date)
            {
                return dateResult;
            }
            return CreateResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);
        }
    }
}

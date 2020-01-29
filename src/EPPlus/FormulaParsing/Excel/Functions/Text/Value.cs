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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    internal class Value : ExcelFunction
    {
        private readonly string _groupSeparator = CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator;
        private readonly string _decimalSeparator = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
        private readonly string _timeSeparator = CultureInfo.CurrentCulture.DateTimeFormat.TimeSeparator;
        private readonly string _shortTimePattern = CultureInfo.CurrentCulture.DateTimeFormat.ShortTimePattern;
        private readonly DateValue _dateValueFunc = new DateValue();
        private readonly TimeValue _timeValueFunc = new TimeValue();

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var val = ArgToString(arguments, 0);
            double result = 0d;
            if (string.IsNullOrEmpty(val)) return CreateResult(result, DataType.Integer);
            val = val.TrimEnd(' ');
            if (Regex.IsMatch(val, $"^[\\d]*({Regex.Escape(_groupSeparator)}?[\\d]*)?({Regex.Escape(_decimalSeparator)}[\\d]*)*?[ ?% ?]?$"))
            {
                if (val.EndsWith("%"))
                {
                    val = val.TrimEnd('%');
                    result = double.Parse(val) / 100;
                }
                else
                {
                    result = double.Parse(val);
                }
                return CreateResult(result, DataType.Decimal);
            }
            if (double.TryParse(val, NumberStyles.Float, CultureInfo.CurrentCulture, out result))
            {
                return CreateResult(result, DataType.Decimal);
            }
            var timeSeparator = Regex.Escape(_timeSeparator);
            if (Regex.IsMatch(val, @"^[\d]{1,2}" + timeSeparator + @"[\d]{2}(" + timeSeparator + @"[\d]{2})?$"))
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

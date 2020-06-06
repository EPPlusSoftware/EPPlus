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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    /// <summary>
    /// Simple implementation of DateValue function, just using .NET built-in
    /// function System.DateTime.TryParse, based on current culture
    /// </summary>
    [FunctionMetadata(
        Category = ExcelFunctionCategory.DateAndTime,
        EPPlusVersion = "4",
        Description = "Converts a text string showing a date, to an integer that represents the date in Excel's date-time code")]
    internal class DateValue : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var dateString = ArgToString(arguments, 0);
            return Execute(dateString);
        }

        internal CompileResult Execute(string dateString)
        {
            System.DateTime result;
            System.DateTime.TryParse(dateString, out result);
            return result != System.DateTime.MinValue ?
                CreateResult(result.ToOADate(), DataType.Date) :
                CreateResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);
        }
    }
}

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
    [FunctionMetadata(
        Category = ExcelFunctionCategory.DateAndTime,
        EPPlusVersion = "4",
        Description = "Returns an integer representing the week number (from 1 to 53) of the year from a user-supplied date")]
    internal class Weeknum : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1, eErrorType.Value);
            var dateSerial = ArgToDecimal(arguments, 0);
            var date = System.DateTime.FromOADate(dateSerial);
            var startDay = DayOfWeek.Sunday;
            if (arguments.Count() > 1)
            {
                var argStartDay = ArgToInt(arguments, 1);
                switch (argStartDay)
                {
                    case 1:
                        startDay = DayOfWeek.Sunday;
                        break;
                    case 2:
                    case 11:
                        startDay = DayOfWeek.Monday;
                        break;
                    case 12:
                        startDay = DayOfWeek.Tuesday;
                        break;
                    case 13:
                        startDay = DayOfWeek.Wednesday;
                        break;
                    case 14:
                        startDay = DayOfWeek.Thursday;
                        break;
                    case 15:
                        startDay = DayOfWeek.Friday;
                        break;
                    case 16:
                        startDay = DayOfWeek.Saturday;
                        break;
                    default:
                        // Not supported 
                        ThrowExcelErrorValueException(eErrorType.Num);
                        break;
                }
            }
            if (DateTimeFormatInfo.CurrentInfo == null)
            {
                throw new InvalidOperationException(
                    "Could not execute Weeknum function because DateTimeFormatInfo.CurrentInfo was null");
            }
            var week = DateTimeFormatInfo.CurrentInfo.Calendar.GetWeekOfYear(date, CalendarWeekRule.FirstDay,
                                                                             startDay);
            return CreateResult(week, DataType.Integer);
        }
        
        
    }
}

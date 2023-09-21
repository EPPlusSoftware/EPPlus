/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/27/2020         EPPlus Software AB       EPPlus 5.5
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateAndTime
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.DateAndTime,
        EPPlusVersion = "5.5",
        Description = "Get days, months, or years between two dates",
        SupportsArrays = true)]
    internal class DateDif : DateParsingFunction
    {

        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.Custom;

        public override void ConfigureArrayBehaviour(ArrayBehaviourConfig config)
        {
            config.SetArrayParameterIndexes(0, 1, 2);
        }

        public override int ArgumentMinLength => 3;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var startDateObj = arguments[0].Value;
            var startDate = ParseDate(arguments, startDateObj, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);

            var endDateObj = arguments.ElementAt(1).Value;
            var endDate = ParseDate(arguments, endDateObj, 1, out ExcelErrorValue e2);
            if (e2 != null) return CompileResult.GetErrorResult(e2.Type);

            if (startDate > endDate) return CreateResult(eErrorType.Num);
            var unit = ArgToString(arguments, 2);
            switch(unit.ToLower())
            {
                case "y":
                    return CreateResult(DateDiffYears(startDate, endDate), DataType.Integer);
                case "m":
                    return CreateResult(DateDiffMonths(startDate, endDate), DataType.Integer);
                case "d":
                    var daysD = endDate.Subtract(startDate).TotalDays;
                    return CreateResult(daysD, DataType.Integer);
                case "ym":
                    var monthsYm = DateDiffMonthsY(startDate, endDate);
                    return CreateResult(monthsYm, DataType.Integer);
                case "yd":
                    var daysYd = GetStartYearEndDateY(startDate, endDate).Subtract(startDate).TotalDays;
                    return CreateResult(daysYd, DataType.Integer);
                case "md":
                    // NB! Excel calculates wrong here sometimes. Example DATEDIF(2001-04-02, 2003-01-01, "md") = 30 (it should be 29)
                    // we have not implemented this bug in EPPlus. Microsoft advices not to use the DateDif function due to this and other bugs.
                    var daysMd = GetStartYearEndDateMd(startDate, endDate).Subtract(startDate).TotalDays;
                    return CreateResult(daysMd, DataType.Integer);
                default:
                    return CompileResult.GetErrorResult(eErrorType.Num);
            }
        }

        private double DateDiffYears(DateTime start, DateTime end)
        {
            var result = Convert.ToDouble(end.Year - start.Year);
            var tmpEnd = GetStartYearEndDate(start, end);
            if (start > tmpEnd)
            {
                result -= 1;
            }
            return result;
        }

        private double DateDiffMonths(DateTime start, DateTime end)
        {
            var years = DateDiffYears(start, end);
            var result = years * 12;
            var tmpEnd = GetStartYearEndDate(start, end);
            if(start > tmpEnd)
            {
                result += 12;
                while (start > tmpEnd)
                {
                    tmpEnd = tmpEnd.AddMonths(1);
                    result--;
                }
            }
            
            return result;
        }

        private double DateDiffMonthsY(DateTime start, DateTime end)
        {
            var endDate = GetStartYearEndDateY(start, end);
            var nMonths = 0d;
            var tmpDate = start;
            if(tmpDate.AddMonths(1) < endDate)
            {
                do
                {
                    tmpDate = tmpDate.AddMonths(1);
                    if(tmpDate < endDate) nMonths++;
                }
                while (tmpDate < endDate);
            }
            
            return nMonths;
        }

        private DateTime GetStartYearEndDate(DateTime start, DateTime end)
        {
            return new DateTime(start.Year, end.Month, end.Day, end.Hour, end.Minute, end.Second, end.Millisecond);
        }

        private DateTime GetStartYearEndDateY(DateTime start, DateTime end)
        {
            var dt = new DateTime(start.Year, end.Month, end.Day, end.Hour, end.Minute, end.Second, end.Millisecond);
            if(dt < start)
            {
                dt = dt.AddYears(1);
            }
            return dt;
        }

        private DateTime GetStartYearEndDateMd(DateTime start, DateTime end)
        {
            var dt = new DateTime(start.Year, start.Month, end.Day, end.Hour, end.Minute, end.Second, end.Millisecond);
            if (dt < start)
            {
                dt = dt.AddMonths(1);
            }
            return dt;
        }
    }
}

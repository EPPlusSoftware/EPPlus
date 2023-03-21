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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.DateAndTime,
        EPPlusVersion = "5.5",
        Description = "Get days, months, or years between two dates",
        SupportsArrays = true)]
    internal class DateDif : DateParsingFunction
    {
        private readonly ArrayBehaviourConfig _arrayConfig = new ArrayBehaviourConfig
        {
            ArrayParameterIndexes = new List<int> { 0, 1, 2 }
        };

        internal override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.Custom;

        internal override ArrayBehaviourConfig GetArrayBehaviourConfig()
        {
            return _arrayConfig;
        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            var startDateObj = arguments.ElementAt(0).Value;
            var startDate = ParseDate(arguments, startDateObj);
            var endDateObj = arguments.ElementAt(1).Value;
            var endDate = ParseDate(arguments, endDateObj, 1);
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
                    return CreateResult(eErrorType.Num);
            }
        }

        private double DateDiffYears(System.DateTime start, System.DateTime end)
        {
            var result = Convert.ToDouble(end.Year - start.Year);
            var tmpEnd = GetStartYearEndDate(start, end);
            if (start > tmpEnd)
            {
                result -= 1;
            }
            return result;
        }

        private double DateDiffMonths(System.DateTime start, System.DateTime end)
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

        private double DateDiffMonthsY(System.DateTime start, System.DateTime end)
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

        private System.DateTime GetStartYearEndDate(System.DateTime start, System.DateTime end)
        {
            return new System.DateTime(start.Year, end.Month, end.Day, end.Hour, end.Minute, end.Second, end.Millisecond);
        }

        private System.DateTime GetStartYearEndDateY(System.DateTime start, System.DateTime end)
        {
            var dt = new System.DateTime(start.Year, end.Month, end.Day, end.Hour, end.Minute, end.Second, end.Millisecond);
            if(dt < start)
            {
                dt = dt.AddYears(1);
            }
            return dt;
        }

        private System.DateTime GetStartYearEndDateMd(System.DateTime start, System.DateTime end)
        {
            var dt = new System.DateTime(start.Year, start.Month, end.Day, end.Hour, end.Minute, end.Second, end.Millisecond);
            if (dt < start)
            {
                dt = dt.AddMonths(1);
            }
            return dt;
        }
    }
}

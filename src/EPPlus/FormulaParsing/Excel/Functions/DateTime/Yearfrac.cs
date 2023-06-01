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
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.DateAndTime,
        EPPlusVersion = "4",
        Description = "Calculates the fraction of the year represented by the number of whole days between two dates")]
    internal class Yearfrac : ExcelFunction
    {
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var date1Num = ArgToDecimal(arguments, 0);
            var date2Num = ArgToDecimal(arguments, 1);
            if (date1Num > date2Num) //Switch to make date1 the lowest date
            {
                var t = date1Num;
                date1Num = date2Num;
                date2Num = t;
                var fa = arguments[1];
                arguments[1] = arguments[0];
                arguments[0] = fa;
            }
            var date1 = System.DateTime.FromOADate(date1Num);
            var date2 = System.DateTime.FromOADate(date2Num);

            var basis = 0;
            if (arguments.Count > 2)
            {
                basis = ArgToInt(arguments, 2);
                if (basis < 0 || basis > 4) return CreateResult(eErrorType.Num);
            }
            var func = context.Configuration.FunctionRepository.GetFunction("days360");
            var calendar = new GregorianCalendar();
            switch (basis)
            {
                case 0:
                    var d360Result = System.Math.Abs(func.Execute(arguments, context).ResultNumeric);
                    // reproducing excels behaviour
                    if (date1.Month == 2 && date2.Day==31)
                    {
                        var daysInFeb = calendar.IsLeapYear(date1.Year) ? 29 : 28;
                        if (date1.Day == daysInFeb) d360Result++;  
                    }
                    return CreateResult(d360Result / 360d, DataType.Decimal);
                case 1:
                    return CreateResult(System.Math.Abs((date2 - date1).TotalDays / CalculateAcutalYear(date1, date2)), DataType.Decimal);
                case 2:
                    return CreateResult(System.Math.Abs((date2 - date1).TotalDays / 360d), DataType.Decimal);
                case 3:
                    return CreateResult(System.Math.Abs((date2 - date1).TotalDays / 365d), DataType.Decimal);
                case 4:
                    arguments.Add(new FunctionArgument(true));
                    double? result = System.Math.Abs(func.Execute(arguments, context).ResultNumeric / 360d);
                    return CreateResult(result.Value, DataType.Decimal);
                default:
                    return null;
            }
        }

        private double CalculateAcutalYear(System.DateTime dt1, System.DateTime dt2)
        {
            var calendar = new GregorianCalendar();
            var perYear = 0d;
            var nYears = dt2.Year - dt1.Year + 1;
            for (var y = dt1.Year; y <= dt2.Year; ++y)
            {
                perYear += calendar.IsLeapYear(y) ? 366 : 365;
            }
            if (new System.DateTime(dt1.Year + 1, dt1.Month, dt1.Day) >= dt2)
            {
                nYears = 1;
                perYear = 365;
                if (calendar.IsLeapYear(dt1.Year) && dt1.Month <= 2)
                    perYear = 366;
                else if (calendar.IsLeapYear(dt2.Year) && dt2.Month > 2)
                    perYear = 366;
                else if (dt2.Month == 2 && dt2.Day == 29)
                    perYear = 366;
            }
            return perYear/(double) nYears;  
        }
    }
}

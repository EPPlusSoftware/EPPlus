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
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateAndTime
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.DateAndTime,
        EPPlusVersion = "4",
        Description = "Returns the ISO week number of the year for a given date",
        IntroducedInExcelVersion = "2013",
        SupportsArrays = true)]
    internal class IsoWeekNum : ExcelFunction
    {

        public override string NamespacePrefix => "_xlfn.";
        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.FirstArgCouldBeARange;
        public override int ArgumentMinLength => 1;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var dateInt = ArgToInt(arguments, 0);
            var date = DateTime.FromOADate(dateInt);
            return CreateResult(WeekNumber(date), DataType.Integer);
        }

        /// <summary>
        /// This implementation was found on http://stackoverflow.com/questions/1285191/get-week-of-date-from-linq-query
        /// </summary>
        /// <param name="fromDate"></param>
        /// <returns></returns>
        private int WeekNumber(DateTime fromDate)
        {
            // Get jan 1st of the year
            var startOfYear = fromDate.AddDays(-fromDate.Day + 1).AddMonths(-fromDate.Month + 1);
            // Get dec 31st of the year
            var endOfYear = startOfYear.AddYears(1).AddDays(-1);
            // ISO 8601 weeks start with Monday
            // The first week of a year includes the first Thursday
            // DayOfWeek returns 0 for sunday up to 6 for saterday
            int[] iso8601Correction = { 6, 7, 8, 9, 10, 4, 5 };
            int nds = fromDate.Subtract(startOfYear).Days + iso8601Correction[(int)startOfYear.DayOfWeek];
            int wk = nds / 7;
            switch (wk)
            {
                case 0:
                    // Return weeknumber of dec 31st of the previous year
                    return WeekNumber(startOfYear.AddDays(-1));
                case 53:
                    // If dec 31st falls before thursday it is week 01 of next year
                    if (endOfYear.DayOfWeek < DayOfWeek.Thursday)
                        return 1;
                    return wk;
                default: return wk;
            }
        }
    }
}

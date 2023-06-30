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
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateAndTime.Workdays;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateAndTime
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.DateAndTime,
        EPPlusVersion = "4",
        Description = "Returns the number of whole networkdays (excluding weekends & holidays), between two supplied dates")]
    internal class Networkdays : ExcelFunction
    {
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var startDate = DateTime.FromOADate(ArgToInt(arguments, 0));
            var endDate = DateTime.FromOADate(ArgToInt(arguments, 1));
            var calculator = new WorkdayCalculator();
            var result = calculator.CalculateNumberOfWorkdays(startDate, endDate);
            if (arguments.Count > 2)
            {
                result = calculator.ReduceWorkdaysWithHolidays(result, arguments[2]);
            }
            
            return new CompileResult(result.NumberOfWorkdays, DataType.Integer);
        }
    }
}

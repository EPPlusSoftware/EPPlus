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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.DateAndTime,
        EPPlusVersion = "4",
        Description = "Returns a date, from a user-supplied year, month and day")]
    internal class Date : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            var year = ArgToInt(arguments, 0);
            var month = ArgToInt(arguments, 1);
            var day = ArgToInt(arguments, 2);
            var date = new System.DateTime(year, 1, 1);
            month -= 1;
            date = date.AddMonths(month);
            date = date.AddDays((double)(day - 1));
            var ret = date.ToOADate();
            if (ret > 0 && ret < 60) ret--;
            return CreateResult(ret, DataType.Date);
        }
    }
}

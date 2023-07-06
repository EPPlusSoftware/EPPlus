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
        Description = "Returns a time, from a user-supplied hour, minute and second")]
    internal class Time : TimeBaseFunction
    {
        public Time()
            : base()
        {

        }

        public override int ArgumentMinLength => 3;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var hour = ArgToInt(arguments, 0);
            var min = ArgToInt(arguments, 1);
            var sec = ArgToInt(arguments, 2);

            if (sec < 0 || sec > 59) return CompileResult.GetErrorResult(eErrorType.Value);
            if (min < 0 || min > 59) return CompileResult.GetErrorResult(eErrorType.Value);
            if (min < 0 || hour > 23) return CompileResult.GetErrorResult(eErrorType.Value);


            var secondsOfThisTime = (double)(hour * 60 * 60 + min * 60 + sec);
            return CreateResult(GetTimeSerialNumber(secondsOfThisTime), DataType.Time);
        }
    }
}

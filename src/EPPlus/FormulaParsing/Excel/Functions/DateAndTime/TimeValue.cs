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
    /// <summary>
    /// Simple implementation of TimeValue function, just using .NET built-in
    /// function System.DateTime.TryParse, based on current culture
    /// </summary>
    [FunctionMetadata(
        Category = ExcelFunctionCategory.DateAndTime,
        EPPlusVersion = "4",
        Description = "Converts a text string showing a time, to a decimal that represents the time in Excel",
        SupportsArrays = true)]
    internal class TimeValue : ExcelFunction
    {
        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.FirstArgCouldBeARange;

        public override int ArgumentMinLength => 1;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var dateString = ArgToString(arguments, 0);
            return Execute(dateString);
        }

        internal CompileResult Execute(string dateString)
        {
            DateTime result;
            DateTime.TryParse(dateString, out result);
            return result != DateTime.MinValue ?
                CreateResult(GetTimeValue(result), DataType.Date) :
                CompileResult.GetErrorResult(eErrorType.Value);
        }

        private double GetTimeValue(DateTime result)
        {
            return (int)result.TimeOfDay.TotalSeconds == 0 ? 0d : result.TimeOfDay.TotalSeconds/ (3600 * 24);
        }
    }
}

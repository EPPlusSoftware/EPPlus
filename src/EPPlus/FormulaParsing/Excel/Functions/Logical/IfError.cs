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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Logical
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Logical,
        EPPlusVersion = "4",
        Description = "Tests if an initial supplied value (or expression) returns an error, and if so, returns a supplied value; Otherwise the function returns the initial value.")]
    internal class IfError : ExcelFunction
    {
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var arg1 = arguments[0];
            var arg2 = arguments[1];
            if (arg1.ValueIsExcelError)
            {
                return GetResultByObject(arg2.Value);
            }
            return GetResultByObject(arg1.Value);
        }
        public override FunctionParameterInformation GetParameterInfo(int argumentIndex)
        {
            return FunctionParameterInformation.IgnoreErrorInPreExecute;
        }
    }
}

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
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Information,
        EPPlusVersion = "7",
        Description = "Tests if the supplied value is a reference, and if so, returns TRUE; Otherwise, returns FALSE")]
    internal class IsRef : ExcelFunction
    {
        public override int ArgumentMinLength => 1;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (arguments[0].DataType==DataType.ExcelRange)
            {
                var ri = arguments[0].ValueAsRangeInfo;
                if(ri != null && ri.IsInMemoryRange == false)
                {
                    return CompileResult.True; 
                }
            }
            return CompileResult.False;
        }
        public override ExcelFunctionParametersInfo ParametersInfo => new ExcelFunctionParametersInfo(new Func<int, FunctionParameterInformation>((argumentIndex) =>
        {
            return FunctionParameterInformation.IgnoreErrorInPreExecute | FunctionParameterInformation.IgnoreAddress;
        }));
    }
}

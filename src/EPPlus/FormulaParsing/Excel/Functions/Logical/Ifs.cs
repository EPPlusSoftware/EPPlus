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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Logical
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Logical,
        EPPlusVersion = "5.0",
        Description = "Returns the largest numeric value that meets one or more criteria in a range of values",
        IntroducedInExcelVersion = "2019")]
    internal class Ifs : ExcelFunction
    {
        public override string NamespacePrefix => "_xlfn.";
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var maxArgs = arguments.Count < 254 ? arguments.Count : 254; 
            if(maxArgs % 2 != 0) 
            {
                return CompileResult.GetErrorResult(eErrorType.Value);
            }
            for(var x = 0; x < maxArgs; x += 2)
            {
                var argResult = ArgToDecimal(arguments, x, out ExcelErrorValue e1);
                if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
                if (System.Math.Round(argResult, 15) != 0d)
                {
                    var arg = arguments.ElementAt(x + 1);
                    if(arg.DataType==DataType.ExcelRange)
                    {
                        return CompileResultFactory.Create(arg, arg.ValueAsRangeInfo.Address);
                    }
                    else
                    {
                        return CompileResultFactory.Create(arg.Value);
                    }
                }
            }
            return CompileResult.GetErrorResult(eErrorType.NA);
        }
        public override bool ReturnsReference => true;
        public override ExcelFunctionParametersInfo ParametersInfo => new ExcelFunctionParametersInfo(new Func<int, FunctionParameterInformation>((argumentIndex) =>
        {
            if (argumentIndex % 2 == 0)
            {
                return FunctionParameterInformation.Condition | FunctionParameterInformation.IgnoreErrorInPreExecute;
            }
            else
            {
                return FunctionParameterInformation.UseIfConditionIsTrue | FunctionParameterInformation.IgnoreErrorInPreExecute;
            }
        }));
		/// <summary>
		/// If the function is allowed in a pivot table calculated field
		/// </summary>
		public override bool IsAllowedInCalculatedPivotTableField => false;
	}
}

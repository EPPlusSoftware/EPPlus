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
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Information,
        EPPlusVersion = "4",
        Description = "Tests if an initial supplied value (or expression) returns an error and if so, returns TRUE; Otherwise returns FALSE",
        SupportsArrays = true)]
    internal class IsError : ExcelFunction
    {
        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.FirstArgCouldBeARange;
        public override int ArgumentMinLength => 1;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            foreach (var argument in arguments)
            {
                if (argument.Value is IRangeInfo)
                {
                    var r = (IRangeInfo)argument.Value;
                    if (ExcelErrorValue.Values.IsErrorValue(r.GetValue(r.Address.FromRow, r.Address.FromCol)))
                    {
                        return CreateResult(true, DataType.Boolean);
                    }
                }
                else
                {
                    if (ExcelErrorValue.Values.IsErrorValue(argument.Value))
                    {
                        return CreateResult(true, DataType.Boolean);
                    }
                }                
            }
            return CreateResult(false, DataType.Boolean);
        }

        //public override CompileResult HandleError(string errorCode)
        //{
        //    return CreateResult(true, DataType.Boolean);
        //}
        public override ExcelFunctionParametersInfo ParametersInfo => new ExcelFunctionParametersInfo(new Func<int, FunctionParameterInformation>((argumentIndex) =>
        {
            return FunctionParameterInformation.IgnoreErrorInPreExecute;
        }));
    }
}

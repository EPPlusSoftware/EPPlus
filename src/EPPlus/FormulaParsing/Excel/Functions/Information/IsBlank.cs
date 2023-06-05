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
        EPPlusVersion = "4",
        Description = "Tests if a supplied cell is blank (empty), and if so, returns TRUE; Otherwise, returns FALSE",
        SupportsArrays = true)]
    internal class IsBlank : ExcelFunction
    {
        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.FirstArgCouldBeARange;

        public override int ArgumentMinLength => 0;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (arguments == null || arguments.Count == 0)
            {
                return CreateResult(true, DataType.Boolean);
            }
            var result = true;
            foreach (var arg in arguments)
            {
                if (arg.Value is IRangeInfo)
                {                    
                    var r=(IRangeInfo)arg.Value;
                    if (r.GetValue(r.Address.FromRow, r.Address.FromCol) != null)
                    {
                        result = false;
                    }
                }
                else
                {
                    if (arg.Value != null && (arg.Value.ToString() != string.Empty))
                    {
                        result = false;
                        break;
                    }
                }
            }
            return CreateResult(result, DataType.Boolean);
        }
        public override FunctionParameterInformation GetParameterInfo(int argumentIndex)
        {
            return FunctionParameterInformation.IgnoreErrorInPreExecute;
        }
    }
}

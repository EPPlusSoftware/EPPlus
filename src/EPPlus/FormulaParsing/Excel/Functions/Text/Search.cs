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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Text,
        EPPlusVersion = "4",
        Description = "Returns the position of a supplied character or text string from within a supplied text string (non-case-sensitive)",
        SupportsArrays = true)]
    internal class Search : ExcelFunction
    {
        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.Custom;
        public override void ConfigureArrayBehaviour(ArrayBehaviourConfig config)
        {
            config.SetArrayParameterIndexes(0, 1);
        }

        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var search = ArgToString(arguments, 0);
            var searchIn = ArgToString(arguments, 1);
            var startIndex = 0;
            if (arguments.Count > 2)
            {
                startIndex = ArgToInt(arguments, 2, out ExcelErrorValue e2) - 1;
                if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
            }
            var result = searchIn.IndexOf(search, startIndex, System.StringComparison.OrdinalIgnoreCase);
            if (result == -1)
            {
                return CompileResult.GetErrorResult(eErrorType.Value);
            }
            // Adding 1 because Excel uses 1-based index
            return CreateResult(result + 1, DataType.Integer);
        }
    }
}

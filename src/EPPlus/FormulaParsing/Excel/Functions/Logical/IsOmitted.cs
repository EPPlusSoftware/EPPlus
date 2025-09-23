/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/04/2025        EPPlus Software AB       Initial release EPPlus 8.1
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
        EPPlusVersion = "8.1",
        Description = "Applies a LAMBDA to each column and returns an array of the results. For example, if the original array is 3 columns by 2 rows, the returned array is 3 columns by 1 row.",
        IntroducedInExcelVersion = "2021")]
    internal class IsOmitted : ExcelFunction
    {
        public override int ArgumentMinLength => 1;

        public override string NamespacePrefix => "_xlfn.";

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (arguments[0].DataType == DataType.Empty)
            {
                if (arguments[0].Value != null && arguments[0].Value.ToString() == "oovs") return CreateResult(false, DataType.Boolean);
                return CreateDynamicArrayResult(true, DataType.Boolean);
            }
            return CreateDynamicArrayResult(false, DataType.Boolean);
        }
    }
}

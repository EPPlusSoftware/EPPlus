/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  XX/XX/XXXX         EPPlus Software AB           EPPlus vX
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    [FunctionMetadata(
       Category = ExcelFunctionCategory.Text,
       EPPlusVersion = "8",
       Description = "Returns the numeric code for the first character of a text string.")]
    internal class CodeFunction : ExcelFunction
    {
        public override int ArgumentMinLength => 1;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var text = ArgToString(arguments, 0);
            if (string.IsNullOrEmpty(text))
            {
                return CompileResult.GetErrorResult(eErrorType.Value);
            }
            // Return the numeric code (Unicode code unit) of the first character.
            // For surrogate pairs we return the leading high surrogate, matching Excel's behavior.
            int code = text[0];
            return CreateResult((double)code, DataType.Decimal);
        }
    }
}
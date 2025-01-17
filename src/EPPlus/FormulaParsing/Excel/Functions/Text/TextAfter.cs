/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/15/2024         EPPlus Software AB       Initial release EPPlus 7.2
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    [FunctionMetadata(
           Category = ExcelFunctionCategory.Text,
           EPPlusVersion = "7.2",
           Description = "Get the text after delimiter",
           SupportsArrays = false)]
    internal class TextAfter : TextDelimiterFunctionBase
    {
        public TextAfter(DelimiterFunction funcType) : base(funcType)
        {
        }

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            return ProcessText(arguments);
        }
    }
}

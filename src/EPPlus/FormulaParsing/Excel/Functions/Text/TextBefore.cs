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
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExcelUtilities.TextUtils;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Text,
        EPPlusVersion = "7.2",
        Description = "Get the text before delimiter",
        SupportsArrays = true)]
    internal class TextBefore : TextDelimiterFunctionBase
    {


        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var text = ArgToString(arguments, 0);
            var delimiters = GetDelimiters(arguments);
            var instanceNum = GetInstanceNum(arguments);
            var matchMode = GetMatchMode(arguments, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            var matchEnd = GetMatchEnd(arguments, out ExcelErrorValue e2);
            if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
            string ifNotFound = GetIfNotFound(arguments);
            
            return ProcessTextBefore(text, delimiters, instanceNum, matchMode, matchEnd, ifNotFound);
        }
    }
}

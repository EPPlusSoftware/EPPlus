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
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    [FunctionMetadata(
       Category = ExcelFunctionCategory.Text,
       EPPlusVersion = "8.6",
       Description = "Returns a URL-encoded string, replacing characters that are not allowed in URLs with their percent-encoded equivalents.", 
       SupportsArrays = true)]
    internal class EncodeUrl : ExcelFunction
    {
        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.FirstArgCouldBeARange;

        public override string NamespacePrefix => "_xlfn.";
        public override int ArgumentMinLength => 1;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var text = ArgToString(arguments, 0);
            if (text == null) text = string.Empty;
            // Uri.EscapeDataString uses UTF-8 percent-encoding, matching Excel's behavior.
            var encoded = Uri.EscapeDataString(text);
            return CreateResult(encoded, DataType.String);
        }
    }
}
/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  25/5/2026         EPPlus Software AB           EPPlus v8.6
 *************************************************************************************************/

using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.TrimFunctions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
     Category = ExcelFunctionCategory.LookupAndReference,
     EPPlusVersion = "8",
     Description = "Trims all trailing empty cells in a given range")]
    internal class TroTrailing : TrimFunctionsBase
    {
        public override int ArgumentMinLength => 1;

        public override string NamespacePrefix => "_xlfn.";

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
    => ExecuteTrim(arguments, TrimMode.Trailing, TrimMode.Trailing);       
    }
}

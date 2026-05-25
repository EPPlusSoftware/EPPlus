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
    Description = "Excludes all empty rows and/or columns from the outer edges of a range")]
    internal class TrimRange : TrimFunctionsBase
    {

        public override int ArgumentMinLength => 1;
        public override string NamespacePrefix => "_xlfn.";

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
        var rowMode = TrimMode.Both;
        var colMode = TrimMode.Both;            

        if (arguments.Count > 1)
        {
            var v = ArgToInt(arguments, 1, RoundingMethod.Convert);
            if (v < 0 || v > 3) return CompileResult.GetErrorResult(eErrorType.Value);
            rowMode = (TrimMode)v;
        }

        if (arguments.Count > 2)
        {
            var v = ArgToInt(arguments, 2, RoundingMethod.Convert);
                if (v < 0 || v > 3) return CompileResult.GetErrorResult(eErrorType.Value);
            colMode = (TrimMode)v;
        }

        return ExecuteTrim(arguments, rowMode, colMode);
        }
    }
}

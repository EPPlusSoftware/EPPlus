/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/25/2026         EPPlus Software AB       Initial release EPPlus 8.6
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Information,
        EPPlusVersion = "8.6",
        Description = "Returns TRUE if the reference is to a cell containing a formula.")]
    internal class IsFormula : ExcelFunction
    {
        public override int ArgumentMinLength => 1;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var reference = arguments[0].ValueAsRangeInfo;
            var ws = reference.Worksheet;
            var adr = reference.Address;
            if(adr.IsSingleCell)
            {
                return CreateResult(ws.HasFormula(adr.FromRow, adr.FromCol), DataType.Boolean);
            }
            var returnRange = new InMemoryRange(reference);
            for(var col = adr.FromCol; col <= adr.ToCol; col++)
            {
                for (var row = adr.FromRow; row <= adr.ToRow; row++)
                {
                    var retVal = ws.HasFormula(row, col);
                    returnRange.SetValue(row - adr.FromRow, col - adr.FromCol, retVal);
                }
            }
            return CreateDynamicArrayResult(returnRange, DataType.ExcelRange);
        }
    }
}

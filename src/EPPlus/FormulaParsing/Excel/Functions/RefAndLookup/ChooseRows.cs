/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  22/3/2023         EPPlus Software AB           EPPlus v7
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.LookupAndReference,
        EPPlusVersion = "7",
        Description = "Returns the specified rows from an array.")]
    internal class ChooseRows : ExcelFunction
    {
        public override string NamespacePrefix => "_xlfn.";
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var firstArg = arguments[0];
            var rows = new List<int>();
            for (var x = 1; x < arguments.Count; x++)
            {
                var r = ArgToInt(arguments, x, out ExcelErrorValue e1);
                if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
                rows.Add(r);
            }
            if (firstArg.IsExcelRange)
            {
                var source = firstArg.ValueAsRangeInfo;
                if(rows.Any(r => Math.Abs(r - 1) > source.Size.NumberOfRows || r == 0))
                {
                    return CompileResult.GetDynamicArrayResultError(eErrorType.Value);
                }
                var nCols = source.Size.NumberOfCols;
                var resultRange = new InMemoryRange(new RangeDefinition(rows.Count, source.Size.NumberOfCols));
                var rIx = 0;
                foreach (var row in rows)
                {
                    for (var col = 0; col < nCols; col++)
                    {
                        var sourceIx = row > 0 ? row - 1 : source.Size.NumberOfRows + row;

                        var val = source.GetOffset(sourceIx, col);
                        resultRange.SetValue(rIx, col, val);
                    }
                    rIx++;
                }
                return CreateDynamicArrayResult(resultRange, DataType.ExcelRange);
            }
            else if (!rows.Any(x => x > 1))
            {
                var resultRange = new InMemoryRange(new RangeDefinition(rows.Count, 1));
                var rIx = 0;
                foreach (var row in rows)
                {
                    resultRange.SetValue(rIx++, 0, firstArg.Value);
                }
                return CreateDynamicArrayResult(resultRange, DataType.ExcelRange);
            }
            return CompileResult.GetDynamicArrayResultError(eErrorType.Value);
        }
		/// <summary>
		/// If the function is allowed in a pivot table calculated field
		/// </summary>
		public override bool IsAllowedInCalculatedPivotTableField => false;
	}
}

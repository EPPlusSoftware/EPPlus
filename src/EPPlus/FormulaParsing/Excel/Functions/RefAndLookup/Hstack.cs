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
        Description = "Combines arrays horizontally into a single array.",
        SupportsArrays = true)]
    internal class Hstack : ExcelFunction
    {
        public override string NamespacePrefix => "_xlfn.";

        public override int ArgumentMinLength => 1;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var ranges = new List<IRangeInfo>();
            foreach (var arg in arguments)
            {
                if (!arg.IsExcelRange)
                {
                    var rng = new InMemoryRange(1, 1);
                    rng.SetValue(0, 0, arg.Value);
                    ranges.Add(rng);
                }
                else
                {
                    ranges.Add(arg.ValueAsRangeInfo);
                }
            }
            var nRows = ranges.Max(x => x.Size.NumberOfRows);
            var nCols = Convert.ToInt16(ranges.Sum(x => x.Size.NumberOfCols));
            var range = new InMemoryRange(nRows, nCols);
            var colIx = 0;
            foreach (var argRange in ranges)
            {
                for (var col = 0; col < argRange.Size.NumberOfCols; col++)
                {
                    var row = 0;
                    for (; row < argRange.Size.NumberOfRows; row++)
                    {
                        var v = argRange.GetOffset(row, col);
                        range.SetValue(row, colIx, v);
                    }
                    for (; row < range.Size.NumberOfRows; row++)
                    {
                        range.SetValue(row, colIx, ErrorValues.NAError);
                    }
                    colIx++;
                }
            }
            return CreateResult(range, DataType.ExcelRange);
        }
		/// <summary>
		/// If the function is allowed in a pivot table calculated field
		/// </summary>
		public override bool IsAllowedInCalculatedPivotTableField => false;
	}
}

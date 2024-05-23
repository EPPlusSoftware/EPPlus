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
        Description = "Returns the array in a single column.",
        SupportsArrays = true)]
    internal class Expand : ExcelFunction
    {
        public override string NamespacePrefix => "_xlfn.";
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var firstArg = arguments[0];
            IRangeInfo range;
            if (!firstArg.IsExcelRange)
            {
                var imr = new InMemoryRange(new RangeDefinition(1, 1));
                imr.SetValue(0, 0, firstArg.Value);
                range = imr;
            }
            else
            {
                range = firstArg.ValueAsRangeInfo;
            }
            var rows = range.Size.NumberOfRows;
            var secondArg = arguments[1];
            if(secondArg.Value != null)
            {
                rows = ArgToInt(arguments, 1, out ExcelErrorValue e1);
                if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
                if(rows < range.Size.NumberOfRows)
                {
                    return CompileResult.GetDynamicArrayResultError(eErrorType.Value);
                }
            }
            var cols = range.Size.NumberOfCols;
            if(arguments.Count > 2 && arguments[2] != null)
            {
                cols = (short)ArgToInt(arguments, 2, out ExcelErrorValue e2);
                if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
                if(cols < range.Size.NumberOfCols)
                {
                    return CompileResult.GetDynamicArrayResultError(eErrorType.Value);
                }
                else if(cols + context.CurrentCell.Column > ExcelPackage.MaxColumns
                    ||
                    rows + context.CurrentCell.Row > ExcelPackage.MaxRows)
                {
                    return new DynamicArrayCompileResult(new ExcelRichDataErrorValue(0, 0));
                }
            }
            object padWith = ErrorValues.NAError;
            if(arguments.Count > 3 && arguments[3] != null)
            {
                padWith = arguments[3].Value;
            }
            // create return range
            var rr = new InMemoryRange(new RangeDefinition(rows, cols));
            for(var row = 0; row < rows; row++)
            {
                for(short col = 0; col < cols; col++)
                {
                    if (row < range.Size.NumberOfRows && col < range.Size.NumberOfCols)
                    {
                        var v = range.GetOffset(row, col);
                        if (v == null) v = 0;
                        rr.SetValue(row, col, v);
                    }
                    else
                    {
                        rr.SetValue(row, col, padWith);
                    }
                }
            }
            return CreateDynamicArrayResult(rr, DataType.ExcelRange);
        }
		/// <summary>
		/// If the function is allowed in a pivot table calculated field
		/// </summary>
		public override bool IsAllowedInCalculatedPivotTableField => false;
	}
}

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
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
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
        Description = "Excludes a specified number of rows or columns from the start or end of an array")]
    internal class Drop : ExcelFunction
    {
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var firstArg = arguments[0];
            int rows, cols;
            rows = ArgToInt(arguments, 1, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            if (arguments.Count > 2)
            {
                cols = ArgToInt(arguments, 2, out ExcelErrorValue e2);
                if(e2 != null) return CompileResult.GetErrorResult(e2.Type);
            }
            else
            {
                cols = 0;
            }

            if(firstArg.DataType == DataType.ExcelRange)
            {
                var r = firstArg.Value as IRangeInfo;
                if(r.Size.NumberOfRows <= Math.Abs(rows) || r.Size.NumberOfCols <= Math.Abs(cols))
                {
                    return CompileResult.GetErrorResult(eErrorType.Calc);
                }

                int fromRow, fromCol, toRow, toCol;

                if(rows<0)
                {
                    fromRow = r.Address.FromRow;
                    toRow = r.Address.ToRow+rows;
                }
                else
                {
                    fromRow = r.Address.FromRow + rows;
                    toRow = r.Address.ToRow;
                }

                if(cols<0)
                {
                    fromCol = r.Address.FromCol;
                    toCol = r.Address.ToCol + cols;
                }
                else
                {
                    fromCol = r.Address.FromCol + cols;
                    toCol = r.Address.ToCol;
                }

                
                IRangeInfo retRange;
                if(r.IsInMemoryRange)
                {
                    retRange = r.GetOffset(fromRow, fromCol, toRow, toCol);
                    return CreateDynamicArrayResult(retRange, DataType.ExcelRange);
                }
                else
                {
                    var address = new FormulaRangeAddress(context, r.Worksheet.IndexInList, fromRow, fromCol, toRow, toCol);
                    retRange = new RangeInfo(r.Worksheet, fromRow, fromCol, toRow, toCol, context, r.Address.ExternalReferenceIx); //External references must be check how they work.
                    return CreateDynamicArrayResult(retRange, DataType.ExcelRange, address);
                }
            }
            // arg was not a range
            if(rows == 0 && cols == 0)
            {                
                return CompileResultFactory.CreateDynamicArray(firstArg.Value);
            }
            return CompileResult.GetDynamicArrayResultError(eErrorType.Calc);
            
        }
        public override string NamespacePrefix => "_xlfn.";
        public override bool ReturnsReference => true;
		/// <summary>
		/// If the function is allowed in a pivot table calculated field
		/// </summary>
		public override bool IsAllowedInCalculatedPivotTableField => false;
	}
}

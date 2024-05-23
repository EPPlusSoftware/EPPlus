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
        Description = "Returns a specified number of contiguous rows or columns from the start or end of an array.")]
    internal class Take : ExcelFunction
    {
        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var firstArg = arguments[0];
            int? rowsParam = default;
            int? colsParam = default;
            int rows = 0, cols = 0;
            var rowsArg = arguments[1];
            if(rowsArg.Value != null) 
            {
                rowsParam = ArgToInt(arguments, 1, out ExcelErrorValue e1);
                if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            }
            
            if (arguments.Count > 2)
            {
                if(arguments[2].Value != null)
                {
                    colsParam = ArgToInt(arguments, 2, out ExcelErrorValue e2);
                    if (e2 != null) return CompileResult.GetErrorResult(e2.Type);
                }
            }

            if (firstArg.DataType == DataType.ExcelRange)
            {
                var r = firstArg.Value as IRangeInfo;
                rows = rowsParam ?? r.Size.NumberOfRows;
                cols = colsParam ?? r.Size.NumberOfCols;
                rows = rows > r.Size.NumberOfRows ? r.Size.NumberOfRows : rows;
                cols = cols > r.Size.NumberOfCols ? r.Size.NumberOfCols : cols;
                if (rows == 0 || cols == 0) return CompileResult.GetDynamicArrayResultError(eErrorType.Calc);
                if (r.Size.NumberOfRows < Math.Abs(rows) || r.Size.NumberOfCols < Math.Abs(cols))
                {
                    return CompileResult.GetDynamicArrayResultError(eErrorType.Calc);
                }

                int fromRow, fromCol, toRow, toCol;

                if (rows > 0)
                {
                    fromRow = r.Address.FromRow;
                    toRow = r.Address.FromRow + rows - 1;
                }
                else
                {
                    fromRow = r.Address.ToRow - Math.Abs(rows) + 1;
                    toRow = r.Address.ToRow;
                }

                if (cols > 0)
                {
                    fromCol = r.Address.FromCol;
                    toCol = r.Address.FromCol + cols - 1;
                }
                else
                {
                    fromCol = r.Address.ToCol - Math.Abs(cols) + 1;
                    toCol = r.Address.ToCol;
                }


                IRangeInfo retRange;
                if (r.IsInMemoryRange)
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
            if (rows != 0 && cols != 0)
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

using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    internal class TroTrailing : ExcelFunction
    {
        public override int ArgumentMinLength => 1;

        public override string NamespacePrefix => "_xlfn.";

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var range = arguments[0].ValueAsRangeInfo;
            var result = TrimRangeTrailing(range, out var error);
            if (error != null)
                return error;
            return CreateDynamicArrayResult(result, DataType.ExcelRange);
        }

        private object TrimRangeTrailing(IRangeInfo range, out CompileResult error)
        {
            int toRow = range.Size.NumberOfRows;
            int toCol = range.Size.NumberOfCols;
            error = null;

            // Hitta sista kolumnen med ett värde (från höger)
            for (int col = range.Size.NumberOfCols - 1; col >= 0; col--)
            {
                bool colHasValue = false;
                for (int row = 0; row < range.Size.NumberOfRows; row++)
                {
                    if (HasValue(range.GetOffset(row, col)))
                    {
                        colHasValue = true;
                        break;
                    }
                }
                if (colHasValue)
                    break;
                toCol--;
            }

            // Hitta sista raden med ett värde (nerifrån)
            for (int row = range.Size.NumberOfRows - 1; row >= 0; row--)
            {
                bool rowHasValue = false;
                for (int col = 0; col < range.Size.NumberOfCols; col++)
                {
                    if (HasValue(range.GetOffset(row, col)))
                    {
                        rowHasValue = true;
                        break;
                    }
                }
                if (rowHasValue)
                    break;
                toRow--;
            }

            if (toRow == 0 || toCol == 0)
            {
                error = CompileResult.GetErrorResult(eErrorType.Ref);
                return new InMemoryRange(1, 1);
            }

            var result = new InMemoryRange(toRow, (short)toCol);
            for (int r = 0; r < toRow; r++)
                for (int c = 0; c < toCol; c++)
                    result.SetValue(r, c, range.GetOffset(r, c));

            return result;
        }        

        private bool HasValue(object val)
        {
            if (val == null) return false;
            if (val.GetType().Name == "RowInternal") return false;
            return true;
        }
    }
}

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
            var toRow = range.Size.NumberOfCols;
            var tocol = range.Size.NumberOfRows;
            error = null;

            for (int col = range.Size.NumberOfCols; col > 0; col--)
            {
                bool lastColFound = false;
                for (int row = range.Size.NumberOfRows; row > 0; row--)
                {
                    if (HasValue(range.GetOffset(row, col))) // Vi hittar en column med värde
                    {
                        lastColFound = true;
                        break;
                    }
                }
                if (lastColFound)
                    break;
                tocol--;
            }

            for (int row = range.Size.NumberOfRows; row > 0; row--)
            {
                bool lastRowFound = false;
                for (int col = range.Size.NumberOfCols; col > 0; col--)
                {
                    if (HasValue(range.GetOffset(row, col))) // Vi hittar en rad med värde
                    {
                        lastRowFound = true;
                        break;
                    }
                }
                if (lastRowFound)
                    break;
                toRow--;
            }

            if (toRow == 0 && tocol == 0) // range is empty
            {
                error = CompileResult.GetErrorResult(eErrorType.Ref);
                return new InMemoryRange(1, 1);
            }

            var result = new InMemoryRange(toRow, (short)tocol);
            for (int r = 0; r < toRow; r++)
                for (int c = 0; c < tocol; c++)
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

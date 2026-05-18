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
        EPPlusVersion = "8",
        Description = "",
        SupportsArrays = true)]
    internal class TroAll : ExcelFunction
    {
        public override int ArgumentMinLength => 1;

        public override string NamespacePrefix => "_xlfn.";

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var range = arguments[0].ValueAsRangeInfo;
            var result = TrimRange(range, out var error);
            if (error != null)
                return error;

            return CreateDynamicArrayResult(result, DataType.ExcelRange);
        }

        private InMemoryRange TrimRange(IRangeInfo range, out CompileResult error)
        {
            int nRows = range.Size.NumberOfRows;
            int nCols = range.Size.NumberOfCols;

            error = null;
            // Hitta första och sista rad med minst ett värde
            int firstRow = -1, lastRow = -1;
            for (int r = 0; r < nRows; r++)
            {
                for (int c = 0; c < nCols; c++)
                {
                    if (HasValue(range.GetOffset(r, c)))
                    {
                        if (firstRow == -1) firstRow = r;
                        lastRow = r;
                        break;
                    }
                }
            }

            // Hitta första och sista kolumn med minst ett värde
            int firstCol = -1, lastCol = -1;
            for (int c = 0; c < nCols; c++)
            {
                for (int r = 0; r < nRows; r++)
                {
                    if (HasValue(range.GetOffset(r, c)))
                    {
                        if (firstCol == -1) firstCol = c;
                        lastCol = c;
                        break;
                    }
                }
            }

            // Om hela ranget är tomt, returnera REF
            if (firstRow == -1 || firstCol == -1)
            {
                error = CompileResult.GetErrorResult(eErrorType.Ref);
                return new InMemoryRange(1, 1);
            }                

            int trimmedRows = lastRow - firstRow + 1;
            int trimmedCols = lastCol - firstCol + 1;

            var result = new InMemoryRange(trimmedRows, (short)trimmedCols);
            for (int r = 0; r < trimmedRows; r++)
                for (int c = 0; c < trimmedCols; c++)
                    result.SetValue(r, c, range.GetOffset(firstRow + r, firstCol + c));

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

/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  19/3/2026         EPPlus Software AB           EPPlus v8.6
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.TrimFunctions
{
    internal abstract partial class TrimFunctionsBase :  ExcelFunction
    {
        public override int ArgumentMinLength => 1;
        public override string NamespacePrefix => "_xlfn.";

        protected CompileResult ExecuteTrim(IList<FunctionArgument> arguments, TrimMode rowMode, TrimMode colMode)
        {
            var range = arguments[0].ValueAsRangeInfo;
            var result = TrimRangeCore(range, rowMode, colMode, out var error);
            if (error != null) return error;
            return CreateDynamicArrayResult(result, DataType.ExcelRange);
        }

        private InMemoryRange TrimRangeCore(IRangeInfo range, TrimMode rowMode, TrimMode colMode, out CompileResult error)
        {
            error = null;
            int nRows = range.Size.NumberOfRows;
            int nCols = range.Size.NumberOfCols;

            int firstRow = 0, lastRow = nRows - 1;
            int firstCol = 0, lastCol = nCols - 1;

            if (rowMode == TrimMode.Leading || rowMode == TrimMode.Both) 
                firstRow = FindFirstNonEmptyRow(range, nRows, nCols);

            if (rowMode == TrimMode.Trailing || rowMode == TrimMode.Both)
                lastRow = FindLastNonEmptyRow(range, nRows, nCols);

            if (colMode == TrimMode.Leading || colMode == TrimMode.Both)
                firstCol = FindFirstNonEmptyCol(range, nRows, nCols);

            if (colMode == TrimMode.Trailing || colMode == TrimMode.Both)
                lastCol = FindLastNonEmptyCol(range, nRows, nCols);

            // Range is empty
            if (firstRow > lastRow || firstCol > lastCol)
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

        private int FindFirstNonEmptyRow(IRangeInfo range, int nRows, int nCols)
        {
            for (int r = 0; r < nRows; r++)
                for (int c = 0; c < nCols; c++)
                    if (HasValue(range.GetOffset(r, c))) return r;
            return nRows;           
        }

        private int FindLastNonEmptyRow(IRangeInfo range, int nRows, int nCols)
        {
            for (int r = nRows - 1; r >= 0; r--)
                for (int c = 0; c < nCols; c++)
                    if (HasValue(range.GetOffset(r, c))) return r;
            return -1;
        }

        private int FindFirstNonEmptyCol(IRangeInfo range, int nRows, int nCols)
        {
            for (int c = 0; c < nCols; c++)
                for (int r = 0; r < nRows; r++)
                    if (HasValue(range.GetOffset(r, c))) return c;
            return nCols;
        }

        private int FindLastNonEmptyCol(IRangeInfo range, int nRows, int nCols)
        {
            for (int c = nCols - 1; c >= 0; c--)
                for (int r = 0; r < nRows; r++)
                    if (HasValue(range.GetOffset(r, c))) return c;
            return -1;
        }

        protected bool HasValue(object val)
        {
            if (val == null) return false;
            if (val.GetType().Name == "RowInternal") return false;
            return true;
        }
    }
}

using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.TrimFunctions
{
    internal abstract class TrimFunctionsBase : ExcelFunction
    {
        public override int ArgumentMinLength => 1;
        public override string NamespacePrefix => "_xlfn.";

        protected CompileResult ExecuteTrim(IList<FunctionArgument> arguments, TrimMode rowMode, TrimMode colMode, ParsingContext context)
        {
            var range = arguments[0].ValueAsRangeInfo;
            var result = TrimRangeCore(range, rowMode, colMode, context, out var error);
            if (error != null) return error;
            return CreateDynamicArrayResult(result, DataType.ExcelRange);
        }

        private InMemoryRange TrimRangeCore(IRangeInfo range, TrimMode rowMode, TrimMode colMode, ParsingContext context, out CompileResult error)
        {
            error = null;
            int nRows = range.Size.NumberOfRows;
            int nCols = range.Size.NumberOfCols;

            var dimension = context.CurrentWorksheet?.Dimension;
            if (dimension == null)
            {
                error = CompileResult.GetErrorResult(eErrorType.Ref);
                return new InMemoryRange(1, 1);
            }

            int rangeFromRow = range.Address.FromRow;
            int rangeFromCol = range.Address.FromCol;

            int scanFirstRow = Math.Max(0, dimension.Start.Row - rangeFromRow);
            int scanLastRow = Math.Min(nRows - 1, dimension.End.Row - rangeFromRow);
            int scanFirstCol = Math.Max(0, dimension.Start.Column - rangeFromCol);
            int scanLastCol = Math.Min(nCols - 1, dimension.End.Column - rangeFromCol);

            if (scanFirstRow > scanLastRow || scanFirstCol > scanLastCol)
            {
                error = CompileResult.GetErrorResult(eErrorType.Ref);
                return new InMemoryRange(1, 1);
            }
            int firstRow = 0, lastRow = nRows - 1;
            int firstCol = 0, lastCol = nCols - 1;

            if (rowMode == TrimMode.Leading || rowMode == TrimMode.Both)
                firstRow = FindFirstNonEmptyRow(range, scanFirstRow, scanLastRow, scanFirstCol, scanLastCol);

            if (rowMode == TrimMode.Trailing || rowMode == TrimMode.Both)
                lastRow = FindLastNonEmptyRow(range, scanFirstRow, scanLastRow, scanFirstCol, scanLastCol);

            if (colMode == TrimMode.Leading || colMode == TrimMode.Both)
                firstCol = FindFirstNonEmptyCol(range, scanFirstRow, scanLastRow, scanFirstCol, scanLastCol);

            if (colMode == TrimMode.Trailing || colMode == TrimMode.Both)
                lastCol = FindLastNonEmptyCol(range, scanFirstRow, scanLastRow, scanFirstCol, scanLastCol);

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

        private int FindFirstNonEmptyRow(IRangeInfo range, int rowStart, int rowEnd, int colStart, int colEnd)
        {
            for (int r = rowStart; r <= rowEnd; r++)
                for (int c = colStart; c <= colEnd; c++)
                    if (HasValue(range.GetOffset(r, c))) return r;
            return int.MaxValue; // not found, empty
        }

        private int FindLastNonEmptyRow(IRangeInfo range, int rowStart, int rowEnd, int colStart, int colEnd)
        {
            for (int r = rowEnd; r >= rowStart; r--)
                for (int c = colStart; c <= colEnd; c++)
                    if (HasValue(range.GetOffset(r, c))) return r;
            return -1;
        }

        private int FindFirstNonEmptyCol(IRangeInfo range, int rowStart, int rowEnd, int colStart, int colEnd)
        {
            for (int c = colStart; c <= colEnd; c++)
                for (int r = rowStart; r <= rowEnd; r++)
                    if (HasValue(range.GetOffset(r, c))) return c;
            return int.MaxValue; // not found, empty
        }

        private int FindLastNonEmptyCol(IRangeInfo range, int rowStart, int rowEnd, int colStart, int colEnd)
        {
            for (int c = colEnd; c >= colStart; c--)
                for (int r = rowStart; r <= rowEnd; r++)
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
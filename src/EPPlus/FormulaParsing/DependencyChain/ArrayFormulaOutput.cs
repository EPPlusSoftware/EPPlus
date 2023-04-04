using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Core.CellStore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Filter;

namespace OfficeOpenXml.FormulaParsing
{
    internal class ArrayFormulaOutput
    {
        internal static void FillArrayFromRangeInfo(RpnFormula f, IRangeInfo array, RangeHashset rd, RpnOptimizedDependencyChain depChain)
        {
            var nr = array.Size.NumberOfRows;
            var nc = array.Size.NumberOfCols;
            var ws = f._ws;
            var sf = ws._sharedFormulas[f._arrayIndex];
            var sr = sf.StartRow;
            var sc = sf.StartCol;
            var rows = sf.EndRow - sf.StartRow + 1;
            var cols = sf.EndCol - sf.StartCol + 1;
            var wsIx = ws.IndexInList;
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    var row = sr + r;
                    var col = sc + c;
                    if (r < nr && c < nc)
                    {
                        ws.SetValueInner(row, col, array.GetOffset(r, c) ?? 0);
                    }
                    else
                    {
                        ws.SetValueInner(row, col, ErrorValues.NAError);
                    }
                    var id = ExcelCellBase.GetCellId(wsIx, row, col);
                    depChain.processedCells.Add(id);
                }
            }

            var formulaAddress = new FormulaRangeAddress(depChain._parsingContext) { WorksheetIx = wsIx, FromRow = sr, ToRow = sf.EndRow, FromCol = sc, ToCol = sf.EndCol };
            rd.Merge(ref formulaAddress);
        }

        private static bool HasSpill(ExcelWorksheet ws, int fIx, int startRow, int startColumn, int rows, short columns)
        {
            for (int r = startRow; r < startRow + rows; r++)
            {
                for (int c = startColumn; c < startColumn + columns; c++)
                {
                    if (r == startRow && c == startColumn) continue;
                    object f = -1;
                    if (fIx!=-1 && ws._formulas.Exists(r, c, ref f) && f != null)
                    {
                        if (f is int intfIx && intfIx == fIx)
                        {
                            continue;
                        }
                        return true;
                    }
                    else
                    {
                        var v = ws.GetValueInner(r, c);
                        if(v!=null)
                        {
                            return false;
                        }
                    }
                }
            }
            return false;
        }       
        internal static void FillDynamicArrayFromRangeInfo(RpnFormula f, IRangeInfo array, RangeHashset rd, RpnOptimizedDependencyChain depChain)
        {
            var nr = array.Size.NumberOfRows;
            var nc = array.Size.NumberOfCols;
            var ws = f._ws;
            var sr = f._row;
            var sc = f._column;
            var wsIx = ws.IndexInList;

            f._isDynamic = true;
            var md = depChain._parsingContext.Package.Workbook.Metadata;
            md.GetDynamicArrayIndex(out int cm);
            f._ws._metadataStore.SetValue(f._row, f._column, new ExcelWorksheet.MetaDataReference() { cm = cm });
            //f._ws._flags.SetFlagValue(f._row, f._column, true, CellFlags.ArrayFormula);

            if(HasSpill(ws, f._arrayIndex, sr, sc, nr,nc))
            {
                ws.SetValueInner(sr, sc, ErrorValues.SpillError); //TODO: Spill should be handled on save updating value meta data.
            }
            if(f._arrayIndex==-1)
            {
                f._ws.Cells[f._row, f._column, f._row + array.Size.NumberOfRows - 1, f._column + array.Size.NumberOfCols - 1].CreateArrayFormula(f._formula);
                f._arrayIndex = f._ws.GetMaxShareFunctionIndex(true) - 1;
            }
            else
            {
                var sf = ws._sharedFormulas[f._arrayIndex];
                var endRow = sf.StartRow + nr - 1;
                var endCol = sf.StartCol + nc - 1;
                if(endRow<sf.EndRow)
                {
                    ClearDynamicFormulaIndex(ws, endRow + 1, sf.StartCol, sf.EndRow, sf.EndCol);
                    //clearValues
                }
                else if(endRow > sf.EndRow)
                {
                    SetDynamicFormulaIndex(ws, sf.EndRow + 1, sf.StartCol, endRow, sf.EndCol, f._arrayIndex);
                }
                if (endCol<sf.EndCol)
                {
                    ClearDynamicFormulaIndex(ws, sf.StartRow, endCol + 1, sf.EndRow, sf.EndCol);
                }
                else
                {
                    SetDynamicFormulaIndex(ws, sf.StartRow, sf.StartCol+1, sf.StartCol, sf.EndCol, f._arrayIndex);
                }
                sf.EndRow = endRow;
                sf.EndCol = endCol;
            }
            FillArrayFromRangeInfo(f, array, rd, depChain);
        }
        private static void ClearDynamicFormulaIndex(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws._formulas.Clear(fromRow, fromCol, toRow, toCol);
            ws._flags.Clear(fromRow, fromCol, toRow, toCol);

            for (int col = fromCol; col <= toCol; col++)
            {
                for (int row = fromRow; row <= toRow; row++)
                {
                    ws.SetValueInner(row, col, null);
                }
            }
        }
        private static void SetDynamicFormulaIndex(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol, int formulaIndex)
        {
            for (int col = fromCol; col <= toCol; col++)
            {
                for (int row = fromRow; row <= toRow; row++)
                {
                    ws._formulas.SetValue(row, col, formulaIndex);
                    ws._flags.SetFlagValue(row, col, true, CellFlags.ArrayFormula);
                    ws.SetValueInner(row, col, null);
                }
            }
        }
    }
}
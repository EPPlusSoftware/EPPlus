using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Core.CellStore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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

            for (int r=0;r < rows; r++)
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
        internal static void FillDynamicArrayFromRangeInfo(RpnFormula f, IRangeInfo array, RangeHashset rd, RpnOptimizedDependencyChain depChain)
        {
            f._isDynamic = true;
            var md = depChain._parsingContext.Package.Workbook.Metadata;
            md.GetDynamicArrayIndex(out int cm);
            f._ws._metadataStore.SetValue(f._row, f._column, new ExcelWorksheet.MetaDataReference() { cm = cm });
            f._ws._flags.SetFlagValue(f._row, f._column, true, CellFlags.ArrayFormula);
            f._ws.Cells[f._row, f._column, f._row+array.Size.NumberOfRows-1, f._column+array.Size.NumberOfCols-1].CreateArrayFormula(f._formula);
            f._arrayIndex = f._ws.GetMaxShareFunctionIndex(true) - 1;
            FillArrayFromRangeInfo(f, array, rd, depChain);
        }
    }
}
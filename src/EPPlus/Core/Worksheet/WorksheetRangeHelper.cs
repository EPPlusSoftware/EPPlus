/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                   Change
 *************************************************************************************************
  02/03/2020         EPPlus Software AB       Added
 *************************************************************************************************/
 using OfficeOpenXml.Drawing;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using static OfficeOpenXml.ExcelAddressBase;

namespace OfficeOpenXml.Core.Worksheet
{
    internal static class WorksheetRangeHelper
    {
        internal static void FixMergedCells(ExcelWorksheet ws, ExcelRangeBase range, eShiftTypeInsert shift)
        {
            if(shift==eShiftTypeInsert.Down)
            {
                FixMergedCellsRow(ws, range._fromRow, range.Rows, false, range._fromCol, range._toCol);
            }
            else
            {
                FixMergedCellsColumn(ws, range._fromCol, range.Columns, false, range._fromRow, range._toRow);
            }
        }
        internal static void FixMergedCells(ExcelWorksheet ws, ExcelRangeBase range, eShiftTypeDelete shift)
        {
            if (shift == eShiftTypeDelete.Up)
            {
                FixMergedCellsRow(ws, range._fromRow, range.Rows, true, range._fromCol, range._toCol);
            }
            else
            {
                FixMergedCellsColumn(ws, range._fromCol, range.Columns, true, range._fromRow, range._toRow);
            }
        }

        internal static void FixMergedCellsRow(ExcelWorksheet ws, int row, int rows, bool delete, int fromCol=1, int toCol=ExcelPackage.MaxColumns)
        {
            List<int> removeIndex = new List<int>();
            for (int i = 0; i < ws._mergedCells.Count; i++)
            {
                if (!string.IsNullOrEmpty(ws._mergedCells[i]))
                {
                    ExcelAddressBase addr = new ExcelAddressBase(ws._mergedCells[i]), newAddr;
                    if (addr._fromCol >= fromCol && addr._toCol <= toCol)
                    {
                        if (delete)
                        {
                            newAddr = addr.DeleteRow(row, rows);
                            if (newAddr == null)
                            {
                                removeIndex.Add(i);
                                continue;
                            }
                        }
                        else
                        {
                            newAddr = addr.AddRow(row, rows);
                            if (newAddr.Address != addr.Address)
                            {
                                ws._mergedCells.SetIndex(newAddr, i);
                            }
                        }

                        if (newAddr.Address != addr.Address)
                        {
                            ws._mergedCells._list[i] = newAddr._address;
                        }
                    }
                }
            }
            for (int i = removeIndex.Count - 1; i >= 0; i--)
            {
                ws._mergedCells._list[removeIndex[i]] = null;
            }
        }
        internal static void FixMergedCellsColumn(ExcelWorksheet ws, int column, int columns, bool delete, int fromRow = 1, int toRow = ExcelPackage.MaxRows)
        {
            List<int> removeIndex = new List<int>();
            for (int i = 0; i < ws._mergedCells.Count; i++)
            {
                if (!string.IsNullOrEmpty(ws._mergedCells[i]))
                {
                    ExcelAddressBase addr = new ExcelAddressBase(ws._mergedCells[i]), newAddr;

                    if (addr._fromRow >= fromRow && addr._toRow <= toRow)
                    {
                        if (delete)
                        {
                            newAddr = addr.DeleteColumn(column, columns);
                            if (newAddr == null)
                            {
                                removeIndex.Add(i);
                                continue;
                            }
                        }
                        else
                        {
                            newAddr = addr.AddColumn(column, columns);
                            if (newAddr.Address != addr.Address)
                            {
                                ws._mergedCells.SetIndex(newAddr, i);
                            }
                        }

                        if (newAddr.Address != addr.Address)
                        {
                            ws._mergedCells._list[i] = newAddr._address;
                        }
                    }
                }
            }
            for (int i = removeIndex.Count - 1; i >= 0; i--)
            {
                ws._mergedCells._list[removeIndex[i]]=null;
            }
        }
        internal static void AdjustDrawingsRow(ExcelWorksheet ws, int rowFrom, int rows, int colFrom=0, int colTo=ExcelPackage.MaxColumns)
        {
            var deletedDrawings = new List<ExcelDrawing>();
            foreach (ExcelDrawing drawing in ws.Drawings)
            {

                if (!drawing.IsWithinColumnRange(colFrom, colTo))
                {
                    continue;
                }

                if(drawing.CellAnchor == eEditAs.TwoCell && 
                    rows < 0 && drawing.From.Row>=rowFrom-1 && 
                    ((drawing.To.Row<=(rowFrom-rows-1) && drawing.To.RowOff==0) || drawing.To.Row <= (rowFrom - rows - 2))) //If delete and the entire drawing is withing the deleted range, remove it.
                {
                    deletedDrawings.Add(drawing);
                    continue;
                }
                if (drawing.EditAs != eEditAs.Absolute)
                {
                    if (drawing.From.Row >= rowFrom-1)
                    {                       
                        if (drawing.From.Row + rows < rowFrom - 1)
                        {
                            drawing.From.Row = rowFrom - 1;
                            drawing.From.RowOff = 0;
                        }
                        else
                        {
                            drawing.From.Row += rows;
                        }

                        if (drawing.CellAnchor == eEditAs.TwoCell)
                        {
                            if (drawing.To.Row >= rowFrom-1)
                            {
                                drawing.To.Row += rows;
                            }
                        }
                    }
                    else if (drawing.To != null && drawing.To.Row >= rowFrom-1)
                    {
                        if (drawing.EditAs == eEditAs.TwoCell)
                        {
                            if (drawing.To.Row + rows < rowFrom - 1)
                            {
                                drawing.To.Row = rowFrom - 1;
                                drawing.To.RowOff = 0;
                            }
                            else
                            {
                                drawing.To.Row += rows;
                            }
                        }
                    }
                    if (drawing.From.Row < 0) drawing.From.Row = 0;
                    drawing.AdjustPositionAndSize();
                }
            }

            deletedDrawings.ForEach(d => ws.Drawings.Remove(d));
        }
        internal static void AdjustDrawingsColumn(ExcelWorksheet ws, int columnFrom, int columns, int rowFrom = 0, int rowTo = ExcelPackage.MaxRows)
        {
            var deletedDrawings = new List<ExcelDrawing>();
            foreach (ExcelDrawing drawing in ws.Drawings)
            {
                if(!drawing.IsWithinRowRange(rowFrom, rowTo))
                {
                    continue;
                }

                if (drawing.CellAnchor==eEditAs.TwoCell && 
                    columns < 0 && drawing.From.Column >= columnFrom - 1 &&
                    ((drawing.To.Column <= (columnFrom - columns - 1) && drawing.To.ColumnOff == 0) || drawing.To.Column <= (columnFrom - columns - 2))) //If delete and the entire drawing is withing the deleted range, remove it.
                {
                    deletedDrawings.Add(drawing);
                    continue;
                }
                if (drawing.EditAs != eEditAs.Absolute)
                {
                    if (drawing.From.Column >= columnFrom - 1)
                    {
                        if (drawing.From.Column + columns < columnFrom - 1)
                        {
                            drawing.From.Column = columnFrom - 1;
                            drawing.From.ColumnOff = 0;
                        }
                        else
                        {
                            drawing.From.Column += columns;
                        }

                        if (drawing.EditAs == eEditAs.TwoCell)
                        {
                            if (drawing.To.Column >= columnFrom - 1)
                            {
                                drawing.To.Column += columns;
                            }
                        }
                    }
                    else if (drawing.To!=null && drawing.To.Column >= columnFrom - 1)
                    {
                        if (drawing.To.Column + columns < columnFrom - 1)
                        {
                            drawing.To.Column = columnFrom - 1;
                            drawing.To.ColumnOff = 0;
                        }
                        else
                        {
                            drawing.To.Column += columns;
                        }
                    }
                    if (drawing.From.Column < 0) drawing.From.Column = 0;
                    drawing.AdjustPositionAndSize();
                }
            }

            deletedDrawings.ForEach(d => ws.Drawings.Remove(d));
        }
        internal static void ConvertEffectedSharedFormulasToCellFormulas(ExcelWorksheet wsUpdate, ExcelAddressBase range)
        {
            foreach (var ws in wsUpdate.Workbook.Worksheets)
            {
                bool isCurrentWs = wsUpdate.Name.Equals(ws.Name, StringComparison.CurrentCultureIgnoreCase);
                var deletedSf = new List<int>();
                foreach (var sf in ws._sharedFormulas.Values)
                {
                    //Do not convert array formulas.
                    if (sf.FormulaType == ExcelWorksheet.FormulaType.Shared && (isCurrentWs || sf.Formula.IndexOf(wsUpdate.Name, StringComparison.CurrentCultureIgnoreCase) >= 0))
                    {
                        if (ConvertEffectedSharedFormulaIfReferenceWithinRange(ws, range, sf, wsUpdate.Name))
                        {
                            deletedSf.Add(sf.Index);
                        }
                    }
                }
                deletedSf.ForEach(x => ws._sharedFormulas.Remove(x));
            }
        }
        private static bool ConvertEffectedSharedFormulaIfReferenceWithinRange(ExcelWorksheet ws, ExcelAddressBase delRange, ExcelWorksheet.Formulas sf, string wsName)
        {
            bool doConvertSF = false;
            var sfAddress = new ExcelAddressBase(sf.Address);
            sf.SetTokens(ws.Name);
            foreach (var token in sf.Tokens)
            {
                if (token.TokenTypeIsSet(TokenType.ExcelAddress))
                {
                    //Check if the address for the entire shared formula collides with the deleted address.
                    var tokenAddress = new ExcelAddressBase(token.Value);
                    if ((ws.Name.Equals(wsName, StringComparison.CurrentCultureIgnoreCase) && string.IsNullOrEmpty(tokenAddress.WorkSheetName)) ||
                        (!string.IsNullOrEmpty(tokenAddress.WorkSheetName) && tokenAddress.WorkSheetName.Equals(wsName, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        if (tokenAddress._toRowFixed == false) tokenAddress._toRow += (sfAddress.Rows - 1);
                        if (tokenAddress._toColFixed == false) tokenAddress._toCol += (sfAddress.Columns - 1);

                        if (tokenAddress.Collide(delRange, true) != ExcelAddressBase.eAddressCollition.No)  //Shared Formula address is effected.
                        {
                            doConvertSF = true;
                            break;
                        }
                    }
                }
            }
            if (doConvertSF)
            {
                ConvertSharedFormulaToCellFormula(ws, sf, sfAddress);
            }
            return doConvertSF;
        }
        private static void ConvertSharedFormulaToCellFormula(ExcelWorksheet ws, ExcelWorksheet.Formulas sf, ExcelAddressBase sfAddress)
        {
            for (var r = 0; r < sfAddress.Rows; r++)
            {
                for (var c = 0; c < sfAddress.Columns; c++)
                {
                    var row = sf.StartRow + r;
                    var col = sf.StartCol + c;
                    ws._formulas.SetValue(row, col, ws.GetFormula(row, col));
                }
            }
        }
        internal static void ValidateIfInsertDeleteIsPossible(ExcelRangeBase range, ExcelAddressBase effectedAddress, ExcelAddressBase effectedAddressTable, bool insert)
        {
            //Validate autofilter
            if (range.Worksheet.AutoFilterAddress!=null && 
                effectedAddress.Collide(range.Worksheet.AutoFilterAddress) == ExcelAddressBase.eAddressCollition.Partly 
                    &&
                    range.Worksheet.AutoFilterAddress.CollideFullRowOrColumn(range) == false)
            {
                throw new InvalidOperationException($"Can't {(insert ? "insert into" : "delete from")} the range. Cells collide with the worksheets autofilter.");
            }

            //Validate merged Cells
            foreach (var a in range.Worksheet.MergedCells)
            {
                if (!string.IsNullOrEmpty(a))
                {
                    var mc = new ExcelAddressBase(a);
                    if (effectedAddress.Collide(mc) == ExcelAddressBase.eAddressCollition.Partly)
                    {
                        throw new InvalidOperationException($"Can't {(insert ? "insert into" : "delete from")} the range. Cells collide with merged range {a}");
                    }
                }
            }

            //Validate pivot tables Cells
            foreach (var pt in range.Worksheet.PivotTables)
            {
                if (effectedAddress.Collide(pt.Address) == ExcelAddressBase.eAddressCollition.Partly)
                {
                    throw new InvalidOperationException($"Can't {(insert ? "insert into" : "delete from")} the range. Cells collide with pivot table {pt.Name}");
                }
            }

            //Validate tables Cells
            foreach (var t in range.Worksheet.Tables)
            {
                if (effectedAddressTable.Collide(t.Address) == ExcelAddressBase.eAddressCollition.Partly &&
                    effectedAddress.Collide(t.Address) != ExcelAddressBase.eAddressCollition.No)
                {
                    throw new InvalidOperationException($"Can't {(insert ? "insert into" : "delete from")} the range. Cells collide with table {t.Name}");
                }
            }

        }

    }
}

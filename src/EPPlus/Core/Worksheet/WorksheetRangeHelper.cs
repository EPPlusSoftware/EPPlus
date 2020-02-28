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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet
{
    internal static class WorksheetRangeHelper
    {
        internal static void FixMergedCellsRow(ExcelWorksheet ws, int row, int rows, bool delete)
        {
            if (delete)
            {
                ws._mergedCells._cells.Delete(row, 0, rows, ExcelPackage.MaxColumns + 1);
            }
            else
            {
                ws._mergedCells._cells.Insert(row, 0, rows, ExcelPackage.MaxColumns + 1);
            }

            List<int> removeIndex = new List<int>();
            for (int i = 0; i < ws._mergedCells.Count; i++)
            {
                if (!string.IsNullOrEmpty(ws._mergedCells[i]))
                {
                    ExcelAddressBase addr = new ExcelAddressBase(ws._mergedCells[i]), newAddr;
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
            for (int i = removeIndex.Count - 1; i >= 0; i--)
            {
                ws._mergedCells._list.RemoveAt(removeIndex[i]);
            }
        }
        internal static void FixMergedCellsColumn(ExcelWorksheet ws, int column, int columns, bool delete)
        {
            if (delete)
            {
                ws._mergedCells._cells.Delete(0, column, 0, columns);
            }
            else
            {
                ws._mergedCells._cells.Insert(0, column, 0, columns);
            }
            List<int> removeIndex = new List<int>();
            for (int i = 0; i < ws._mergedCells.Count; i++)
            {
                if (!string.IsNullOrEmpty(ws._mergedCells[i]))
                {
                    ExcelAddressBase addr = new ExcelAddressBase(ws._mergedCells[i]), newAddr;
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
            for (int i = removeIndex.Count - 1; i >= 0; i--)
            {
                ws._mergedCells._list.RemoveAt(removeIndex[i]);
            }
        }
        internal static void AdjustDrawingsRow(ExcelWorksheet ws, int rowFrom, int rows)
        {
            var deletedDrawings = new List<ExcelDrawing>();
            foreach (ExcelDrawing drawing in ws.Drawings)
            {                
                if(rows < 0 && drawing.From.Row>=rowFrom-1 && 
                    ((drawing.To.Row<=(rowFrom-rows-1) && drawing.To.RowOff==0) || drawing.To.Row <= (rowFrom - rows - 2))) //If delete and the entire drawing is withing the deleted range, remove it.
                {
                    deletedDrawings.Add(drawing);
                    continue;
                }
                if (drawing.EditAs != eEditAs.Absolute)
                {
                    if (drawing.From.Row >= rowFrom-1)
                    {
                        drawing.From.RowOff = 0;
                        if (drawing.From.Row + rows < rowFrom - 1)
                        {
                            drawing.From.Row = rowFrom - 1;
                        }
                        else
                        {
                            drawing.From.Row += rows;
                        }

                        if (drawing.EditAs == eEditAs.TwoCell)
                        {
                            if (drawing.To.Row >= rowFrom-1)
                            {                                
                                drawing.To.Row += rows;
                            }
                        }
                    }
                    else if (drawing.To.Row >= rowFrom-1)
                    {
                        drawing.To.RowOff = 0;
                        if (drawing.To.Row+rows < rowFrom-1)
                        {
                            drawing.To.Row = rowFrom-1;
                        }
                        else
                        {
                            drawing.To.Row += rows;
                        }
                    }
                    if (drawing.From.Row < 0) drawing.From.Row = 0;
                    drawing.AdjustPositionAndSize();
                }
            }

            deletedDrawings.ForEach(d => ws.Drawings.Remove(d));
        }
        internal static void AdjustDrawingsColumn(ExcelWorksheet ws, int columnFrom, int columns)
        {
            var deletedDrawings = new List<ExcelDrawing>();
            foreach (ExcelDrawing drawing in ws.Drawings)
            {
                if (columns < 0 && drawing.From.Column >= columnFrom - 1 &&
                    ((drawing.To.Column <= (columnFrom - columns - 1) && drawing.To.ColumnOff == 0) || drawing.To.Column <= (columnFrom - columns - 2))) //If delete and the entire drawing is withing the deleted range, remove it.
                {
                    deletedDrawings.Add(drawing);
                    continue;
                }
                if (drawing.EditAs != eEditAs.Absolute)
                {
                    if (drawing.From.Column >= columnFrom - 1)
                    {
                        drawing.From.ColumnOff = 0;
                        if (drawing.From.Column + columns < columnFrom - 1)
                        {
                            drawing.From.Column = columnFrom - 1;
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
                    else if (drawing.To.Column >= columnFrom - 1)
                    {
                        drawing.To.ColumnOff = 0;
                        if (drawing.To.Column + columns < columnFrom - 1)
                        {
                            drawing.To.Column = columnFrom - 1;
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
    }
}

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
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet
{
    internal static class WorksheetRangeDeleteHelper 
    {
        internal static void DeleteRow(ExcelWorksheet ws, int rowFrom, int rows)
        {
            ws.CheckSheetType();
            ValidateRow(rowFrom, rows);
            lock (ws)
            {
                DeleteCellStores(ws, rowFrom, 0, rows, ExcelPackage.MaxColumns + 1, true);

                AdjustFormulasRow(ws, rowFrom, rows);
                WorksheetRangeHelper.FixMergedCellsRow(ws, rowFrom, rows, true);

                foreach (var tbl in ws.Tables)
                {
                    tbl.Address = tbl.Address.DeleteRow(rowFrom, rows);
                }

                foreach (var ptbl in ws.PivotTables)
                {
                    if (ptbl.Address.Start.Row > rowFrom + rows)
                    {
                        ptbl.Address = ptbl.Address.DeleteRow(rowFrom, rows);
                    }
                }
                //Issue 15573
                foreach (ExcelDataValidation dv in ws.DataValidations)
                {
                    var addr = dv.Address;
                    if (addr.Start.Row > rowFrom + rows)
                    {
                        var newAddr = addr.DeleteRow(rowFrom, rows).Address;
                        if (addr.Address != newAddr)
                        {
                            dv.SetAddress(newAddr);
                        }
                    }
                }

                WorksheetRangeHelper.AdjustDrawingsRow(ws, rowFrom, -rows);
            }
        }
        internal static void DeleteColumn(ExcelWorksheet ws, int columnFrom, int columns)
        {
            ValidateColumn(columnFrom, columns);
            lock (ws)
            {
                //Set previous column Max to Row before if it spans the deleted column range.
                ExcelColumn col = ws.GetValueInner(0, columnFrom) as ExcelColumn;
                if (col == null)
                {
                    var r = 0;
                    var c = columnFrom;
                    if (ws._values.PrevCell(ref r, ref c))
                    {
                        col = ws.GetValueInner(0, c) as ExcelColumn;
                        if (col._columnMax >= columnFrom)
                        {
                            col.ColumnMax = columnFrom - 1;
                        }
                    }
                }

                DeleteCellStores(ws, 0, columnFrom, 0, columns, true);

                AdjustFormulasColumn(ws, columnFrom, columns);
                WorksheetRangeHelper.FixMergedCellsColumn(ws, columnFrom, columns, true);

                foreach (var tbl in ws.Tables)
                {
                    if (columnFrom >= tbl.Address.Start.Column && columnFrom <= tbl.Address.End.Column)
                    {
                        var node = tbl.Columns[0].TopNode.ParentNode;
                        var ix = columnFrom - tbl.Address.Start.Column;
                        for (int i = 0; i < columns; i++)
                        {
                            if (node.ChildNodes.Count > ix)
                            {
                                node.RemoveChild(node.ChildNodes[ix]);
                            }
                        }
                        tbl._cols = new ExcelTableColumnCollection(tbl);
                    }

                    tbl.Address = tbl.Address.DeleteColumn(columnFrom, columns);

                    foreach (var ptbl in ws.PivotTables)
                    {
                        if (ptbl.Address.Start.Column > columnFrom + columns)
                        {
                            ptbl.Address = ptbl.Address.DeleteColumn(columnFrom, columns);
                        }
                        if (ptbl.CacheDefinition.SourceRange.Start.Column > columnFrom + columns)
                        {
                            ptbl.CacheDefinition.SourceRange.Address = ptbl.CacheDefinition.SourceRange.DeleteColumn(columnFrom, columns).Address;
                        }
                    }
                }

                //Adjust DataValidation
                foreach (ExcelDataValidation dv in ws.DataValidations)
                {
                    var addr = dv.Address;
                    if (addr.Start.Column > columnFrom + columns)
                    {
                        var newAddr = addr.DeleteColumn(columnFrom, columns).Address;
                        if (addr.Address != newAddr)
                        {
                            dv.SetAddress(newAddr);
                        }
                    }
                }

                //Adjust drawing positions.
                WorksheetRangeHelper.AdjustDrawingsColumn(ws, columnFrom, -columns);
            }
        }

        private static void ValidateRow(int rowFrom, int rows)
        {
            if (rowFrom < 1 || rowFrom + rows > ExcelPackage.MaxRows)
            {
                throw (new ArgumentException("rowFrom", "Row out of range. Spans from 1 to " + ExcelPackage.MaxRows.ToString(CultureInfo.InvariantCulture)));
            }
        }
        private static void ValidateColumn(int columnFrom, int columns)
        {
            if (columnFrom < 1 || columnFrom + columns > ExcelPackage.MaxColumns)
            {
                throw (new ArgumentException("columnFrom", "Column out of range. Spans from 1 to " + ExcelPackage.MaxColumns.ToString(CultureInfo.InvariantCulture)));
            }
        }

        private static void DeleteCellStores(ExcelWorksheet ws, int rowFrom, int columnFrom, int rows, int columns, bool shift)
        {
            //Store
            ws._values.Delete(rowFrom, columnFrom, rows, columns, shift);
            ws._formulas.Delete(rowFrom, columnFrom, rows, columns, shift);
            ws._flags.Delete(rowFrom, columnFrom, rows, columns, shift);
            ws._commentsStore.Delete(rowFrom, columnFrom, rows, columns, shift);
            ws._hyperLinks.Delete(rowFrom, columnFrom, rows, columns, shift);

            ws._names.Delete(rowFrom, columnFrom, rows, columns);
            ws.Comments.Delete(rowFrom, columnFrom, rows, columns);
            ws.Workbook.Names.Delete(rowFrom, columnFrom, rows, columns, n => n.Worksheet == ws);

            if (rowFrom == 0 && rows >= ExcelPackage.MaxRows) //Delete full column
            {
                AdjustColumnMinMax(ws, columnFrom, columns);
            }
        }
        private static void AdjustColumnMinMax(ExcelWorksheet ws, int columnFrom, int columns)
        {
            var csec = new CellStoreEnumerator<ExcelValue>(ws._values, 0, columnFrom, 0, columnFrom + columns - 1);
            foreach (var val in csec)
            {
                var column = val._value;
                if (column is ExcelColumn)
                {
                    var c = (ExcelColumn)column;
                    if (c._columnMin >= columnFrom)
                    {
                        c._columnMin += columns;
                        c._columnMax += columns;
                    }
                }
            }
        }
        static void AdjustFormulasRow(ExcelWorksheet ws, int rowFrom, int rows)
        {
            var delSF = new List<int>();
            foreach (var sf in ws._sharedFormulas.Values)
            {
                var a = new ExcelAddress(sf.Address).DeleteRow(rowFrom, rows);
                if (a == null)
                {
                    delSF.Add(sf.Index);
                }
                else
                {
                    sf.Address = a.Address;
                    if (sf.StartRow > rowFrom)
                    {
                        var r = Math.Min(sf.StartRow - rowFrom, rows);
                        sf.Formula = ExcelCellBase.UpdateFormulaReferences(sf.Formula, -r, 0, rowFrom, 0, ws.Name, ws.Name);
                        sf.StartRow -= r;
                    }
                }
            }
            foreach (var ix in delSF)
            {
                ws._sharedFormulas.Remove(ix);
            }
            var cse = new CellStoreEnumerator<object>(ws._formulas, 1, 1, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
            while (cse.Next())
            {
                if (cse.Value is string)
                {
                    cse.Value = ExcelCellBase.UpdateFormulaReferences(cse.Value.ToString(), -rows, 0, rowFrom, 0, ws.Name, ws.Name);
                }
            }
        }
        internal static void AdjustFormulasColumn(ExcelWorksheet ws, int columnFrom, int columns)
        {
            var delSF = new List<int>();
            foreach (var sf in ws._sharedFormulas.Values)
            {
                var a = new ExcelAddress(sf.Address).DeleteColumn(columnFrom, columns);
                if (a == null)
                {
                    delSF.Add(sf.Index);
                }
                else
                {
                    sf.Address = a.Address;
                    if (sf.StartCol > columnFrom)
                    {
                        var c = Math.Min(sf.StartCol - columnFrom, columns);
                        sf.Formula = ExcelCellBase.UpdateFormulaReferences(sf.Formula, 0, -c, 0, 1, ws.Name, ws.Name);
                        sf.StartCol -= c;
                    }
                }
            }
            foreach (var ix in delSF)
            {
                ws._sharedFormulas.Remove(ix);
            }
            delSF = null;
            var cse = new CellStoreEnumerator<object>(ws._formulas, 1, 1, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
            while (cse.Next())
            {
                if (cse.Value is string)
                {
                    cse.Value = ExcelCellBase.UpdateFormulaReferences(cse.Value.ToString(), 0, -columns, 0, columnFrom, ws.Name, ws.Name);
                }
            }
        }
    }
}

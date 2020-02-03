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
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Core.Worksheet
{
    internal static class WorksheetRangeInsertHelper
    {
        internal static void InsertRow(ExcelWorksheet ws, int rowFrom, int rows, int copyStylesFromRow)
        {
            ValidateInsertRow(ws, rowFrom, rows);

            lock (ws)
            {
                InsertCellStores(ws, rowFrom, 0, rows, 0);

                foreach (var f in ws._sharedFormulas.Values)
                {
                    if (f.StartRow >= rowFrom) f.StartRow += rows;
                    var a = new ExcelAddressBase(f.Address);
                    if (a._fromRow >= rowFrom)
                    {
                        a._fromRow += rows;
                        a._toRow += rows;
                    }
                    else if (a._toRow >= rowFrom)
                    {
                        a._toRow += rows;
                    }
                    f.Address = ExcelCellBase.GetAddress(a._fromRow, a._fromCol, a._toRow, a._toCol);
                    f.Formula = ExcelCellBase.UpdateFormulaReferences(f.Formula, rows, 0, rowFrom, 0, ws.Name, ws.Name);
                }

                var cse = new CellStoreEnumerator<object>(ws._formulas);
                while (cse.Next())
                {
                    if (cse.Value is string)
                    {
                        cse.Value = ExcelCellBase.UpdateFormulaReferences(cse.Value.ToString(), rows, 0, rowFrom, 0, ws.Name, ws.Name);
                    }
                }
                WorksheetRangeHelper.FixMergedCellsRow(ws, rowFrom, rows, false);
                if (copyStylesFromRow > 0)
                {
                    CopyFromStyleRow(ws, rowFrom, rows, copyStylesFromRow);
                }
                foreach (var tbl in ws.Tables)
                {
                    tbl.Address = tbl.Address.AddRow(rowFrom, rows);
                }
                foreach (var ptbl in ws.PivotTables)
                {
                    ptbl.Address = ptbl.Address.AddRow(rowFrom, rows);
                    ptbl.CacheDefinition.SourceRange.Address = ptbl.CacheDefinition.SourceRange.AddRow(rowFrom, rows).Address;
                }

                //Update data validation references
                foreach (ExcelDataValidation dv in ws.DataValidations)
                {
                    var addr = dv.Address;
                    var newAddr = addr.AddRow(rowFrom, rows).Address;
                    if (addr.Address != newAddr)
                    {
                        dv.SetAddress(newAddr);
                    }
                }

                // Update sheetname in cross reference formulas
                foreach (var sheet in ws.Workbook.Worksheets.Where(otherWs => otherWs != ws))
                {
                    sheet.UpdateSheetNameInFormulas(ws.Name, rowFrom, rows, 0, 0);
                }

                WorksheetRangeHelper.AdjustDrawingsRow(ws, rowFrom, rows);
            }
        }
        internal static void InsertColumn(ExcelWorksheet ws, int columnFrom, int columns, int copyStylesFromColumn)
        {
            ValidateInsertColumn(ws, columnFrom, columns);

            lock (ws)
            {
                InsertCellStores(ws, 0, columnFrom, 0, columns);

                InsertColumnFormulas(ws, columnFrom, columns);

                WorksheetRangeHelper.FixMergedCellsColumn(ws, columnFrom, columns, false);

                var csec = new CellStoreEnumerator<ExcelValue>(ws._values, 0, 1, 0, ExcelPackage.MaxColumns);
                var lst = new List<ExcelColumn>();
                foreach (var val in csec)
                {
                    var col = val._value;
                    if (col is ExcelColumn)
                    {
                        lst.Add((ExcelColumn)col);
                    }
                }

                for (int i = lst.Count - 1; i >= 0; i--)
                {
                    var c = lst[i];
                    if (c._columnMin >= columnFrom)
                    {
                        if (c._columnMin + columns <= ExcelPackage.MaxColumns)
                        {
                            c._columnMin += columns;
                        }
                        else
                        {
                            c._columnMin = ExcelPackage.MaxColumns;
                        }

                        if (c._columnMax + columns <= ExcelPackage.MaxColumns)
                        {
                            c._columnMax += columns;
                        }
                        else
                        {
                            c._columnMax = ExcelPackage.MaxColumns;
                        }
                    }
                    else if (c._columnMax >= columnFrom)
                    {
                        var cc = c._columnMax - columnFrom;
                        c._columnMax = columnFrom - 1;
                        ws.CopyColumn(c, columnFrom + columns, columnFrom + columns + cc);
                    }
                }


                //Copy style from another column?
                if (copyStylesFromColumn > 0)
                {
                    if (copyStylesFromColumn >= columnFrom)
                    {
                        copyStylesFromColumn += columns;
                    }

                    //Get styles to a cached list, 
                    var l = new List<int[]>();
                    var sce = new CellStoreEnumerator<ExcelValue>(ws._values, 0, copyStylesFromColumn, ExcelPackage.MaxRows, copyStylesFromColumn);
                    lock (sce)
                    {
                        while (sce.Next())
                        {
                            if (sce.Value._styleId == 0) continue;
                            l.Add(new int[] { sce.Row, sce.Value._styleId });
                        }
                    }

                    //Set the style id's from the list.
                    foreach (var sc in l)
                    {
                        for (var c = 0; c < columns; c++)
                        {
                            if (sc[0] == 0)
                            {
                                var col = ws.Column(columnFrom + c);   //Create the column
                                col.StyleID = sc[1];
                            }
                            else
                            {
                                ws.SetStyleInner(sc[0], columnFrom + c, sc[1]);
                            }
                        }
                    }
                    var newOutlineLevel = ws.Column(copyStylesFromColumn).OutlineLevel;
                    for (var c = 0; c < columns; c++)
                    {
                        ws.Column(columnFrom + c).OutlineLevel = newOutlineLevel;
                    }
                }

                //Adjust tables
                foreach (var tbl in ws.Tables)
                {
                    if (columnFrom > tbl.Address.Start.Column && columnFrom <= tbl.Address.End.Column)
                    {
                        InsertTableColumns(columnFrom, columns, tbl);
                    }

                    tbl.Address = tbl.Address.AddColumn(columnFrom, columns);
                }
                foreach (var ptbl in ws.PivotTables)
                {
                    if (columnFrom <= ptbl.Address.End.Column)
                    {
                        ptbl.Address = ptbl.Address.AddColumn(columnFrom, columns);
                    }
                    if (columnFrom <= ptbl.CacheDefinition.SourceRange.End.Column)
                    {
                        if (ptbl.CacheDefinition.CacheSource == eSourceType.Worksheet)
                        {
                            ptbl.CacheDefinition.SourceRange.Address = ptbl.CacheDefinition.SourceRange.AddColumn(columnFrom, columns).Address;
                        }
                    }
                }

                //Adjust DataValidation
                foreach (ExcelDataValidation dv in ws.DataValidations)
                {
                    var addr = dv.Address;
                    var newAddr = addr.AddColumn(columnFrom, columns).Address;
                    if (addr.Address != newAddr)
                    {
                        dv.SetAddress(newAddr);
                    }
                }

                // Update cross-sheet references.
                foreach (var sheet in ws.Workbook.Worksheets.Where(otherSheet => otherSheet != ws))
                {
                    sheet.UpdateSheetNameInFormulas(ws.Name, 0, 0, columnFrom, columns);
                }

                //Adjust drawing positions.
                WorksheetRangeHelper.AdjustDrawingsColumn(ws, columnFrom, columns);
            }
        }
        private static void InsertColumnFormulas(ExcelWorksheet ws, int columnFrom, int columns)
        {
            foreach (var f in ws._sharedFormulas.Values)
            {
                if (f.StartCol >= columnFrom) f.StartCol += columns;
                var a = new ExcelAddressBase(f.Address);
                if (a._fromCol >= columnFrom)
                {
                    a._fromCol += columns;
                    a._toCol += columns;
                }
                else if (a._toCol >= columnFrom)
                {
                    a._toCol += columns;
                }
                f.Address = ExcelCellBase.GetAddress(a._fromRow, a._fromCol, a._toRow, a._toCol);
                f.Formula = ExcelCellBase.UpdateFormulaReferences(f.Formula, 0, columns, 0, columnFrom, ws.Name, ws.Name);
            }

            var cse = new CellStoreEnumerator<object>(ws._formulas);
            while (cse.Next())
            {
                if (cse.Value is string)
                {
                    cse.Value = ExcelCellBase.UpdateFormulaReferences(cse.Value.ToString(), 0, columns, 0, columnFrom, ws.Name, ws.Name);
                }
            }
        }

        private static void ValidateInsertColumn(ExcelWorksheet ws, int columnFrom, int columns)
        {
            ws.CheckSheetType();
            var d = ws.Dimension;

            if (columnFrom < 1)
            {
                throw (new ArgumentOutOfRangeException("columnFrom can't be lesser that 1"));
            }
            //Check that cells aren't shifted outside the boundries
            if (d != null && d.End.Column > columnFrom && d.End.Column + columns > ExcelPackage.MaxColumns)
            {
                throw (new ArgumentOutOfRangeException("Can't insert. Columns will be shifted outside the boundries of the worksheet."));
            }
        }

        #region private methods
        private static void ValidateInsertRow(ExcelWorksheet ws, int rowFrom, int rows)
        {
            ws.CheckSheetType();
            var d = ws.Dimension;

            if (rowFrom < 1)
            {
                throw (new ArgumentOutOfRangeException("rowFrom can't be lesser that 1"));
            }

            //Check that cells aren't shifted outside the boundries
            if (d != null && d.End.Row > rowFrom && d.End.Row + rows > ExcelPackage.MaxRows)
            {
                throw (new ArgumentOutOfRangeException("Can't insert. Rows will be shifted outside the boundries of the worksheet."));
            }
        }
        internal static void InsertCellStores(ExcelWorksheet ws, int rowFrom, int columnFrom, int rows, int columns)
        {
            ws._values.Insert(rowFrom, columnFrom, rows, columns);
            ws._formulas.Insert(rowFrom, columnFrom, rows, columns);
            ws._commentsStore.Insert(rowFrom, columnFrom, rows, columns);
            ws._hyperLinks.Insert(rowFrom, columnFrom, rows, columns);
            ws._flags.Insert(rowFrom, columnFrom, rows, columns);

            ws.Comments.Insert(rowFrom, columnFrom, rows, columns);
            ws._names.Insert(rowFrom, columnFrom, rows, columns);
            ws.Workbook.Names.Insert(rowFrom, columnFrom, rows, columns, n => n.Worksheet == ws);
        }
        private static void CopyFromStyleRow(ExcelWorksheet ws, int rowFrom, int rows, int copyStylesFromRow)
        {
            //Copy style from style row
            using (var cseS = new CellStoreEnumerator<ExcelValue>(ws._values, copyStylesFromRow, 0, copyStylesFromRow, ExcelPackage.MaxColumns))
            {
                while (cseS.Next())
                {
                    if (cseS.Value._styleId == 0) continue;
                    for (var r = 0; r < rows; r++)
                    {
                        ws.SetStyleInner(rowFrom + r, cseS.Column, cseS.Value._styleId);
                    }
                }
            }

            //Copy outline
            var styleRowOutlineLevel = ws.Row(copyStylesFromRow + rows).OutlineLevel;
            for (var r = rowFrom; r < rowFrom + rows; r++)
            {
                ws.Row(r).OutlineLevel = styleRowOutlineLevel;
            }
        }
        private static void InsertTableColumns(int columnFrom, int columns, ExcelTable tbl)
        {
            var node = tbl.Columns[0].TopNode.ParentNode;
            var ix = columnFrom - tbl.Address.Start.Column - 1;
            var insPos = node.ChildNodes[ix];
            ix += 2;
            for (int i = 0; i < columns; i++)
            {
                var name =
                    tbl.Columns.GetUniqueName(string.Format("Column{0}",
                        (ix++).ToString(CultureInfo.InvariantCulture)));
                XmlElement tableColumn =
                    (XmlElement)tbl.TableXml.CreateNode(XmlNodeType.Element, "tableColumn", ExcelPackage.schemaMain);
                tableColumn.SetAttribute("id", (tbl.Columns.Count + i + 1).ToString(CultureInfo.InvariantCulture));
                tableColumn.SetAttribute("name", name);
                insPos = node.InsertAfter(tableColumn, insPos);
            } //Create tbl Column
            tbl._cols = new ExcelTableColumnCollection(tbl);
        }

    }
    #endregion
}

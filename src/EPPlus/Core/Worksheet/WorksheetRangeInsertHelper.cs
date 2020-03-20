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
using OfficeOpenXml.ConditionalFormatting;
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

                FixFormulasInsertRow(ws, rowFrom, rows);

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

                WorksheetRangeHelper.AdjustDrawingsRow(ws, rowFrom, rows);
            }
        }
        internal static void InsertColumn(ExcelWorksheet ws, int columnFrom, int columns, int copyStylesFromColumn)
        {
            ValidateInsertColumn(ws, columnFrom, columns);

            lock (ws)
            {
                InsertCellStores(ws, 0, columnFrom, 0, columns);

                FixFormulasInsertColumn(ws, columnFrom, columns);

                WorksheetRangeHelper.FixMergedCellsColumn(ws, columnFrom, columns, false);

                AdjustColumns(ws, columnFrom, columns);

                CopyStylesFromColumn(ws, columnFrom, columns, copyStylesFromColumn);

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

                //Adjust drawing positions.
                WorksheetRangeHelper.AdjustDrawingsColumn(ws, columnFrom, columns);
            }
        }
        internal static void Insert(ExcelRangeBase range, eShiftTypeInsert shift)
        {
            WorksheetRangeHelper.ValidateIfInsertDeleteIsPossible(range, GetEffectedRange(range, shift, 1));

            var ws = range.Worksheet;
            var styleList = GetStylesForRange(range, shift);
            WorksheetRangeHelper.ConvertEffectedSharedFormulasToCellFormulas(ws, range);

            if (shift == eShiftTypeInsert.Down)
            {
                InsertCellStores(range._worksheet, range._fromRow, range._fromCol, range.Rows, range.Columns, range._toCol);
            }
            else
            {
                InsertCellStoreShiftRight(range._worksheet, range);
            }
            var effectedAddress = GetEffectedRange(range, shift);
            FixFormulasInsert(range, effectedAddress, shift);

            WorksheetRangeHelper.FixMergedCells(ws, range, shift);

            SetStylesForRange(range, shift, styleList);

            InsertTableAddresses(ws, range, shift, effectedAddress);
            InsertPivottableAddresses(ws, range, shift, effectedAddress);

            //Update data validation references
            foreach (var dv in ws.DataValidations)
            {
                ((ExcelDataValidation)dv).SetAddress(InsertSplitAddress(dv.Address, range, effectedAddress, shift).Address);
            }

            //Update Conditional formatting references
            foreach (var cf in ws.ConditionalFormatting)
            {
                ((ExcelConditionalFormattingRule)cf).Address = new ExcelAddress(InsertSplitAddress(cf.Address, range, effectedAddress, shift).Address);
            }

            if(shift==eShiftTypeInsert.Down)
            {
                WorksheetRangeHelper.AdjustDrawingsRow(ws, range._fromRow, range.Rows, range._fromCol, range._toCol);
            }
            else
            {
                WorksheetRangeHelper.AdjustDrawingsColumn(ws, range._fromCol, range.Columns, range._fromRow, range._toRow);
            }
        }

        private static ExcelAddressBase InsertSplitAddress(ExcelAddressBase address, ExcelAddressBase range, ExcelAddressBase effectedAddress, eShiftTypeInsert shift)
        {
            var collide = effectedAddress.Collide(address);
            if (collide == ExcelAddressBase.eAddressCollition.Partly)
            {
                var addressToShift = effectedAddress.Intersect(address);
                var shiftedAddress = ShiftAddress(addressToShift, range, shift);
                var newAddress = "";
                if(address._fromRow < addressToShift._fromRow)
                {
                    newAddress = ExcelCellBase.GetAddress(address._fromRow, address._fromCol, addressToShift._fromRow-1, address._toCol)+",";
                }
                if (address._fromCol < addressToShift._fromCol)
                {
                    var fromRow = Math.Max(address._fromRow, addressToShift._fromRow);
                    newAddress += ExcelCellBase.GetAddress(fromRow, address._fromCol, address._toRow, addressToShift._fromCol-1)+ ",";
                }

                newAddress += $"{shiftedAddress},";

                if (address._toRow > addressToShift._toRow)
                {
                    newAddress += ExcelCellBase.GetAddress(addressToShift._toRow + 1, address._fromCol, address._toRow, address._toCol)+",";
                }
                if (address._toCol > addressToShift._toCol)
                {
                    newAddress += ExcelCellBase.GetAddress(address._fromRow, addressToShift._toCol + 1, address._toRow, address._toCol)+",";
                }
                return new ExcelAddressBase(newAddress.Substring(0, newAddress.Length-1));
            }
            else if(collide!=ExcelAddressBase.eAddressCollition.No)
            {
                return ShiftAddress(address, range, shift);
            }
            return address;
        }

        private static ExcelAddressBase ShiftAddress(ExcelAddressBase address, ExcelAddressBase range, eShiftTypeInsert shift)
        {
            if (shift == eShiftTypeInsert.Down)
            {
                return address.AddRow(range._fromRow, range.Rows);
            }
            else
            {
                return address.AddColumn(range._fromCol, range.Columns);
            }
        }

        private static void InsertPivottableAddresses(ExcelWorksheet ws, ExcelRangeBase range, eShiftTypeInsert shift, ExcelAddressBase effectedAddress)
        {
            foreach (var ptbl in ws.PivotTables)
            {
                if (shift == eShiftTypeInsert.Down)
                {
                    if (ptbl.Address._fromCol >= range._fromCol && ptbl.Address._toCol <= range._toCol)
                    {
                        ptbl.Address = ptbl.Address.AddRow(range._fromRow, range.Rows);
                    }
                }
                else
                {
                    if (ptbl.Address._fromRow >= range._fromRow && ptbl.Address._toRow <= range._toRow)
                    {
                        ptbl.Address = ptbl.Address.AddColumn(range._fromCol, range.Columns);
                    }
                }

                if (ptbl.CacheDefinition.SourceRange.Worksheet==ws)
                {
                    var address = ptbl.CacheDefinition.SourceRange;
                    if (shift == eShiftTypeInsert.Down)
                    {
                        if (address._fromCol >= range._fromCol && address._toCol <= range._toCol)
                        {
                            ptbl.CacheDefinition.SourceRange = ws.Cells[address.AddRow(range._fromRow, range.Rows).Address];
                        }
                    }
                    else
                    {
                        if (address._fromRow >= range._fromRow && address._toRow <= range._toRow)
                        {
                            ptbl.CacheDefinition.SourceRange = ws.Cells[address.AddColumn(range._fromCol, range.Columns).Address];
                        }
                    }
                }
            }
        }

        private static void InsertTableAddresses(ExcelWorksheet ws, ExcelRangeBase range, eShiftTypeInsert shift, ExcelAddressBase effectedAddress)
        {
            foreach (var tbl in ws.Tables)
            {               
                if (shift == eShiftTypeInsert.Down)
                {
                    if (tbl.Address._fromCol >= range._fromCol && tbl.Address._toCol <= range._toCol)
                    {
                        tbl.Address = tbl.Address.AddRow(range._fromRow, range.Rows);
                    }
                }
                else
                {
                    if (tbl.Address._fromRow >= range._fromRow && tbl.Address._toRow <= range._toRow)
                    {
                        tbl.Address = tbl.Address.AddColumn(range._fromCol, range.Columns);
                    }
                }
            }
        }

        private static List<int> GetStylesForRange(ExcelRangeBase range, eShiftTypeInsert shift)
        {
            var list=new List<int>();
            if(shift==eShiftTypeInsert.Down)
            {
                for(int i=0;i<range.Columns;i++)
                {
                    if(range._fromRow == 1)
                    {
                        list.Add(0);
                    }
                    else
                    {
                        list.Add(range.Offset(-1, i).StyleID);
                    }
                }
            }
            else
            {
                for (int i = 0; i < range.Rows; i++)
                {
                    if (range._fromCol == 1)
                    {
                        list.Add(0);
                    }
                    else
                    {
                        list.Add(range.Offset(i, -1).StyleID);
                    }
                }
            }
            return list;
        }

        private static void SetStylesForRange(ExcelRangeBase range, eShiftTypeInsert shift, List<int> list)
        {
            if (shift == eShiftTypeInsert.Down)
            {
                for (int i = 0; i < range.Columns; i++)
                {
                    range.Offset(0, i,range.Rows,1).StyleID=list[i];
                }
            }
            else
            {
                for (int i = 0; i < range.Rows; i++)
                {
                    
                    range.Offset(i, 0, 1, range.Columns).StyleID = list[i];
                }
            }
        }

        private static ExcelAddressBase GetEffectedRange(ExcelRangeBase range, eShiftTypeInsert shift, int? start=null)
        {
            if (shift == eShiftTypeInsert.Down)
            {                
                return new ExcelAddressBase(start ?? range._fromRow, range._fromCol, ExcelPackage.MaxRows, range._toCol);
            }
            else if (shift == eShiftTypeInsert.Right)
            {
                return new ExcelAddressBase(range._fromRow, start ?? range._fromCol, range._toRow, ExcelPackage.MaxColumns);
            }
            else if (shift == eShiftTypeInsert.EntireColumn)
            {
                return new ExcelAddressBase(1, range._fromCol, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
            }
            else
            {
                return new ExcelAddressBase(range._fromRow, 1, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
            }
        }

        private static void CopyStylesFromColumn(ExcelWorksheet ws, int columnFrom, int columns, int copyStylesFromColumn)
        {
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
        }

        private static void AdjustColumns(ExcelWorksheet ws, int columnFrom, int columns)
        {
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
        }
        private static void FixFormulasInsert(ExcelRangeBase range, ExcelAddressBase effectedAddress, eShiftTypeInsert shift)
        {
            //Adjust formulas
            foreach (var ws in range._workbook.Worksheets)
            {
                var workSheetName = range.Worksheet.Name;
                var rowFrom = range._fromRow;
                var columnFrom = range._fromCol;
                var rows = range.Rows;

                foreach (var f in ws._sharedFormulas.Values)
                {
                    if (workSheetName == ws.Name)
                    {
                        if (f.StartCol >= columnFrom)
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
                            f.Formula = ExcelCellBase.UpdateFormulaReferences(f.Formula, range, effectedAddress, shift, ws.Name, workSheetName);
                        }
                    }
                    else if (f.Formula.Contains(workSheetName))
                    {
                        f.Formula = ExcelCellBase.UpdateFormulaReferences(f.Formula, range, effectedAddress, shift, ws.Name, workSheetName);
                    }
                }

                var cse = new CellStoreEnumerator<object>(ws._formulas);
                while (cse.Next())
                {
                    if (cse.Value is string v)
                    {
                        if (workSheetName == ws.Name)
                        {
                            cse.Value = ExcelCellBase.UpdateFormulaReferences(v, range, effectedAddress, shift, ws.Name, workSheetName);
                        }
                        else if (v.Contains(workSheetName))
                        {
                            cse.Value = ExcelCellBase.UpdateFormulaReferences(v, range, effectedAddress, shift, ws.Name, workSheetName);
                        }
                    }
                }
            }
        }

        private static void FixFormulasInsertRow(ExcelWorksheet ws, int rowFrom, int rows, int columnFrom=0, int columnTo=ExcelPackage.MaxColumns)
        {
            //Adjust formulas
            foreach (var wsToUpdate in ws.Workbook.Worksheets)
            {
                foreach (var f in wsToUpdate._sharedFormulas.Values)
                {
                    if (ws.Name == wsToUpdate.Name)
                    {
                        if (f.StartCol >= columnFrom)
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
                            f.Formula = ExcelCellBase.UpdateFormulaReferences(f.Formula, rows, 0, rowFrom, 0, wsToUpdate.Name, ws.Name);
                        }
                    }
                    else if (f.Formula.Contains(ws.Name))
                    {
                        f.Formula = ExcelCellBase.UpdateFormulaReferences(f.Formula, rows, 0, rowFrom, columnFrom, wsToUpdate.Name, ws.Name);
                    }
                }

                var cse = new CellStoreEnumerator<object>(wsToUpdate._formulas);
                while (cse.Next())
                {
                    if (cse.Value is string v)
                    {
                        if (ws.Name == wsToUpdate.Name)
                        {
                            cse.Value = ExcelCellBase.UpdateFormulaReferences(v, rows, 0, rowFrom, 0, wsToUpdate.Name, ws.Name);
                        }
                        else if (v.Contains(ws.Name))
                        {
                            cse.Value = ExcelCellBase.UpdateFormulaReferences(v, rows, 0, rowFrom, 0, wsToUpdate.Name, ws.Name);
                        }
                    }
                }
}        }
        private static void FixFormulasInsertColumn(ExcelWorksheet ws, int columnFrom, int columns)
        {
            foreach (var wsToUpdate in ws.Workbook.Worksheets)
            {
                foreach (var f in wsToUpdate._sharedFormulas.Values)
                {
                    if (ws.Name == wsToUpdate.Name)
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
                        f.Formula = ExcelCellBase.UpdateFormulaReferences(f.Formula, 0, columns, 0, columnFrom, wsToUpdate.Name, ws.Name);
                    }
                    else if (f.Formula.Contains(ws.Name))
                    {
                        f.Formula = ExcelCellBase.UpdateFormulaReferences(f.Formula, 0, columns, 0, columnFrom, wsToUpdate.Name, ws.Name);
                    }
                }

                var cse = new CellStoreEnumerator<object>(wsToUpdate._formulas);
                while (cse.Next())
                {
                    if (cse.Value is string v)
                    {
                        if (ws.Name == wsToUpdate.Name)
                        {
                            cse.Value = ExcelCellBase.UpdateFormulaReferences(v, 0, columns, 0, columnFrom, wsToUpdate.Name, ws.Name);
                        }
                        else if (v.Contains(ws.Name))
                        {
                            cse.Value = ExcelCellBase.UpdateFormulaReferences(v, 0, columns, 0, columnFrom, wsToUpdate.Name, ws.Name);
                        }
                    }
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
        internal static void InsertCellStores(ExcelWorksheet ws, int rowFrom, int columnFrom, int rows, int columns, int columnTo=ExcelPackage.MaxColumns)
        {
            ws._values.Insert(rowFrom, columnFrom, rows, columns);
            ws._formulas.Insert(rowFrom, columnFrom, rows, columns);
            ws._commentsStore.Insert(rowFrom, columnFrom, rows, columns);
            ws._hyperLinks.Insert(rowFrom, columnFrom, rows, columns);
            ws._flags.Insert(rowFrom, columnFrom, rows, columns);
            ws._vmlDrawings?._drawings.Insert(rowFrom, columnFrom, rows, columns);

            if(rows==0||columns==0)
            {
                ws.Comments.Insert(rowFrom, columnFrom, rows, columns);
                ws._names.Insert(rowFrom, columnFrom, rows, columns, 0, columnTo);
                ws.Workbook.Names.Insert(rowFrom, columnFrom, rows, columns, n => n.Worksheet == ws, 0, columnTo);
            }
            else
            {
                ws.Comments.Insert(rowFrom, columnFrom, rows, 0, 0, columnTo);
                ws._names.Insert(rowFrom, columnFrom, rows, 0, columnFrom, columnTo);
                ws.Workbook.Names.Insert(rowFrom, columnFrom, rows, 0, n => n.Worksheet == ws, columnFrom, columnTo);
            }
        }
        internal static void InsertCellStoreShiftRight(ExcelWorksheet ws, ExcelAddressBase fromAddress)
        {
            ws._values.InsertShiftRight(fromAddress);
            ws._formulas.InsertShiftRight(fromAddress);
            ws._commentsStore.InsertShiftRight(fromAddress);
            ws._hyperLinks.InsertShiftRight(fromAddress);
            ws._flags.InsertShiftRight(fromAddress);
            ws._vmlDrawings?._drawings.InsertShiftRight(fromAddress);

            ws.Comments.Insert(fromAddress._fromRow, fromAddress._fromCol, 0, fromAddress.Columns);
            ws._names.Insert(fromAddress._fromRow, fromAddress._fromCol, 0, fromAddress.Columns, fromAddress._fromRow, fromAddress._toRow);
            ws.Workbook.Names.Insert(fromAddress._fromRow, fromAddress._fromCol, 0, fromAddress.Columns, n => n.Worksheet == ws, fromAddress._fromRow, fromAddress._toRow);
        }

        private static void CopyFromStyleRow(ExcelWorksheet ws, int rowFrom, int rows, int copyStylesFromRow)
        {
            if (copyStylesFromRow >= rowFrom) copyStylesFromRow += rows;

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
            var styleRowOutlineLevel = ws.Row(copyStylesFromRow).OutlineLevel;
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

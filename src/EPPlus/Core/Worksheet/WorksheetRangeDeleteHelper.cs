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
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
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
                var delRange = new ExcelAddressBase(rowFrom, 1, rowFrom + rows - 1, ExcelPackage.MaxColumns);
                WorksheetRangeHelper.ConvertEffectedSharedFormulasToCellFormulas(ws, delRange);

                DeleteCellStores(ws, rowFrom, 0, rows, ExcelPackage.MaxColumns + 1);

                foreach (var wsToUpdate in ws.Workbook.Worksheets)
                {
                    FixFormulasDeleteRow(wsToUpdate, rowFrom, rows, ws.Name);
                }


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
                var delRange = new ExcelAddressBase(1, columnFrom, ExcelPackage.MaxRows, columnFrom + columns - 1);
                WorksheetRangeHelper.ConvertEffectedSharedFormulasToCellFormulas(ws, delRange);

                DeleteCellStores(ws, 0, columnFrom, 0, columns);

                foreach (var wsToUpdate in ws.Workbook.Worksheets)
                {
                    FixFormulasDeleteColumn(wsToUpdate, columnFrom, columns, ws.Name);
                }                
                
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

        private static void DeleteCellStores(ExcelWorksheet ws, int rowFrom, int columnFrom, int rows, int columns, int columnTo = ExcelPackage.MaxColumns)
        {
            //Store
            ws._values.Delete(rowFrom, columnFrom, rows, columns, true);
            ws._formulas.Delete(rowFrom, columnFrom, rows, columns, true);
            ws._flags.Delete(rowFrom, columnFrom, rows, columns, true);
            ws._commentsStore.Delete(rowFrom, columnFrom, rows, columns, true);
            ws._vmlDrawings?._drawings.Delete(rowFrom, columnFrom, rows, columns, true);
            ws._hyperLinks.Delete(rowFrom, columnFrom, rows, columns, true);

            if (rows == 0 || columns == 0)
            {
                ws._names.Delete(rowFrom, columnFrom, rows, columns);
                ws.Workbook.Names.Delete(rowFrom, columnFrom, rows, columns, n => n.Worksheet == ws);
                ws.Comments.Delete(rowFrom, columnFrom, rows, columns);

                if (rowFrom == 0 && rows >= ExcelPackage.MaxRows) //Delete full column
                {
                    AdjustColumnMinMax(ws, columnFrom, columns);
                }
            }
            else
            {
                ws.Comments.Delete(rowFrom, columnFrom, rows, 0, 0, columnTo);
                ws._names.Delete(rowFrom, columnFrom, rows, 0, columnFrom, columnTo);
                ws.Workbook.Names.Delete(rowFrom, columnFrom, rows, 0, n => n.Worksheet == ws, columnFrom, columnTo);
            }
        }
        private static void DeleteCellStoresShiftLeft(ExcelWorksheet ws, ExcelRangeBase fromAddress)
        {
            //Store
            ws._values.DeleteShiftLeft(fromAddress);
            ws._formulas.DeleteShiftLeft(fromAddress);
            ws._flags.DeleteShiftLeft(fromAddress);
            ws._commentsStore.DeleteShiftLeft(fromAddress);
            ws._vmlDrawings?._drawings.DeleteShiftLeft(fromAddress);
            ws._hyperLinks.DeleteShiftLeft(fromAddress);

            ws.Comments.Delete(fromAddress._fromRow, fromAddress._fromCol, 0, fromAddress.Columns);
            ws._names.Delete(fromAddress._fromRow, fromAddress._fromCol, 0, fromAddress.Columns, fromAddress._fromRow, fromAddress._toRow);
            ws.Workbook.Names.Delete(fromAddress._fromRow, fromAddress._fromCol, 0, fromAddress.Columns, n => n.Worksheet == ws, fromAddress._fromRow, fromAddress._toRow);
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
        static void FixFormulasDeleteRow(ExcelWorksheet ws, int rowFrom, int rows, string workSheetName)
        {
            var delSF = new List<int>();
            foreach (var sf in ws._sharedFormulas.Values)
            {
                if (workSheetName == ws.Name)
                {
                    var a = new ExcelAddress(sf.Address).DeleteRow(rowFrom, rows);
                    if (a == null)
                    {
                        delSF.Add(sf.Index);
                    }
                    else
                    {
                        sf.Address = a.Address;
                        sf.Formula = ExcelCellBase.UpdateFormulaReferences(sf.Formula, -rows, 0, rowFrom, 0, ws.Name, workSheetName);
                        if (sf.StartRow >= rowFrom)
                        {
                            var r = Math.Max(rowFrom, sf.StartRow - rows);
                            sf.StartRow = r;
                        }
                    }
                }
                else if (sf.Formula.Contains(workSheetName))
                {
                    sf.Formula = ExcelCellBase.UpdateFormulaReferences(sf.Formula, -rows, 0, rowFrom, 0, ws.Name, workSheetName);
                }
            }

            foreach (var ix in delSF)
            {
                ws._sharedFormulas.Remove(ix);
            }
            var cse = new CellStoreEnumerator<object>(ws._formulas, 1, 1, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
            while (cse.Next())
            {
                if (cse.Value is string v)
                {
                    if (workSheetName == ws.Name || v.IndexOf(workSheetName, StringComparison.CurrentCultureIgnoreCase)>=0)
                    {
                        cse.Value = ExcelCellBase.UpdateFormulaReferences(v, -rows, 0, rowFrom, 0, ws.Name, workSheetName);
                    }
                }
            }
        }
        internal static void FixFormulasDeleteColumn(ExcelWorksheet ws, int columnFrom, int columns, string workSheetName)
        {
            var delSF = new List<int>();
            foreach (var sf in ws._sharedFormulas.Values)
            {

                if (workSheetName == ws.Name)
                {
                    var a = new ExcelAddress(sf.Address).DeleteColumn(columnFrom, columns);
                    if (a == null)
                    {
                        delSF.Add(sf.Index);
                    }
                    else
                    {
                        sf.Address = a.Address;
                        sf.Formula = ExcelCellBase.UpdateFormulaReferences(sf.Formula, 0, -columns, 0, columnFrom, ws.Name, workSheetName);

                        if (sf.StartCol > columnFrom)
                        {
                            var c = Math.Max(columnFrom, sf.StartCol - columns);
                            sf.StartCol -= c;
                        }
                    }
                }
                else if (sf.Formula.Contains(workSheetName))
                {
                    sf.Formula = ExcelCellBase.UpdateFormulaReferences(sf.Formula, 0, -columns, 0, columnFrom, ws.Name, workSheetName);
                }
            }
            foreach (var ix in delSF)
            {
                ws._sharedFormulas.Remove(ix);
            }

            var cse = new CellStoreEnumerator<object>(ws._formulas, 1, 1, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
            while (cse.Next())
            {
                if (cse.Value is string v)
                {
                    if (workSheetName == ws.Name || v.IndexOf(workSheetName, StringComparison.CurrentCultureIgnoreCase)>=0)
                    {
                        cse.Value = ExcelCellBase.UpdateFormulaReferences(v, 0, -columns, 0, columnFrom, ws.Name, workSheetName);
                    }
                }
            }
        }

        internal static void Delete(ExcelRangeBase range, eShiftTypeDelete shift)
        {
            if (shift == eShiftTypeDelete.Left)
            {
                ValidateColumn(range._fromCol, range.Columns);
            }
            else
            {
                ValidateRow(range._fromRow, range.Rows);
            }

            var effectedAddress = GetEffectedRange(range, shift);
            WorksheetRangeHelper.ValidateIfInsertDeleteIsPossible(range, effectedAddress, GetEffectedRange(range, shift, 1));

            var ws = range.Worksheet;
            WorksheetRangeHelper.ConvertEffectedSharedFormulasToCellFormulas(ws, effectedAddress);
            if (shift == eShiftTypeDelete.Up)
            {
                DeleteCellStores(ws, range._fromRow, range._fromCol, range.Rows, range.Columns, range._toCol);
            }
            else
            {
                DeleteCellStoresShiftLeft(ws, range);
            }

            FixFormulasDelete(range, effectedAddress, shift);
            WorksheetRangeHelper.FixMergedCells(ws, range, shift);

            DeleteTableAddresses(ws, range, shift, effectedAddress);
            DeletePivottableAddresses(ws, range, shift, effectedAddress);

            //Adjust/delete data validations and conditional formatting
            DeleteDataValidations(range, shift, ws, effectedAddress);
            DeleteConditionalForatting(range, shift, ws, effectedAddress);

            AdjustDrawings(range, shift);
        }

        private static void AdjustDrawings(ExcelRangeBase range, eShiftTypeDelete shift)
        {
            if (shift == eShiftTypeDelete.Up)
            {
                WorksheetRangeHelper.AdjustDrawingsRow(range.Worksheet, range._fromRow, -range.Rows, range._fromCol, range._toCol);
            }
            else
            {
                WorksheetRangeHelper.AdjustDrawingsColumn(range.Worksheet, range._fromCol, -range.Columns, range._fromRow, range._toRow);
            }
        }

        private static void DeleteConditionalForatting(ExcelRangeBase range, eShiftTypeDelete shift, ExcelWorksheet ws, ExcelAddressBase effectedAddress)
        {
            //Update Conditional formatting references
            var deletedCF = new List<IExcelConditionalFormattingRule>();
            foreach (var cf in ws.ConditionalFormatting)
            {
                var address = DeleteSplitAddress(cf.Address, range, effectedAddress, shift);
                if (address == null)
                {
                    deletedCF.Add(cf);
                }
                else
                {
                    ((ExcelConditionalFormattingRule)cf).Address = new ExcelAddress(address.Address);
                }
            }
            deletedCF.ForEach(cf => ws.ConditionalFormatting.Remove(cf));
        }

        private static void DeleteDataValidations(ExcelRangeBase range, eShiftTypeDelete shift, ExcelWorksheet ws, ExcelAddressBase effectedAddress)
        {
            //Update data validation references
            var deletedDV = new List<IExcelDataValidation>();
            foreach (var dv in ws.DataValidations)
            {
                var address = DeleteSplitAddress(dv.Address, range, effectedAddress, shift);
                if (address == null)
                {
                    deletedDV.Add(dv);
                }
                else
                {
                    ((ExcelDataValidation)dv).SetAddress(address.Address);
                }
            }
            deletedDV.ForEach(dv => ws.DataValidations.Remove(dv));
        }

        private static ExcelAddressBase DeleteSplitAddress(ExcelAddressBase address, ExcelAddressBase range, ExcelAddressBase effectedAddress, eShiftTypeDelete shift)
        {
            var collide = effectedAddress.Collide(address);
            if (collide == ExcelAddressBase.eAddressCollition.Partly)
            {
                var addressToShift = effectedAddress.Intersect(address);
                var shiftedAddress = ShiftAddress(addressToShift, range, shift);
                var newAddress = "";
                if (address._fromRow < addressToShift._fromRow)
                {
                    newAddress = ExcelCellBase.GetAddress(address._fromRow, address._fromCol, addressToShift._fromRow - 1, address._toCol) + ",";
                }
                if (address._fromCol < addressToShift._fromCol)
                {
                    var fromRow = Math.Max(address._fromRow, addressToShift._fromRow);
                    newAddress += ExcelCellBase.GetAddress(fromRow, address._fromCol, address._toRow, addressToShift._fromCol - 1) + ",";
                }
                
                if(shiftedAddress != null)
                {
                    newAddress += $"{shiftedAddress.Address},";
                }

                if (address._toRow > addressToShift._toRow)
                {
                    newAddress += ExcelCellBase.GetAddress(addressToShift._toRow + 1, address._fromCol, address._toRow, address._toCol) + ",";
                }
                if (address._toCol > addressToShift._toCol)
                {
                    newAddress += ExcelCellBase.GetAddress(address._fromRow, addressToShift._toCol + 1, address._toRow, address._toCol) + ",";
                }
                return new ExcelAddressBase(newAddress.Substring(0, newAddress.Length - 1));
            }
            else if (collide != ExcelAddressBase.eAddressCollition.No)
            {
                return ShiftAddress(address, range, shift);
            }
            return address;
        }

        private static ExcelAddressBase ShiftAddress(ExcelAddressBase address, ExcelAddressBase range, eShiftTypeDelete shift)
        {
            if (shift == eShiftTypeDelete.Up)
            {
                return address.DeleteRow(range._fromRow, range.Rows);
            }
            else
            {
                return address.DeleteColumn(range._fromCol, range.Columns);
            }
        }
        private static void DeletePivottableAddresses(ExcelWorksheet ws, ExcelRangeBase range, eShiftTypeDelete shift, ExcelAddressBase effectedAddress)
        {
            var deletedPt = new List<ExcelPivotTable>();
            foreach (var ptbl in ws.PivotTables)
            {
                if (shift == eShiftTypeDelete.Up)
                {
                    if (ptbl.Address._fromCol >= range._fromCol && ptbl.Address._toCol <= range._toCol)
                    {
                        ptbl.Address = ptbl.Address.DeleteRow(range._fromRow, range.Rows);
                    }
                }
                else
                {
                    if (ptbl.Address._fromRow >= range._fromRow && ptbl.Address._toRow <= range._toRow)
                    {
                        ptbl.Address = ptbl.Address.DeleteColumn(range._fromCol, range.Columns);
                    }
                }
                if (ptbl.Address == null)
                {
                    deletedPt.Add(ptbl);
                }
                else
                {
                    foreach (var wsSource in ws.Workbook.Worksheets)
                    {
                        if (ptbl.CacheDefinition.SourceRange.Worksheet == wsSource)
                        {
                            var address = ptbl.CacheDefinition.SourceRange;
                            if (shift == eShiftTypeDelete.Up)
                            {
                                if (address._fromCol >= range._fromCol && address._toCol <= range._toCol)
                                {
                                    var deletedRange = ws.Cells[address.DeleteRow(range._fromRow, range.Rows).Address];
                                    if (deletedRange != null)
                                    {
                                        ptbl.CacheDefinition.SourceRange = deletedRange;
                                    }
                                }
                            }
                            else
                            {
                                if (address._fromRow >= range._fromRow && address._toRow <= range._toRow)
                                {
                                    var deletedRange = ws.Cells[address.DeleteColumn(range._fromCol, range.Columns).Address];
                                    if (deletedRange != null)
                                    {
                                        ptbl.CacheDefinition.SourceRange = deletedRange;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            deletedPt.ForEach(x => ws.PivotTables.Delete(x));

        }

        private static void DeleteTableAddresses(ExcelWorksheet ws, ExcelRangeBase range, eShiftTypeDelete shift, ExcelAddressBase effectedAddress)
        {
            var deletedTbl = new List<ExcelTable>();
            foreach (var tbl in ws.Tables)
            {
                if (shift == eShiftTypeDelete.Up)
                {
                    if (tbl.Address._fromCol >= range._fromCol && tbl.Address._toCol <= range._toCol)
                    {
                        tbl.Address = tbl.Address.DeleteRow(range._fromRow, range.Rows);
                    }
                }
                else
                {
                    if (tbl.Address._fromRow >= range._fromRow && tbl.Address._toRow <= range._toRow)
                    {
                        tbl.Address = tbl.Address.DeleteColumn(range._fromCol, range.Columns);
                    }
                }
                if(tbl.Address==null) deletedTbl.Add(tbl);
            }

            deletedTbl.ForEach(x => ws.Tables.Delete(x));
        }
        private static void FixFormulasDelete(ExcelRangeBase range, ExcelAddressBase effectedRange, eShiftTypeDelete shift)
        {
            foreach (var ws in range.Worksheet.Workbook.Worksheets)
            {
                var workSheetName = range.WorkSheetName;
                var rowFrom = range._fromRow;
                var rows = range.Rows;

                var delSF = new List<int>();
                foreach (var sf in ws._sharedFormulas.Values)
                {
                    if (workSheetName == ws.Name)
                    {
                        if (effectedRange.Collide(new ExcelAddressBase(sf.Address)) != ExcelAddressBase.eAddressCollition.No)
                        {
                            ExcelAddressBase a;
                            if (shift == eShiftTypeDelete.Up)
                            {
                                a = new ExcelAddress(sf.Address).DeleteRow(range._fromRow, rows);
                            }
                            else
                            {
                                a = new ExcelAddress(sf.Address).DeleteColumn(range._fromRow, rows);
                            }

                            if (a == null)
                            {
                                delSF.Add(sf.Index);
                            }
                            else
                            {
                                sf.Address = a.Address;
                                sf.Formula = ExcelCellBase.UpdateFormulaReferences(sf.Formula, range, effectedRange, shift, ws.Name, workSheetName);
                                if (sf.StartRow >= rowFrom)
                                {
                                    var r = Math.Max(rowFrom, sf.StartRow - rows);
                                    sf.StartRow = r;
                                }
                            }
                        }
                    }
                    else if (sf.Formula.Contains(workSheetName))
                    {
                        sf.Formula = ExcelCellBase.UpdateFormulaReferences(sf.Formula, -rows, 0, rowFrom, 0, ws.Name, workSheetName);
                    }
                }

                foreach (var ix in delSF)
                {
                    ws._sharedFormulas.Remove(ix);
                }
                var cse = new CellStoreEnumerator<object>(ws._formulas, 1, 1, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
                while (cse.Next())
                {
                    if (cse.Value is string v)
                    {
                        if (workSheetName == ws.Name || v.IndexOf(workSheetName, StringComparison.CurrentCultureIgnoreCase) >= 0)
                        {
                            cse.Value = ExcelCellBase.UpdateFormulaReferences(v, range, effectedRange, shift, ws.Name, workSheetName);
                        }
                    }
                }
            }
        }

        private static ExcelAddressBase GetEffectedRange(ExcelRangeBase range, eShiftTypeDelete shift, int? start = null)
        {
            if (shift == eShiftTypeDelete.Up)
            {
                return new ExcelAddressBase(start ?? range._fromRow, range._fromCol, ExcelPackage.MaxRows, range._toCol);
            }
            else if (shift == eShiftTypeDelete.Left)
            {
                return new ExcelAddressBase(range._fromRow, start ?? range._fromCol, range._toRow, ExcelPackage.MaxColumns);
            }
            else if (shift == eShiftTypeDelete.EntireColumn)
            {
                return new ExcelAddressBase(1, range._fromCol, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
            }
            else
            {
                return new ExcelAddressBase(range._fromRow, 1, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
            }
        }

    }
}

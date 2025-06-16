using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;


namespace OfficeOpenXml.Utils
{
    internal static class ValueFinder
    {
        internal static List<int> FirstValueCell(ExcelWorksheet sheet, FormulaRangeAddress address)
        {
            var fromRow = address.FromRow;
            var fromCol = address.FromCol;
            var someValue = sheet._values.GetValue(address.FromRow, address.FromCol);
            var valueOfValue = sheet._values.GetValue(address.FromRow, address.FromCol)._value;
            if (valueOfValue == null)
            {
                while (sheet._values.NextCell(ref fromRow, ref fromCol, address.FromRow, address.FromCol, address.ToRow, address.ToCol))
                {
                    if (sheet._values.GetValue(fromRow, fromCol)._value != null)
                    {
                        return new List<int> { fromRow, fromCol - 1 };
                    }
                }
            }
            return new List<int> { fromRow, fromCol };
        }

        internal static List<int> FirstValueCell(FormulaRangeAddress address, CellStore<object> values)
        {
            var fromRow = address.FromRow;
            var fromCol = address.FromCol;

            var valueStart = values.GetValue(fromRow, fromCol);
            if (valueStart == null)
            {
                while (values.NextCell(ref fromRow, ref fromCol, address.FromRow, address.ToCol - address.FromCol, address.ToRow, address.ToCol))
                {
                    if (values.GetValue(fromRow, fromCol) != null)
                    {
                        return new List<int> { fromRow, fromCol - 1 };
                    }
                }
            }
            return new List<int> { fromRow, fromCol };
        }

        internal static List<int> LastValueCell(FormulaRangeAddress address, CellStore<object> values)
        {
            var toRow = address.ToRow;
            var toCol = address.ToCol;
            var firstVal = values.GetValue(toRow, toCol);
            if (firstVal == null)
            {
                while (values.PrevCell(ref toRow, ref toCol) && toRow > 0)
                {
                    if (toCol > 0 && values.GetValue(toRow, toCol) != null)
                    {
                        return new List<int> { toRow, toCol };
                    }
                }
                return null;
            }
            return new List<int> { toRow, toCol };
        }

        internal static List<int> LastValueCell(ExcelWorksheet sheet, FormulaRangeAddress address)
        {
            //The range might refer to cells outside the worksheet dimension if no value has been set.
            //Therefore take the lower value between toRow and toCol and the dimension version of them
            var toRow = sheet.Dimension._toRow < address.ToRow ? sheet.Dimension._toRow : address.ToRow;
            var toCol = sheet.Dimension._toCol < address.ToCol ? sheet.Dimension._toCol : address.ToCol;

            if (sheet._values.GetValue(toRow, toCol)._value == null)
            {
                while (sheet._values.PrevCell(ref toRow, ref toCol) && toRow > 0)
                {
                    if (toCol > 0 && sheet._values.GetValue(toRow, toCol)._value != null)
                    {
                        return new List<int> { toRow, toCol };
                    }
                }
                return null;
            }
            return new List<int> { toRow, toCol };
        }

        internal static IRangeInfo RangeByValue(IRangeInfo rInfo)
        {
            //int fromRow = -1;
            //int fromCol = -1;
            //int toRow = -1;
            //int toCol = -1;

            List<int> fvc;
            List<int> lvc;

            var sheet = rInfo.Worksheet;
            var address = rInfo.Address;

            CellStore<object> csValues;

            var baseAddress = rInfo.Address.ToExcelAddressBase();
            if (baseAddress.IsExternal)
            {
                var extRangeInfo = (EpplusExcelExternalRangeInfo)rInfo;
                
                var cellValues = extRangeInfo?._externalWs.CellValues._values;
                var rAddress = rInfo.Address;

                if (extRangeInfo._externalWs != null)
                {
                    var dimension = extRangeInfo._externalWs.GetDimension();
                    rAddress.ToRow = dimension._toRow < address.ToRow ? dimension._toRow : address.ToRow;
                    rAddress.ToCol = dimension._toCol < address.ToCol ? dimension._toCol : address.ToCol;
                }

                fvc = FirstValueCell(rAddress, cellValues);
                lvc = LastValueCell(rAddress, cellValues);

                csValues = cellValues;
                //rInfo.GetOffset(rInfo.Address.FromRow, rInfo.Address.FromCol);
            }
            else
            {
                //fromRow = rInfo.Address.FromRow;
                //toRow = rInfo.Address.ToRow;

                //toRow = sheet.Dimension._toRow < address.ToRow ? sheet.Dimension._toRow : address.ToRow;
                //toCol = sheet.Dimension._toCol < address.ToCol ? sheet.Dimension._toCol : address.ToCol;

                fvc = FirstValueCell(sheet, rInfo.Address);
                lvc = LastValueCell(sheet, rInfo.Address);

                CellStore<object> cellStore = (CellStore<object>)Convert.ChangeType(sheet._values, typeof(CellStore<object>));
                csValues = cellStore;
            }

            //var baseValueAddresses = new List<ExcelAddressBase>();
            // foreach (var cellInfo in rInfo)
            // {
            //     if(cellInfo.Value != null && cellInfo.IsExcelError == false)
            //     {
            //         baseValueAddresses.Add(new ExcelAddressBase(cellInfo.Address));
            //     }
            // }

            // return baseValueAddresses;

            //if (rInfo.IsInMemoryRange)
            //{
            //    //rInfo.GetValue()
            //}
            //else
            //{
            //    var baseAddress = rInfo.Address.ToExcelAddressBase();
            //    if (baseAddress.IsExternal)
            //    {
            //        rInfo.MoveNext();
            //        rInfo.Current.Value
            //        rInfo.GetOffset(rInfo.Address.FromRow, rInfo.Address.FromCol);
            //    }
            //    else
            //    {
            //        fromRow = rInfo.Address.FromRow;
            //        toRow = rInfo.Address.ToRow;

            //        toRow = sheet.Dimension._toRow < address.ToRow ? sheet.Dimension._toRow : address.ToRow;
            //        toCol = sheet.Dimension._toCol < address.ToCol ? sheet.Dimension._toCol : address.ToCol;
            //    }
            //}
            ////    var baseAddress = rInfo.Address.ToExcelAddressBase();
            ////string worksheetName = null;
            ////if(rInfo.IsInMemoryRange)
            ////{
            ////    worksheetName = baseAddress.GetAddressWorkBookWorkSheet();
            ////}

            //var sheet = rInfo.Worksheet;
            //var address = rInfo.Address;

            //var fvc = FirstValueCell(sheet, address);
            //var lvc = LastValueCell(sheet, address);

            var fromRow = fvc[0];
            var toRow = lvc[0];
            int fromCol, toCol;

            if (fvc[1] == address.FromRow)
            {
                fromCol = fvc[1];
            }
            else
            {
                int r = fromRow, c = address.FromCol;
                while (csValues.NextCellByColumn(ref r, ref c, fromRow, toRow, address.ToCol - address.FromCol))
                {
                    if (csValues.GetValue(r, c) != null)
                    {
                        break;
                    }
                    r++;
                }
                fromCol = c;
            }

            if (lvc[1] == address.ToCol)
            {
                toCol = lvc[1];
            }
            else
            {
                int r = toRow, c = address.ToCol;
                while (csValues.PrevCellByColumn(ref r, ref c, fromRow, toRow, address.ToCol - address.FromCol))
                {
                    if (csValues.GetValue(r, c) != null)
                    {
                        break;
                    }
                    r--;
                }
                toCol = c;
            }

            var offsetFromRow = Math.Min(fromRow, toRow) - rInfo.Address.FromRow;
            var offsetFromCol = Math.Min(fromCol, toCol) - rInfo.Address.FromCol;
            var offsetToRow =  Math.Max(fromRow, toRow) - rInfo.Address.FromRow;
            var offsetToCol = Math.Max(fromRow, toCol) - rInfo.Address.ToCol;

            var subRange = rInfo.GetOffset(offsetFromRow, offsetFromCol, offsetToRow, offsetToCol);
            return subRange;

            //SimpleAddress subRange = new SimpleAddress { FromRow = Math.Min(fromRow, toRow), FromCol = Math.Min(fromCol, toCol), ToRow = Math.Max(fromRow, toRow), ToCol = Math.Max(fromCol, toCol) };
            //return subRange;
        }
    }
}
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
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
                    var cellValue = values.GetValue(fromRow, fromCol);
                    if (cellValue != null)
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

        internal static SimpleAddress IterateColumns<T>(List<int> fvc, List<int> lvc, FormulaRangeAddress address, CellStore<T> csValues)
        {
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

            var offsetFromRow = Math.Min(fromRow, toRow) - address.FromRow;
            var offsetFromCol = Math.Min(fromCol, toCol) - address.FromCol;
            var offsetToRow = Math.Max(fromRow, toRow) - address.FromRow;
            var offsetToCol = Math.Max(fromRow, toCol) - address.Address.ToCol;

            return new SimpleAddress(offsetFromRow, offsetFromCol, offsetToRow, offsetToCol);
        }

        internal static IRangeInfo RangeByValue(IRangeInfo rInfo)
        {
            List<int> fvc;
            List<int> lvc;

            ExcelWorksheet sheet = null;
            FormulaRangeAddress address = null;
            ExcelAddressBase baseAddress = null;
            CellStore<object> csValues;
            SimpleAddress subrangeAddress;

            baseAddress = rInfo.Address.ToExcelAddressBase();

            if (rInfo.IsInMemoryRange)
            {
                var memRange = rInfo as InMemoryRange;
                return rInfo;
                //var cellInfocpy = memRange.GetCellInfoCopy();


                //var str = "What";

                //fvc = FirstValueCell(sheet, rInfo.Address);
                //lvc = LastValueCell(sheet, rInfo.Address);

                //CellStore<object> cellStore = (CellStore<object>)Convert.ChangeType(sheet._values, typeof(CellStore<object>));
                //csValues = cellStore;
            }
            else if (baseAddress.IsExternal)
            {
                var extRangeInfo = (EpplusExcelExternalRangeInfo)rInfo;

                var cellValues = extRangeInfo?._externalWs.CellValues._values;

                var rAddress = rInfo.Address;
                address = rInfo.Address;

                if (extRangeInfo._externalWs != null)
                {
                    var dimension = extRangeInfo._externalWs.GetDimension();
                    rAddress.ToRow = dimension._toRow < address.ToRow ? dimension._toRow : address.ToRow;
                    rAddress.ToCol = dimension._toCol < address.ToCol ? dimension._toCol : address.ToCol;
                }

                fvc = FirstValueCell(rAddress, cellValues);
                lvc = LastValueCell(rAddress, cellValues);

                subrangeAddress = IterateColumns(fvc, lvc, address, cellValues);
            }
            else
            {
                sheet = rInfo.Worksheet;
                address = rInfo.Address;
                fvc = FirstValueCell(sheet, rInfo.Address);
                lvc = LastValueCell(sheet, rInfo.Address);

                subrangeAddress = IterateColumns(fvc, lvc, address, sheet._values);
            }

            var subRange = rInfo.GetOffset(subrangeAddress.FromRow, subrangeAddress.FromCol, subrangeAddress.ToRow, subrangeAddress.ToCol);
            return subRange;
        }
    }
}
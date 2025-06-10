using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
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

        internal static SimpleAddress RangeByValue(ExcelWorksheet sheet, FormulaRangeAddress address)
        {
            var fvc = FirstValueCell(sheet, address);
            var lvc = LastValueCell(sheet, address);

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
                while (sheet._values.NextCellByColumn(ref r, ref c, fromRow, toRow, address.ToCol - address.FromCol))
                {
                    if (sheet._values.GetValue(r, c)._value != null)
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
                while (sheet._values.PrevCellByColumn(ref r, ref c, fromRow, toRow, address.ToCol - address.FromCol))
                {
                    if (sheet._values.GetValue(r, c)._value != null)
                    {
                        break;
                    }
                    r--;
                }
                toCol = c;
            }

            SimpleAddress subRange = new SimpleAddress { FromRow = Math.Min(fromRow, toRow), FromCol = Math.Min(fromCol, toCol), ToRow = Math.Max(fromRow, toRow), ToCol = Math.Max(fromCol, toCol) };
            return subRange;
        }
    }
}
/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    public struct ExcelDoubleCellValue : IComparable<ExcelDoubleCellValue>, IComparable
    {
        public ExcelDoubleCellValue(double val)
        {
            Value = val;
            CellRow = default(int?);
        }

        public ExcelDoubleCellValue(double val, int cellRow)
        {
            Value = val;
            CellRow = cellRow;
        }

        public int? CellRow;

        public double Value;

        public static implicit operator double(ExcelDoubleCellValue d)
        {
            return d.Value;
        }
        //  User-defined conversion from double to Digit
        public static implicit operator ExcelDoubleCellValue(double d)
        {
            return new ExcelDoubleCellValue(d);
        }

        public int CompareTo(ExcelDoubleCellValue other)
        {
            return Value.CompareTo(other.Value);
        }

        public int CompareTo(object obj)
        {
            if(obj is double)
            {
                return Value.CompareTo((double)obj);
            }
            return Value.CompareTo(((ExcelDoubleCellValue)obj).Value);
        }

        public override bool Equals(object obj)
        {
            return CompareTo(obj) == 0;
        }
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }        
        public static bool operator ==(ExcelDoubleCellValue a, ExcelDoubleCellValue b)
        {
            return a.Value.CompareTo(b.Value) == 0;
        }

        public static bool operator ==(ExcelDoubleCellValue a, double b)
        {
            return a.Value.CompareTo(b) == 0;
        }

        public static bool operator !=(ExcelDoubleCellValue a, ExcelDoubleCellValue b)
        {
            return a.Value.CompareTo(b.Value) != 0;
        }

        public static bool operator !=(ExcelDoubleCellValue a, double b)
        {
            return a.Value.CompareTo(b) != 0;
        }
    }
}

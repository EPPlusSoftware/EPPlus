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
using System.Collections;

namespace OfficeOpenXml.Core.CellStore
{
    /// <summary>
    /// This class represents
    /// </summary>
    internal class CellStoreValue : CellStore<ExcelValue> 
    {
        public CellStoreValue() : base()
        {

        }
        internal void SetValueRange_Value(int row, int col, object[,] array)
        {
            var rowBound = array.GetUpperBound(0);
            var colBound = array.GetUpperBound(1);
            
            for (int r = 0; r <= rowBound; r++)            
            {
                for (int c = 0; c <= colBound; c++)
                {
                    SetValue_Value(row + r, col + c, array[r, c]);
                }
            }
        }

        internal void SetValueRow_Value(int row, int col, object[] array)
        {
            for (int c = 0; c < array.Length; c++)
            {
                if(array[c] == DBNull.Value)
                {
                    SetValue_Value(row, col + c, null);
                }
                else
                {
                    SetValue_Value(row, col + c, array[c]);
                }
            }
        }
        internal void SetValueRow_Value(int row, int col, IEnumerable collection)
        {
            int offset=0;
            foreach (var v in collection)
            {
                SetValue_Value(row, col + offset, v);
                offset++;
            }
        }
        internal void SetValue_Value(int Row, int Column, object value)
        {
            var c = GetColumnIndex(Column);
            if(c != null)
            {
                int i = c.GetPointer(Row);
                if (i >= 0)
                {
                    c._values[i] = new ExcelValue { _value = value, _styleId = c._values[i]._styleId };
                    return;
                }
            }
            var v = new ExcelValue { _value = value };
            SetValue(Row, Column, v);
        }
        internal void SetValue_Style(int Row, int Column, int styleId)
        {
            var c = GetColumnIndex(Column);
            if (c != null)
            {
                int i = c.GetPointer(Row);
                if (i >= 0)
                {
                    c._values[i] = new ExcelValue { _styleId = styleId, _value = c._values[i]._value };
                    return;
                }
            }
            var v = new ExcelValue { _styleId = styleId };
            SetValue(Row, Column, v);
        }
        internal void SetValue(int Row, int Column, object value, int styleId)
        {
            var c = GetColumnIndex(Column);
            if (c != null)
            {
                int i = c.GetPointer(Row);
                if (i >= 0)
                {
                    c._values[i] = new ExcelValue { _value = value, _styleId = styleId };
                    return;
                }
            }
            var v = new ExcelValue { _value = value, _styleId = styleId};
            SetValue(Row, Column, v);
        }
    }
}
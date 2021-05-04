/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2021         EPPlus Software AB       EPPlus 5.7
 *************************************************************************************************/
using OfficeOpenXml.Core.CellStore;
using System;
using System.Collections;
using System.Collections.Generic;

namespace OfficeOpenXml.ExternalReferences
{
    public class ExcelExternalReferenceCellCollection : IEnumerable<ExcelExternalCellValue>, IEnumerator<ExcelExternalCellValue>
    {
        internal CellStore<object> _values;
        private CellStore<int> _metaData;
        CellStoreEnumerator<object> _valuesEnum;

        internal ExcelExternalReferenceCellCollection(CellStore<object> values, CellStore<int> metaData)
        {
            _values = values;
            _metaData = metaData;
        }
        public ExcelExternalCellValue this[string cellAddress]
        {
            get
            {
                if(ExcelCellBase.GetRowColFromAddress(cellAddress, out int row, out int column))
                {
                    return this[row, column];
                }
                throw (new ArgumentException("Address is not valid"));
            }
        }
        public ExcelExternalCellValue this[int row, int column]
        {
            get
            {
                if(row < 1 || column < 1 || row > ExcelPackage.MaxRows || column > ExcelPackage.MaxColumns)
                {
                    throw (new ArgumentOutOfRangeException());
                }

                return new ExcelExternalCellValue()
                {
                    Row = row,
                    Column = column,
                    Value = _values.GetValue(row, column),
                    MetaDataReference = _metaData.GetValue(row, column)
                };
            }
    }
    public ExcelExternalCellValue Current
        {
            get 
            {
                if (_valuesEnum == null) return null;
                return new ExcelExternalCellValue()
                {
                    Row = _valuesEnum.Row,
                    Column = _valuesEnum.Column,
                    Value = _valuesEnum.Value,
                    MetaDataReference = _metaData.GetValue(_valuesEnum.Row, _valuesEnum.Column)
                };
            }
        }

        object IEnumerator.Current
        {
            get
            {
                return Current;
            }
        }
        public void Dispose()
        {
            _valuesEnum.Dispose();
        }

        public IEnumerator<ExcelExternalCellValue> GetEnumerator()
        {
            return this;
        }

        public bool MoveNext()
        {
            if (_valuesEnum == null) Reset();
            return _valuesEnum.Next();
        }

        public void Reset()
        {
            _valuesEnum = new CellStoreEnumerator<object>(_values);
            _valuesEnum.Init();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this;
        }
        internal CellStoreEnumerator<object> GetCellStore(int fromRow, int fromCol, int toRow, int toCol)
        {
            return new CellStoreEnumerator<object>(_values, fromRow, fromCol, toRow, toCol);
        }
        internal object GetValue(int row, int col)
        {
            return _values.GetValue(row, col);
        }

    }
}

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
    /// <summary>
    /// A collection of <see cref="ExcelExternalCellValue" />
    /// </summary>
    public class ExcelExternalCellCollection
        : IEnumerable<ExcelExternalCellValue>, IEnumerator<ExcelExternalCellValue>
    {
        internal CellStore<object> _values;
        private CellStore<int> _metaData;
        CellStoreEnumerator<object> _valuesEnum;

        internal ExcelExternalCellCollection(CellStore<object> values, CellStore<int> metaData)
        {
            _values = values;
            _metaData = metaData;
        }
        /// <summary>
        /// An indexer to access the the external cell values 
        /// </summary>
        /// <param name="cellAddress">The cell address</param>
        /// <returns>The <see cref="ExcelExternalCellValue"/></returns>
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
        /// <summary>
        /// An indexer to access the the external cell values 
        /// </summary>
        /// <param name="row">The row of the cell to get the value from</param>
        /// <param name="column">The column of the cell to get the value from</param>
        /// <returns>The <see cref="ExcelExternalCellValue"/></returns>
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
        /// <summary>
        /// The current value of the <see cref="IEnumerable"/>
        /// </summary>
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

        /// <summary>
        /// The current value of the <see cref="IEnumerable"/>
        /// </summary>
        object IEnumerator.Current
        {
            get
            {
                return Current;
            }
        }
        /// <summary>
        /// Disposed the object
        /// </summary>
        public void Dispose()
        {
            _valuesEnum.Dispose();
        }
        /// <summary>
        /// Get the enumerator for this collection
        /// </summary>
        /// <returns></returns>
        public IEnumerator<ExcelExternalCellValue> GetEnumerator()
        {
            Reset();
            return this;
        }
        /// <summary>
        /// Move to the next item in the collection
        /// </summary>
        /// <returns>true if more items exists</returns>
        public bool MoveNext()
        {
            if (_valuesEnum == null) Reset();
            return _valuesEnum.Next();
        }
        /// <summary>
        /// Resets the enumeration
        /// </summary>
        public void Reset()
        {
            _valuesEnum = new CellStoreEnumerator<object>(_values);
            _valuesEnum.Init();
        }
        /// <summary>
        /// Get the enumerator for this collection
        /// </summary>
        /// <returns></returns>
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

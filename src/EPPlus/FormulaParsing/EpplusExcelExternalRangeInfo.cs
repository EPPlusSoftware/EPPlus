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
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Table;
using System;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.ExternalReferences;

namespace OfficeOpenXml.FormulaParsing
{
    public class EpplusExcelExternalRangeInfo : ExcelDataProvider.IRangeInfo
    {
        internal ExcelExternalWorksheet _externalWs;
        internal CellStoreEnumerator<object> _values = null;
        int _fromRow, _toRow, _fromCol, _toCol;
        int _cellCount = 0;
        ExcelAddressBase _address;
        ExcelDataProvider.ICellInfo _cell;

        public EpplusExcelExternalRangeInfo(ExcelExternalWorkbook externalWb, ExcelWorkbook wb, ExcelAddressBase address)
        {
            SetAddress(wb, address, externalWb);
        }
        private void SetAddress(ExcelWorkbook wb, ExcelAddressBase address, ExcelExternalWorkbook externalWb)
        {
            if (externalWb != null)
            {
                _externalWs = externalWb.CachedWorksheets[address.WorkSheetName];
                _fromRow = address._fromRow;
                _fromCol = address._fromCol;
                _toRow = address._toRow;
                _toCol = address._toCol;
                _address = address;
                if (_externalWs != null)
                {
                    _values = _externalWs.CellValues.GetCellStore(_fromRow, _fromCol, _toRow, _toCol);
                    _cell = new ExternalCellInfo(_externalWs, _values);
                }
            }
        }

        public int GetNCells()
        {
            return ((_toRow - _fromRow) + 1) * ((_toCol - _fromCol) + 1);
        }

        public bool IsRef
        {
            get
            {
                return _externalWs == null || _fromRow < 0 || _toRow < 0;
            }
        }
        public bool IsEmpty
        {
            get
            {
                if (_cellCount > 0)
                {
                    return false;
                }
                else if (_values == null) return true;
                else if (_values.Next())
                {
                    _values.Reset();
                    return false;
                }
                else
                {
                    return true;
                }
            }
        }
        public bool IsMulti
        {
            get
            {
                if (_cellCount == 0)
                {
                    if (_values == null) return false;
                    if (_values.Next() && _values.Next())
                    {
                        _values.Reset();
                        return true;
                    }
                    else
                    {
                        _values.Reset();
                        return false;
                    }
                }
                else if (_cellCount > 1)
                {
                    return true;
                }
                return false;
            }
        }

        public ExcelDataProvider.ICellInfo Current
        {
            get { return _cell; }
        }

        public ExcelWorksheet Worksheet
        {
            get { return null; }
        }

        public void Dispose()
        {
        }

        object System.Collections.IEnumerator.Current
        {
            get
            {
                return this;
            }
        }

        public bool MoveNext()
        {
            if (_values == null) return false;
            _cellCount++;
            return _values.MoveNext();
        }

        public void Reset()
        {
            _cellCount = 0;
            _values?.Init();
        }


        public bool NextCell()
        {
            if (_values == null) return false;
            _cellCount++;
            return _values.MoveNext();
        }

        public IEnumerator<ExcelDataProvider.ICellInfo> GetEnumerator()
        {
            Reset();
            return this;
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this;
        }

        public ExcelAddressBase Address
        {
            get { return _address; }
        }

        public object GetValue(int row, int col)
        {
            return _externalWs?.CellValues.GetValue(row, col);
        }

        public object GetOffset(int rowOffset, int colOffset)
        {
            if (_values == null) return null;
            if (_values.Row < _fromRow || _values.Column < _fromCol)
            {
                return _externalWs?.CellValues.GetValue(_fromRow + rowOffset, _fromCol + colOffset);
            }
            else
            {
                return _externalWs?.CellValues.GetValue(_values.Row + rowOffset, _values.Column + colOffset);
            }
        }
    }

    public class ExternalCellInfo : ExcelDataProvider.ICellInfo
    {
        ExcelExternalWorksheet _ws;
        CellStoreEnumerator<object> _values;
        internal ExternalCellInfo(ExcelExternalWorksheet ws, CellStoreEnumerator<object> values)
        {
            _ws = ws;
            _values = values;
        }
        public string Address
        {
            get { return _values.CellAddress; }
        }

        public int Row
        {
            get { return _values.Row; }
        }

        public int Column
        {
            get { return _values.Column; }
        }

        public string Formula
        {
            get
            {
                return "";
            }
        }

        public object Value
        {
            get
            {
                return _values.Value;
            }
        }

        public double ValueDouble
        {
            get { return ConvertUtil.GetValueDouble(_values.Value, true); }
        }
        public double ValueDoubleLogical
        {
            get { return ConvertUtil.GetValueDouble(_values.Value, false); }
        }
        public bool IsHiddenRow
        {
            get
            {
                return false;
            }
        }

        public bool IsExcelError
        {
            get { return ExcelErrorValue.Values.IsErrorValue(_values.Value); }
        }

        public IList<Token> Tokens
        {
            get
            {
                return new List<Token>();
            }
        }

        public ulong Id
        {
            get
            {
                return ExcelCellBase.GetCellId(_ws.SheetId, _values.Row, _values.Column);
            }
        }

        public string WorksheetName
        {
            get { return _ws.Name; }
        }
    }
}
    
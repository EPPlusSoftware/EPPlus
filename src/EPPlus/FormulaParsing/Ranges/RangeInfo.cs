/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/31/2022         EPPlus Software AB           EPPlus 6.1
 *************************************************************************************************/
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Ranges
{
    /// <summary>
    /// EPPlus implementation of the <see cref="IRangeInfo"/> interface
    /// </summary>
    public class RangeInfo : IRangeInfo
    {
        internal ExcelWorksheet _ws;
        CellStoreEnumerator<ExcelValue> _values = null;
        private RangeDefinition _size;
        ParsingContext _context;
        //int _fromRow, _toRow, _fromCol, _toCol;
        int _cellCount = 0;
        FormulaRangeAddress _address;
        ICellInfo _cell;

        /// <summary>
        /// Constructor
        /// </summary>
        public RangeInfo(FormulaRangeAddress address, ParsingContext ctx)
        {
            _context = ctx;
            _address = address;
            var wsIx = address.WorksheetIx >= 0 ? address.WorksheetIx : ctx.CurrentCell.WorksheetIx;
            if (wsIx >= 0 && wsIx < ctx.Package.Workbook.Worksheets.Count)
            {
                _ws = ctx.Package.Workbook.Worksheets[wsIx];
                _values = new CellStoreEnumerator<ExcelValue>(_ws._values, address.FromRow, address.FromCol, address.ToRow, address.ToCol);
                _cell = new CellInfo(_ws, _values);
            }
            _size = new RangeDefinition(address.ToRow - address.FromRow + 1, (short)(address.ToCol - address.FromCol + 1));
        }
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="ws">The worksheet</param>
        /// <param name="fromRow"></param>
        /// <param name="fromCol"></param>
        /// <param name="toRow"></param>
        /// <param name="toCol"></param>
        /// <param name="ctx">Parsing context</param>
        public RangeInfo(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol, ParsingContext ctx)
        {
            _context = ctx;
            var address = new ExcelAddressBase(fromRow, fromCol, toRow, toCol);
            address._ws = ws.Name;
            SetAddress(ws, address);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ws">The worksheet</param>
        public RangeInfo(ExcelWorksheet ws, ParsingContext ctx)
        {
            _context = ctx;
            _address = new FormulaRangeAddress(ctx) { WorksheetIx = (short)ws.PositionId };
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="address"></param>
        public RangeInfo(ExcelWorksheet ws, ExcelAddressBase address)
        {
            SetAddress(ws, address);
        }

        private void SetAddress(ExcelWorksheet ws, ExcelAddressBase address)
        {
            _ws = ws;
            _address = new FormulaRangeAddress(null) { FromRow = address._fromRow, FromCol = address._fromCol, ToRow = address._toRow, ToCol = address._toCol, WorksheetIx = (short)ws.PositionId };
            if (_ws != null && _ws.IsDisposed == false)
            {
                _values = new CellStoreEnumerator<ExcelValue>(_ws._values, address._fromRow, address._fromCol, address._toRow, address._toCol);
                _cell = new CellInfo(_ws, _values);
            }
            _size = new RangeDefinition(address._toRow - address._fromRow + 1, (short)(address._toCol - address._fromCol + 1));
        }

        /// <summary>
        /// The total number of cells (including empty) of the range
        /// </summary>
        /// <returns></returns>
        public int GetNCells()
        {
            return ((_address.ToRow - _address.FromRow) + 1) * ((_address.ToCol - _address.FromCol) + 1);
        }

        /// <summary>
        /// Returns true if the range represents a reference
        /// </summary>
        public bool IsRef
        {
            get
            {
                return _ws == null || _address.FromRow < 0 || _address.ToRow < 0;
            }
        }
        /// <summary>
        /// Returns true if the range is empty
        /// </summary>
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

        /// <summary>
        /// Returns true if more than one cell
        /// </summary>
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

        /// <summary>
        /// Size of the range
        /// </summary>
        public RangeDefinition Size => _size;

        /// <summary>
        /// Returns true if the range is an <see cref="InMemoryRange"/>
        /// </summary>
        public bool IsInMemoryRange => false;

        /// <summary>
        /// Current cell
        /// </summary>
        public ICellInfo Current
        {
            get { return _cell; }
        }

        /// <summary>
        /// The worksheet
        /// </summary>
        public ExcelWorksheet Worksheet
        {
            get { return _ws; }
        }

        /// <summary>
        /// Runs at dispose of this instance
        /// </summary>
        public void Dispose()
        {
            //_values = null;
            //_ws = null;
            //_cell = null;
        }

        /// <summary>
        /// IEnumerator.Current
        /// </summary>
        object System.Collections.IEnumerator.Current
        {
            get
            {
                return this;
            }
        }

        /// <summary>
        /// Moves to next cell
        /// </summary>
        /// <returns></returns>
        public bool MoveNext()
        {
            if (_values == null) return false;
            _cellCount++;
            return _values.MoveNext();
        }

        /// <summary>
        /// 
        /// </summary>
        public void Reset()
        {
            _cellCount = 0;
            _values?.Init();
        }


        /// <summary>
        /// Moves to next cell
        /// </summary>
        /// <returns></returns>
        public bool NextCell()
        {
            if (_values == null) return false;
            _cellCount++;
            return _values.MoveNext();
        }

        /// <summary>
        /// Returns enumerator for cells
        /// </summary>
        /// <returns></returns>
        public IEnumerator<ICellInfo> GetEnumerator()
        {
            Reset();
            return this;
        }

        /// <summary>
        /// Returns enumerator for cells
        /// </summary>
        /// <returns></returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this;
        }

        //public ExcelAddressBase Address
        //{
        //    get { return _address; }
        //}


        /// <summary>
        /// Address of the range
        /// </summary>
        public FormulaRangeAddress Address { get { return _address; } }

        /// <summary>
        /// Returns the cell value by 0-based index
        /// </summary>
        /// <param name="row">0-based row index</param>
        /// <param name="col">0-based col index</param>
        /// <returns>Cell value</returns>
        public object GetValue(int row, int col)
        {
            return _ws?.GetValue(row, col);
        }

        /// <summary>
        /// Return value by offset
        /// </summary>
        /// <param name="rowOffset"></param>
        /// <param name="colOffset"></param>
        /// <returns></returns>
        public object GetOffset(int rowOffset, int colOffset)
        {
            if (_values == null) return null;
            if (_values.Row < _address.FromRow || _values.Column < _address.FromCol)
            {
                return _ws.GetValue(_address.FromRow + rowOffset, _address.FromCol + colOffset);
            }
            else
            {
                return _ws.GetValue(_values.Row + rowOffset, _values.Column + colOffset);
            }
        }
    }
}

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
        //ParsingContext _context;
        //int _fromRow, _toRow, _fromCol, _toCol;
        int _cellCount = 0;
        FormulaRangeAddress _address;
        ICellInfo _cell;

        /// <summary>
        /// Constructor
        /// </summary>
        public RangeInfo(FormulaRangeAddress address)
        {
            _address = address;
            if(address.WorksheetIx==-1)
            {
                return;
            }
            var ctx = address._context;
            var wsIx = address.WorksheetIx >= 0 ? address.WorksheetIx : ctx.CurrentCell.WorksheetIx;
            if (wsIx >= 0 && wsIx < ctx.Package.Workbook.Worksheets.Count)
            {
                _ws = ctx.Package.Workbook.GetWorksheetByIndexInList(wsIx);
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
            /// <param name="extRef">External reference id</param>
        public RangeInfo(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol, ParsingContext ctx, int extRef = -1)
        {
            var address = new ExcelAddressBase(fromRow, fromCol, toRow, toCol);
            address._ws = ws.Name;
            SetAddress(ws, address, ctx);
            _address.ExternalReferenceIx = extRef;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ws">The worksheet</param>
        /// <param name="ctx">Parsing context</param>
        public RangeInfo(ExcelWorksheet ws, ParsingContext ctx)
        {
            _address = new FormulaRangeAddress(ctx) { WorksheetIx = (short)ws.PositionId };
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="address"></param>
        /// <param name="ctx">Parsing context</param>
        public RangeInfo(ExcelWorksheet ws, ExcelAddressBase address, ParsingContext ctx=null)
        {
            SetAddress(ws, address, ctx);
        }

        private void SetAddress(ExcelWorksheet ws, ExcelAddressBase address, ParsingContext ctx)
        {
            _ws = ws;
            _address = new FormulaRangeAddress(ctx, address) 
            { 
                WorksheetIx = (short)ws.PositionId,
            };
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
        /// Returns true if the range represents an invalid reference
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
        FormulaRangeAddress _dimension = null;
        /// <summary>
        /// Dimension
        /// </summary>
        public FormulaRangeAddress Dimension
        {
            get
            {
                if(_dimension == null)
                {
                    var d = _ws.Dimension;
                    _dimension =new FormulaRangeAddress() 
                    { 
                        FromRow=d._fromRow,
                        FromCol=d._fromCol,
                        ToRow=d._toRow,
                        ToCol=d._toCol                        
                    };
                }
                return _dimension; 
            }
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
            return _ws.GetValue(_address.FromRow + rowOffset, _address.FromCol + colOffset);
        }

        /// <summary>
        /// Returns a subrange
        /// </summary>
        /// <param name="rowOffsetStart"></param>
        /// <param name="colOffsetStart"></param>
        /// <param name="rowOffsetEnd"></param>
        /// <param name="colOffsetEnd"></param>
        /// <returns></returns>
        public IRangeInfo GetOffset(int rowOffsetStart, int colOffsetStart, int rowOffsetEnd, int colOffsetEnd)
        {
            if (_values == null) return null;

            var sr = _address.FromRow;
            var sc = _address.FromCol;
            return new RangeInfo(_ws, sr + rowOffsetStart, sc + colOffsetStart, sr + rowOffsetEnd, sc + colOffsetEnd, _address._context, _address.ExternalReferenceIx);
        }
        /// <summary>
        /// Is hidden
        /// </summary>
        /// <param name="rowOffset"></param>
        /// <param name="colOffset"></param>
        /// <returns></returns>
        public bool IsHidden(int rowOffset, int colOffset)
        {
            var row = _ws.GetValueInner(_address.FromRow + rowOffset, 0) as RowInternal;
            if (row != null)
            {
                return row.Hidden || row.Height == 0;
            }
            else
            {
                return false;
            }

        }
    }
}

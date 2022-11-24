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

namespace OfficeOpenXml.FormulaParsing.Ranges
{
    /// <summary>
    /// Provide the formula parser with information about an workbook external range.
    /// </summary>
    public class EpplusExcelExternalRangeInfo : IRangeInfo
    {
        internal ExcelExternalWorksheet _externalWs;
        internal CellStoreEnumerator<object> _values = null;
        private RangeDefinition _size;
        int _fromRow, _toRow, _fromCol, _toCol;
        int _cellCount = 0;
        FormulaRangeAddress _address;
        ICellInfo _cell;

        /// <summary>
        /// The constructor
        /// </summary>
        /// <param name="externalReferenceIx">Index of the external workbook</param>
        /// <param name="worksheetIx">The external worksheet index/id</param>
        /// <param name="fromRow">The from row of the address</param>
        /// <param name="fromCol">The from column of the address</param>
        /// <param name="toRow">The to row of the address</param>
        /// <param name="toCol">The to column of the address</param>
        /// <param name="ctx">Parsing context</param>
        public EpplusExcelExternalRangeInfo(int externalReferenceIx, int worksheetIx, int fromRow, int fromCol, int toRow, int toCol, ParsingContext ctx)
        {
            _fromRow = fromRow;
            _fromCol = fromCol;
            _toRow = toRow;
            _toCol = toCol;
            _address = new FormulaRangeAddress(ctx)
            {
                FromRow = fromRow,
                FromCol = fromCol,
                ToRow = toRow,
                ToCol = toCol
            };
            _size = new RangeDefinition(toRow + fromCol + 1, (short)(toCol - fromCol + 1));
            if (externalReferenceIx > 0 && ctx.Package != null && ctx.Package.Workbook.ExternalLinks.Count < externalReferenceIx)
            {
                var externalWb = ctx.Package.Workbook.ExternalLinks[externalReferenceIx].As.ExternalWorkbook;
                if (externalWb != null)
                {

                }
                _externalWs = externalWb.CachedWorksheets[worksheetIx];
                if (_externalWs != null)
                {
                    _address.ExternalReferenceIx = (short)externalReferenceIx;
                    _address.WorksheetIx = (short)_externalWs.SheetId;
                    _values = _externalWs.CellValues.GetCellStore(_fromRow, _fromCol, _toRow, _toCol);
                    _cell = new ExternalCellInfo(_externalWs, _values);
                }
            }
            else if (ctx.Package != null && ctx.Package.Workbook.Worksheets[worksheetIx] != null)
            {
                _address.WorksheetIx = (short)worksheetIx;
            }
        }
        /// <summary>
        /// Get the number of cells in the range
        /// </summary>
        /// <returns></returns>
        public int GetNCells()
        {
            return (_toRow - _fromRow + 1) * (_toCol - _fromCol + 1);
        }
        /// <summary>
        /// If the range is invalid (#REF!)
        /// </summary>
        public bool IsRef
        {
            get
            {
                return _externalWs == null || _fromRow < 0 || _toRow < 0;
            }
        }
        /// <summary>
        /// If the range is empty, ie contains no set cells.
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
        /// If the range contains more than one set cell.
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
        /// True if this is a range that doesn't is connected to a worksheet.
        /// </summary>
        public bool IsInMemoryRange => false;

        /// <summary>
        /// Return the current object in the enumeration
        /// </summary>
        public ICellInfo Current
        {
            get { return _cell; }
        }
        /// <summary>
        /// Not applicable for external ranges.. Returns null
        /// </summary>
        public ExcelWorksheet Worksheet
        {
            get { return null; }
        }
        /// <summary>
        /// Called when the object is disposed.
        /// </summary>
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

        /// <summary>
        /// Moves to the next item in the enumeration
        /// </summary>
        /// <returns>returns true until the enumeration has reached the last cell.</returns>
        public bool MoveNext()
        {
            if (_values == null) return false;
            _cellCount++;
            return _values.MoveNext();
        }
        /// <summary>
        /// Resets the enumeration
        /// </summary>
        public void Reset()
        {
            _cellCount = 0;
            _values?.Init();
        }

        /// <summary>
        /// Moves to the next item in the enumeration
        /// </summary>
        /// <returns></returns>
        public bool NextCell()
        {
            if (_values == null) return false;
            _cellCount++;
            return _values.MoveNext();
        }

        /// <summary>
        /// Gets the enumerator
        /// </summary>
        /// <returns>The enumerator</returns>
        public IEnumerator<ICellInfo> GetEnumerator()
        {
            Reset();
            return this;
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this;
        }

        /// <summary>
        /// The address of the range
        /// </summary>
        //public ExcelAddressBase Address
        //{
        //    get { return _address; }
        //}

        public FormulaRangeAddress Address { get { return _address; } }

        /// <summary>
        /// Gets the value 
        /// </summary>
        /// <param name="row">The row</param>
        /// <param name="col">The column</param>
        /// <returns></returns>
        public object GetValue(int row, int col)
        {
            return _externalWs?.CellValues.GetValue(row, col);
        }
        /// <summary>
        /// Get the value from the range with the offset from the top-left cell
        /// </summary>
        /// <param name="rowOffset">The row offset.</param>
        /// <param name="colOffset">The column offset.</param>
        /// <returns></returns>
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

    /// <summary>
    /// Provides information about an external cell in an external range.
    /// </summary>
    public class ExternalCellInfo : ICellInfo
    {
        ExcelExternalWorksheet _ws;
        CellStoreEnumerator<object> _values;
        internal ExternalCellInfo(ExcelExternalWorksheet ws, CellStoreEnumerator<object> values)
        {
            _ws = ws;
            _values = values;
        }
        /// <summary>
        /// The cell address.
        /// </summary>
        public string Address
        {
            get { return _values.CellAddress; }
        }
        /// <summary>
        /// The row of the cell
        /// </summary>
        public int Row
        {
            get { return _values.Row; }
        }

        /// <summary>
        /// The column of the cell
        /// </summary>
        public int Column
        {
            get { return _values.Column; }
        }
        /// <summary>
        /// Formula. Always return Empty.String for external cells.
        /// </summary>
        public string Formula
        {
            get
            {
                return "";
            }
        }
        /// <summary>
        /// The value of the current cell.
        /// </summary>
        public object Value
        {
            get
            {
                return _values.Value;
            }
        }
        /// <summary>
        /// The value as double of the current cell. Bools will be ignored.
        /// </summary>
        public double ValueDouble
        {
            get { return ConvertUtil.GetValueDouble(_values.Value, true); }
        }
        /// <summary>
        /// The value as double of the current cell.
        /// </summary>
        public double ValueDoubleLogical
        {
            get { return ConvertUtil.GetValueDouble(_values.Value, false); }
        }
        /// <summary>
        /// If the row of the cell is hidden
        /// </summary>
        public bool IsHiddenRow
        {
            get
            {
                return false;
            }
        }

        /// <summary>
        /// If the value of the cell is an Excel Error
        /// </summary>
        public bool IsExcelError
        {
            get { return ExcelErrorValue.Values.IsErrorValue(_values.Value); }
        }
        /// <summary>
        /// Tokens for the formula. Not applicable to External cells.
        /// </summary>
        public IList<Token> Tokens
        {
            get
            {
                return new List<Token>();
            }
        }
        /// <summary>
        /// The cell id
        /// </summary>
        public ulong Id
        {
            get
            {
                return ExcelCellBase.GetCellId(_ws.SheetId, _values.Row, _values.Column);
            }
        }
        /// <summary>
        /// The name of the worksheet.
        /// </summary>
        public string WorksheetName
        {
            get { return _ws.Name; }
        }
    }
}

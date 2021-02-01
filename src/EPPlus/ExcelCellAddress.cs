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
using System.Text;

namespace OfficeOpenXml
{
    /// <summary>
    /// A single cell address 
    /// </summary>
    public class ExcelCellAddress
    {
        /// <summary>
        /// Initializes a new instance of the ExcelCellAddress class.
        /// </summary>
        public ExcelCellAddress()
            : this(1, 1)
        {

        }

        private int _row;
        private bool _isRowFixed;
        private int _column;
        private bool _isColumnFixed;
        private string _address;
        /// <summary>
        /// Initializes a new instance of the ExcelCellAddress class.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="column">The column.</param>
        /// <param name="isRowFixed">If the row is fixed, prefixed with $</param>
        /// <param name="isColumnFixed">If the column is fixed, prefixed with $</param>
        public ExcelCellAddress(int row, int column, bool isRowFixed = false, bool isColumnFixed = false)
        {
            Row = row;
            Column = column;
            _isRowFixed = isRowFixed;
            _isColumnFixed = isColumnFixed;


        }
        /// <summary>
        /// Initializes a new instance of the ExcelCellAddress class.
        /// </summary>
        ///<param name="address">The address</param>
        public ExcelCellAddress(string address)
        {
            Address = address; 
        }
        /// <summary>
        /// Row
        /// </summary>
        public int Row
        {
            get
            {
                return this._row;
            }
            private set
            {
                if (value <= 0)
                {
                    throw new ArgumentOutOfRangeException("value", "Row cannot be less than 1.");
                }
                this._row = value;
                if(_column>0) 
                    _address = ExcelCellBase.GetAddress(_row, _column);
                else
                    _address = "#REF!";
            }
        }
        /// <summary>
        /// Column
        /// </summary>
        public int Column
        {
            get
            {
                return this._column;
            }
            private set
            {
                if (value <= 0)
                {
                    throw new ArgumentOutOfRangeException("value", "Column cannot be less than 1.");
                }
                this._column = value;
                if (_row > 0)
                    _address = ExcelCellBase.GetAddress(_row, _column);
                else
                    _address = "#REF!";
            }
        }
        /// <summary>
        /// Celladdress
        /// </summary>
        public string Address
        {
            get
            {
                return _address;
            }
            internal set
            {
                _address = value;
                ExcelCellBase.GetRowColFromAddress(_address, out _row, out _column,out _isRowFixed, out _isColumnFixed);
            }
        }
        public bool IsRowFixed 
        { 
            get
            {
                return _isRowFixed;
            }
        }
        public bool IsColumnFixed
        {
            get
            {
                return _isColumnFixed;
            }
        }

    /// <summary>
    /// If the address is an invalid reference (#REF!)
    /// </summary>
    public bool IsRef
        {
            get
            {
                return _row <= 0;
            }
        }

        /// <summary>
        /// Returns the letter corresponding to the supplied 1-based column index.
        /// </summary>
        /// <param name="column">Index of the column (1-based)</param>
        /// <returns>The corresponding letter, like A for 1.</returns>
        public static string GetColumnLetter(int column)
        {
            if (column > ExcelPackage.MaxColumns || column < 1)
            {
                throw new InvalidOperationException("Invalid 1-based column index: " + column + ". Valid range is 1 to " + ExcelPackage.MaxColumns);
            }
            return ExcelCellBase.GetColumnLetter(column);
        }
    }
}


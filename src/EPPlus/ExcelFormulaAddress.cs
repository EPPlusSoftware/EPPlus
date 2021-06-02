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

namespace OfficeOpenXml
{
    /// <summary>
    /// Range address used in the formula parser
    /// </summary>
    public class ExcelFormulaAddress : ExcelAddressBase
    {
        /// <summary>
        /// Creates a Address object
        /// </summary>
        internal ExcelFormulaAddress()
            : base()
        {
        }

        /// <summary>
        /// Creates an Address object
        /// </summary>
        /// <param name="fromRow">start row</param>
        /// <param name="fromCol">start column</param>
        /// <param name="toRow">End row</param>
        /// <param name="toColumn">End column</param>
        public ExcelFormulaAddress(int fromRow, int fromCol, int toRow, int toColumn)
            : base(fromRow, fromCol, toRow, toColumn)
        {
            _ws = "";
        }
        /// <summary>
        /// Creates an Address object
        /// </summary>
        /// <param name="address">The formula address</param>
        /// <param name="worksheet">The worksheet</param>
        public ExcelFormulaAddress(string address, ExcelWorksheet worksheet)
            : base(address, worksheet?.Workbook, worksheet?.Name)
        {
            SetFixed();
        }
        
        internal ExcelFormulaAddress(string ws, string address)
            : base(address)
        {
            if (string.IsNullOrEmpty(_ws)) _ws = ws;
            SetFixed();
        }
        internal ExcelFormulaAddress(string ws, string address, bool isName)
            : base(address, isName)
        {
            if (string.IsNullOrEmpty(_ws)) _ws = ws;
            if(!isName)
                SetFixed();
        }

        private void SetFixed()
        {
            if (Address.IndexOf('[') >= 0) return;
            var address=FirstAddress;
            if(_fromRow==_toRow && _fromCol==_toCol)
            {
                GetFixed(address, out _fromRowFixed, out _fromColFixed);
            }
            else
            {
                var cells = address.Split(':');                
                if(cells.Length>1) //If 1 then the address is the entire worksheet
                {
                    GetFixed(cells[0], out _fromRowFixed, out _fromColFixed);
                    GetFixed(cells[1], out _toRowFixed, out _toColFixed);
                }
            }
        }

        private void GetFixed(string address, out bool rowFixed, out bool colFixed)
        {            
            rowFixed=colFixed=false;
            var ix=address.IndexOf('$');
            while(ix>-1)
            {
                ix++;
                if(ix < address.Length)
                {
                    if(address[ix]>='0' && address[ix]<='9')
                    {
                        rowFixed=true;
                        break;
                    }
                    else
                    {
                        colFixed=true;
                    }
                }
                ix = address.IndexOf('$', ix);
            }
        }
        /// <summary>
        /// The address for the range
        /// </summary>
        /// <remarks>Examples of addresses are "A1" "B1:C2" "A:A" "1:1" "A1:E2,G3:G5" </remarks>
        public new string Address
        {
            get
            {
                if (string.IsNullOrEmpty(_address) && _fromRow>0)
                {
                    _address = GetAddress(_fromRow, _fromCol, _toRow, _toCol, _fromRowFixed, _toRowFixed, _fromColFixed, _toColFixed);
                }
                return _address;
            }
            set
            {                
                SetAddress(value, null, null);
                ChangeAddress();
                SetFixed();
            }
        }
        internal new List<ExcelFormulaAddress> _addresses;
        /// <summary>
        /// Addresses can be separated by a comma. If the address contains multiple addresses this list contains them.
        /// </summary>
        public new List<ExcelFormulaAddress> Addresses
        {
            get
            {
                if (_addresses == null)
                {
                    _addresses = new List<ExcelFormulaAddress>();
                }
                return _addresses;

            }
        }
        internal string GetOffset(int row, int column, bool withWbWs=false)
        {
            int fromRow = _fromRow, fromCol = _fromCol, toRow = _toRow, tocol = _toCol;
            var isMulti = (fromRow != toRow || fromCol != tocol);
            if (!_fromRowFixed)
            {
                fromRow += row;
            }
            if (!_fromColFixed)
            {
                fromCol += column;
            }
            if (isMulti)
            {
                if (!_toRowFixed)
                {
                    toRow += row;
                }
                if (!_toColFixed)
                {
                    tocol += column;
                }
            }
            else
            {
                toRow = fromRow;
                tocol = fromCol;
            }
            string a = GetAddress(fromRow, fromCol, toRow, tocol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
            if (Addresses != null)
            {
                foreach (var sa in Addresses)
                {
                    a+="," + sa.GetOffset(row, column, withWbWs);
                }
            }
            if(withWbWs)
            {
                return GetAddressWorkBookWorkSheet() + a;
            }
            else
            {
                return a;
            }
        }
    }
}

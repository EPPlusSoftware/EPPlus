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
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml
{
    /// <summary>
    /// A range address
    /// </summary>
    /// <remarks>Examples of addresses are "A1" "B1:C2" "A:A" "1:1" "A1:E2,G3:G5" </remarks>
    public class ExcelAddressBase : ExcelCellBase
    {
        internal int _fromRow=-1, _toRow, _fromCol, _toCol;
        internal bool _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed;
        internal string _wb;
        internal string _ws;
        internal string _address;

        internal enum eAddressCollition
        {
            No,
            Partly,
            Inside,
            Equal
        }
        #region "Constructors"
        internal ExcelAddressBase()
        {
        }
        /// <summary>
        /// Creates an Address object
        /// </summary>
        /// <param name="fromRow">start row</param>
        /// <param name="fromCol">start column</param>
        /// <param name="toRow">End row</param>
        /// <param name="toColumn">End column</param>
        public ExcelAddressBase(int fromRow, int fromCol, int toRow, int toColumn)
        {
            _fromRow = fromRow;
            _toRow = toRow;
            _fromCol = fromCol;
            _toCol = toColumn;
            Validate();

            _address = GetAddress(_fromRow, _fromCol, _toRow, _toCol);
        }
        /// <summary>
        /// Creates an Address object
        /// </summary>
        /// <param name="worksheetName">Worksheet name</param>
        /// <param name="fromRow">Start row</param>
        /// <param name="fromCol">Start column</param>
        /// <param name="toRow">End row</param>
        /// <param name="toColumn">End column</param>
        public ExcelAddressBase(string worksheetName, int fromRow, int fromCol, int toRow, int toColumn)
        {
            _ws = worksheetName;
            _fromRow = fromRow;
            _toRow = toRow;
            _fromCol = fromCol;
            _toCol = toColumn;
            Validate();

            _address = GetAddress(_fromRow, _fromCol, _toRow, _toCol);
        }

        internal static bool IsTableAddress(string address)
        {
            SplitAddress(address, out string wb, out string ws, out string intAddress);
            var lPos = intAddress.IndexOf('[');
            if(lPos >= 0) 
            {
                var rPos= intAddress.IndexOf(']',lPos);
                if(rPos>lPos)
                {
                    var c=intAddress[lPos+1];
                    return !((c >= '0' && c <= '9') || c == '-');
                }
            }
            return false;
        }

        /// <summary>
        /// Creates an Address object
        /// </summary>
        /// <param name="fromRow">Start row</param>
        /// <param name="fromCol">Start column</param>
        /// <param name="toRow">End row</param>
        /// <param name="toColumn">End column</param>
        /// <param name="fromRowFixed">Start row fixed</param>
        /// <param name="fromColFixed">Start column fixed</param>
        /// <param name="toRowFixed">End row fixed</param>
        /// <param name="toColFixed">End column fixed</param>
        public ExcelAddressBase(int fromRow, int fromCol, int toRow, int toColumn, bool fromRowFixed, bool fromColFixed, bool toRowFixed, bool toColFixed) :
            this(fromRow, fromCol, toRow, toColumn, fromRowFixed, fromColFixed, toRowFixed, toColFixed, null, null)
        {

        }
        internal ExcelAddressBase(int fromRow, int fromCol, int toRow, int toColumn, bool fromRowFixed, bool fromColFixed, bool toRowFixed, bool toColFixed, string worksheetName, string prevAddress)
        {
            _fromRow = fromRow;
            _toRow = toRow;
            _fromCol = fromCol;
            _toCol = toColumn;
            _fromRowFixed = fromRowFixed;
            _fromColFixed = fromColFixed;
            _toRowFixed = toRowFixed;
            _toColFixed = toColFixed;
            _ws = worksheetName;
            Validate();
            var prevAddressHasWs = prevAddress != null && prevAddress.IndexOf("!") > 0 && !prevAddress.EndsWith("!");
            _address = GetAddress(_fromRow, _fromCol, _toRow, _toCol, _fromRowFixed, fromColFixed, _toRowFixed, _toColFixed );
            if(prevAddressHasWs && !string.IsNullOrEmpty(_ws))
            {
                if(ExcelWorksheet.NameNeedsApostrophes(_ws))
                {
                    _address = $"'{_ws.Replace("'","''")}'!{_address}";
                }
                else
                {
                    _address = $"{_ws}!{_address}";
                }
            }
        }
        /// <summary>
        /// Creates an Address object
        /// </summary>
        /// <remarks>Examples of addresses are "A1" "B1:C2" "A:A" "1:1" "A1:E2,G3:G5" </remarks>
        /// <param name="address">The Excel Address</param>
        /// <param name="wb">The workbook to verify any defined names from</param>
        /// <param name="wsName">The name of the worksheet the address referes to</param>
        /// <ws></ws>
        public ExcelAddressBase(string address, ExcelWorkbook wb=null, string wsName=null)
        {
            SetAddress(address, wb, wsName);
            if (string.IsNullOrEmpty(_ws) && string.IsNullOrEmpty(_wb)) _ws = wsName;
        }
        /// <summary>
        /// Creates an Address object
        /// </summary>
        /// <remarks>Examples of addresses are "A1" "B1:C2" "A:A" "1:1" "A1:E2,G3:G5" </remarks>
        /// <param name="address">The Excel Address</param>
        /// <param name="pck">Reference to the package to find information about tables and names</param>
        /// <param name="referenceAddress">The address</param>
        public ExcelAddressBase(string address, ExcelPackage pck, ExcelAddressBase referenceAddress)
        {
            SetAddress(address, null, null);
            SetRCFromTable(pck, referenceAddress);
        }

        internal void SetRCFromTable(ExcelPackage pck, ExcelAddressBase referenceAddress)
        {
            if (string.IsNullOrEmpty(_wb) && Table != null)
            {
                foreach (var ws in pck.Workbook.Worksheets)
                {
                    if (ws is ExcelChartsheet) continue;
                    foreach (var t in ws.Tables)
                    {
                        if (t.Name.Equals(Table.Name, StringComparison.OrdinalIgnoreCase))
                        {
                            _ws = ws.Name;
                            if (Table.IsAll)
                            {
                                _fromRow = t.Address._fromRow;
                                _toRow = t.Address._toRow;
                            }
                            else
                            {
                                if (Table.IsThisRow)
                                {
                                    if (referenceAddress == null)
                                    {
                                        _fromRow = -1;
                                        _toRow = -1;
                                    }
                                    else
                                    {
                                        _fromRow = referenceAddress._fromRow;
                                        _toRow = _fromRow;
                                    }
                                }
                                else if (Table.IsHeader && Table.IsData)
                                {
                                    _fromRow = t.Address._fromRow;
                                    _toRow = t.ShowTotal ? t.Address._toRow - 1 : t.Address._toRow;
                                }
                                else if (Table.IsData && Table.IsTotals)
                                {
                                    _fromRow = t.ShowHeader ? t.Address._fromRow + 1 : t.Address._fromRow;
                                    _toRow = t.Address._toRow;
                                }
                                else if (Table.IsHeader)
                                {
                                    _fromRow = t.ShowHeader ? t.Address._fromRow : -1;
                                    _toRow = t.ShowHeader ? t.Address._fromRow : -1;
                                }
                                else if (Table.IsTotals)
                                {
                                    _fromRow = t.ShowTotal ? t.Address._toRow : -1;
                                    _toRow = t.ShowTotal ? t.Address._toRow : -1;
                                }
                                else
                                {
                                    _fromRow = t.ShowHeader ? t.Address._fromRow + 1 : t.Address._fromRow;
                                    _toRow = t.ShowTotal ? t.Address._toRow - 1 : t.Address._toRow;
                                }
                            }

                            if (string.IsNullOrEmpty(Table.ColumnSpan))
                            {
                                _fromCol = t.Address._fromCol;
                                _toCol = t.Address._toCol;
                                return;
                            }
                            else
                            {
                                var col = t.Address._fromCol;
                                var cols = Table.ColumnSpan.Split(':');
                                foreach (var c in t.Columns)
                                {
                                    if (_fromCol <= 0 && cols[0].Equals(c.Name, StringComparison.OrdinalIgnoreCase))   //Issue15063 Add invariant igore case
                                    {
                                        _fromCol = col;
                                        if (cols.Length == 1)
                                        {
                                            _toCol = _fromCol;
                                            return;
                                        }
                                    }
                                    else if (cols.Length > 1 && _fromCol > 0 && cols[1].Equals(c.Name, StringComparison.OrdinalIgnoreCase)) //Issue15063 Add invariant igore case
                                    {
                                        _toCol = col;
                                        return;
                                    }

                                    col++;
                                }
                            }
                        }
                    }
                }
            }
        }
        internal string ChangeTableName(string prevName, string name)
        {
            if (LocalAddress.StartsWith(prevName +"[", StringComparison.CurrentCultureIgnoreCase))
            {
                var wsPart = "";
                var ix = _address.TrimEnd().LastIndexOf('!', _address.Length - 2);  //Last index can be ! if address is #REF!, so check from                 
                if (ix >= 0)
                {
                    wsPart=_address.Substring(0, ix);
                }

                return wsPart + name + LocalAddress.Substring(prevName.Length);
            }
            else
            {
                return _address;
            }
        }
        internal ExcelAddressBase Intersect(ExcelAddressBase address)
        {
            if(address._fromRow > _toRow || _toRow < address._fromRow ||
               address._fromCol > _toCol || _toCol < address._fromCol ||
               _fromRow > address._toRow || address._toRow < _fromRow ||
               _fromCol > address._toCol || address._toCol < _fromCol
               )
            {
                return null;
            }
            
            var fromRow = Math.Max(address._fromRow, _fromRow);
            var toRow = Math.Min(address._toRow, _toRow);
            var fromCol = Math.Max(address._fromCol, _fromCol);
            var toCol = Math.Min(address._toCol, _toCol);

            return new ExcelAddressBase(fromRow, fromCol, toRow, toCol);
        }
        /// <summary>
        /// Returns the parts of this address that not intersects with <paramref name="address"/>
        /// </summary>
        /// <param name="address">The address to intersect with</param>
        /// <returns>The addresses not intersecting with <paramref name="address"/></returns>
        internal ExcelAddressBase IntersectReversed(ExcelAddressBase address)
        {
            if (address._fromRow > _toRow || _toRow < address._fromRow ||
               address._fromCol > _toCol || _toCol < address._fromCol ||
               _fromRow > address._toRow || address._toRow < _fromRow ||
               _fromCol > address._toCol || address._toCol < _fromCol ||
               (string.IsNullOrEmpty(address._ws) == false && string.IsNullOrEmpty(_ws) == false && address._ws != _ws))
            {
                return this;
            }
            string retAddress = "";
            int fromRow = _fromRow, fromCol = _fromCol, toCol = _toCol;

            if (_fromCol < address._fromCol)
            {
                retAddress = GetAddress(fromRow, fromCol, _toRow, address._fromCol - 1) + ",";
                fromCol = address._fromCol;
            }

            if (_fromRow < address._fromRow)
            {
                retAddress += GetAddress(fromRow, fromCol, address._fromRow - 1, toCol) + ",";
                fromRow = address._fromRow;
            }

            if (_toCol > address._toCol)
            {
                retAddress += GetAddress(fromRow, address._toCol + 1, _toRow, toCol) + ",";
                toCol = address._toCol;
            }

            if (_toRow > address._toRow)
            {
                retAddress += GetAddress(address._toRow + 1, fromCol, _toRow, toCol) + ",";
            }
            return string.IsNullOrEmpty(retAddress) ? null : new ExcelAddressBase(retAddress.Substring(0, retAddress.Length - 1));
        }

        internal bool IsInside(ExcelAddressBase effectedAddress)
        {
            var c = Collide(effectedAddress);
            return c == ExcelAddressBase.eAddressCollition.Equal ||
                   c == ExcelAddressBase.eAddressCollition.Inside;
        }
        /// <summary>
        /// Address is an defined name
        /// </summary>
        /// <param name="address">the name</param>
        /// <param name="isName">Should always be true</param>
        internal ExcelAddressBase(string address, bool isName)
        {
            if (isName)
            {
                _address = address;
                _fromRow = -1;
                _fromCol = -1;
                _toRow = -1;
                _toCol = -1;
                _start = null;
                _end = null;
            }
            else
            {
                SetAddress(address, null, null);
            }
        }
        /// <summary>
        /// Sets the address
        /// </summary>
        /// <param name="address">The address</param>
        /// <param name="wb"></param>
        /// <param name="wsName"></param>
        protected internal void SetAddress(string address, ExcelWorkbook wb, string wsName)
        {
            address = address.Trim();
            if (address.Length > 0 && (address[0] == '\'' || address[0] == '['))
            {
                SetWbWs(address);
            }
            else
            {
                _address = address;
            }
            _addresses = null;
            if (_address.IndexOfAny(new char[] {',','!', '['}) > -1)
            {
                _firstAddress = null;
                //Advanced address. Including Sheet or multi or table.
                ExtractAddress(_address);
            }
            else
            {
                //Simple address
                GetRowColFromAddress(_address, out _fromRow, out _fromCol, out _toRow, out  _toCol, out _fromRowFixed, out _fromColFixed,  out _toRowFixed, out _toColFixed, wb, wsName);
                _start = null;
                _end = null;
            }
            _address = address;
            Validate();
        }

        internal ExcelAddressBase ToInternalAddress()
        {
            if(_address.StartsWith("["))
            {
                var ix = _address.IndexOf("]", 1);
                if (ix > 0)
                {
                    if(_address[ix+1]=='!')
                    {
                        ix++;
                    }
                    var a = _address.Substring(ix+1);
                    
                    return new ExcelAddressBase(a);
                }
                return this;
            }
            else
            {
                return this;
            }
        }

        /// <summary>
        /// Called when the address changes
        /// </summary>
        internal protected virtual void ChangeAddress()
        {
        }
        private void SetWbWs(string address)
        {
            int pos;
            if (address[0] == '[')
            {
                pos = address.IndexOf(']');
                _wb = address.Substring(1, pos - 1);                
                _ws = address.Substring(pos + 1);                
            }
            else
            {
                _wb = "";
                _ws = address;
            }
            if(_ws.StartsWith("'", StringComparison.OrdinalIgnoreCase))
            {
                pos = _ws.IndexOf("'",1, StringComparison.OrdinalIgnoreCase);
                while(pos>0 && pos+1<_ws.Length && _ws[pos+1]=='\'')
                {
                    _ws = _ws.Substring(0, pos) + _ws.Substring(pos+1);
                    pos = _ws.IndexOf("'", pos+1, StringComparison.OrdinalIgnoreCase);
                }
                if (pos>0)
                {
                    if(_ws.Length-1==pos)
                    {
                        _address = "A:XFD";
                    }
                    else if (_ws[pos+1]!='!')
                    {
                        throw new InvalidOperationException($"Address is not valid {address}. Missing ! after sheet name.");
                    }
                    else
                    {
                        _address = _ws.Substring(pos + 2);
                    }
                    _ws = _ws.Substring(1, pos-1);
                    if(_ws.StartsWith("["))
                    {
                        var ix = _ws.IndexOf("]", 1);
                        if(ix>0)
                        {
                            _wb = _ws.Substring(1, ix - 1);
                            _ws = _ws.Substring(ix+1);
                        }
                    }
                    pos = _address.IndexOf(":'", StringComparison.OrdinalIgnoreCase);
                    if(pos>0)
                    {
                        var a1 = _address.Substring(0,pos);
                        pos = _address.LastIndexOf("\'!", StringComparison.OrdinalIgnoreCase);
                        if (pos > 0)
                        {
                            var a2 = _address.Substring(pos+2);
                            _address=a1 + ":" + a2; //Remove any worksheet on second reference of the address. 
                        }
                    }
                    return;
                }
            }
            pos = _ws.IndexOf('!');

            if (pos==0)
            {
                _address = _ws.Substring(1);
                _ws = "";
                //_wb = "";
            }
            else if (pos > -1)
            {
                _address = _ws.Substring(pos + 1);
                _ws = _ws.Substring(0, pos);
            }
            else
            {
                _address = address;
            }
            if(string.IsNullOrEmpty(_address))
            {
                _address = "A:XFD";
            }
        }
        internal void ChangeWorksheet(string wsName, string newWs)
        {
            if (_ws == wsName) _ws = newWs;
            var fullAddress = GetAddress();
            
            if (Addresses != null)
            {
                foreach (var a in Addresses)
                {
                    if (a._ws == wsName)
                    {
                        a._ws = newWs;
                        fullAddress += "," + a.GetAddress();
                    }
                    else
                    {
                        fullAddress += "," + a._address;
                    }
                }
            }
            _address = fullAddress;
        }

        private string GetAddress()
        {
            string address = GetAddressWorkBookWorkSheet();
            if (IsName)
                return address + GetAddress(_fromRow, _fromCol, _toRow, _toCol);
            else
                return address + GetAddress(_fromRow, _fromCol, _toRow, _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
        }

        internal string GetAddressWorkBookWorkSheet()
        {
            var address = "";

            if (string.IsNullOrEmpty(_ws) == false)
            {
                if (string.IsNullOrEmpty(_wb) == false)
                {
                    address = "[" + _wb + "]";
                }

                if (_address.IndexOf("'!", StringComparison.OrdinalIgnoreCase) >=0 || ExcelWorksheet.NameNeedsApostrophes(_ws))
                {
                    address += string.Format("'{0}'!", _ws.Replace("'","''"));
                }
                else
                {
                    address += string.Format("{0}!", _ws);
                }
            }

            return address;
        }
        #endregion
        internal ExcelCellAddress _start = null;
        /// <summary>
        /// Gets the row and column of the top left cell.
        /// </summary>
        /// <value>The start row column.</value>
        public ExcelCellAddress Start
        {
            get
            {
                if (_start == null)
                {
                    _start = new ExcelCellAddress(_fromRow, _fromCol, _fromRowFixed, _fromColFixed);
                }
                return _start;
            }
        }
        internal ExcelCellAddress _end = null;
        /// <summary>
        /// Gets the row and column of the bottom right cell.
        /// </summary>
        /// <value>The end row column.</value>
        public ExcelCellAddress End
        {
            get
            {
                if (_end == null)
                {
                    _end = new ExcelCellAddress(_toRow, _toCol, _toRowFixed, _toColFixed);
                }
                return _end;
            }
        }
        /// <summary>
        /// The index to the external reference. Return 0, the current workbook, if no reference exists.
        /// </summary>
        public int ExternalReferenceIndex
        {
            get
            {
                if(Address.StartsWith("["))
                {
                   if(_wb.Any(x=>char.IsDigit(x)))
                   {
                      return int.Parse(_wb);
                   }
                   else
                   {
                      return -1;
                   }
                }
                else
                {
                    return 0;
                }
            }
        }
        internal ExcelTableAddress _table = null;
        /// <summary>
        /// If the address is refering a table, this property contains additional information 
        /// </summary>
        public ExcelTableAddress Table
        {
            get
            {
                return _table;
            }
        }

        /// <summary>
        /// The address for the range
        /// </summary>
        public virtual string Address
        {
            get
            {
                return _address;
            }
        }
        /// <summary>
        /// The full address including the worksheet
        /// </summary>
        public string FullAddress
        {
            get
            {
                string a="";
                if(_addresses != null)
                {
                    foreach(var sa in _addresses)
                    {
                        a += ","+sa.GetAddress();
                    }
                    a = a.TrimStart(',');
                }
                else
                {
                    a = GetAddress();
                }
                return a;
            }
        }
        /// <summary>
        /// If the address is a defined name
        /// </summary>
        public bool IsName
        {
            get
            {
                return _fromRow < 0;
            }
        }
        /// <summary>
        /// Returns the address text
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return _address;
        }
        /// <summary>
        /// Serves as the default hash function.
        /// </summary>
        /// <returns>A hash code for the current object.</returns>
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
        string _firstAddress;
        /// <summary>
        /// returns the first address if the address is a multi address.
        /// A1:A2,B1:B2 returns A1:A2
        /// </summary>
        internal string FirstAddress
        {
            get
            {
                if (string.IsNullOrEmpty(_firstAddress))
                {
                    return _address;
                }
                else
                {
                    return _firstAddress;
                }
            }
        }
        internal string AddressSpaceSeparated
        {
            get
            {
                return _address.Replace(',', ' '); //Conditional formatting and a few other places use space as separator for mulit addresses.
            }
        }
        /// <summary>
        /// Validate the address
        /// </summary>
        protected void Validate()
        {
            if ((_fromRow > _toRow || _fromCol > _toCol) && (_toRow!=0)) //_toRow==0 is #REF!
            {
                throw new ArgumentOutOfRangeException("Start cell Address must be less or equal to End cell address");
            }
        }
        internal string WorkSheetName
        {
            get
            {
                return _ws;
            }
        }
        internal List<ExcelAddress> _addresses = null;
        internal virtual List<ExcelAddress> Addresses
        {
            get
            {
                return _addresses;
            }
        }

        private bool ExtractAddress(string fullAddress)
        {
            var brackPos=new Stack<int>();
            var bracketParts=new List<string>();
            string first="", second="";
            bool isText=false, hasSheet=false, hasColon=false;
            string ws="";
            _addresses = null;            
            try
            {
                if (fullAddress == "#REF!")
                {
                    SetAddress(ref fullAddress, ref second, ref hasSheet);
                    return true;
                }
                else if (Utils.ConvertUtil._invariantCompareInfo.IsPrefix(fullAddress, "!"))
                {
                    // invalid address!
                    return false;
                }
                for (int i = 0; i < fullAddress.Length; i++)
                {
                    var c = fullAddress[i];
                    if (c == '\'')
                    {
                        if (isText && i + 1 < fullAddress.Length && fullAddress[i + 1] == '\'')
                        {
                            if (hasColon)
                            {
                                second += c;
                            }
                            else
                            {
                                first += c;
                            }
                        }
                        isText = !isText;
                    }
                    else
                    {
                        if (brackPos.Count > 0)
                        {
                            if (c == '[' && !isText)
                            {
                                brackPos.Push(i);
                            }
                            else if (c == ']' && !isText)
                            {
                                if (brackPos.Count > 0)
                                {
                                    var from = brackPos.Pop();
                                    bracketParts.Add(fullAddress.Substring(from + 1, i - from - 1));

                                    if (brackPos.Count == 0)
                                    {
                                        HandleBrackets(first, second, bracketParts);
                                    }
                                }
                                else
                                {
                                    //Invalid address!
                                    return false;
                                }
                            }
                        }
                        else if (c == ':' && !isText)
                        {
                            hasColon = true;
                        }
                        else if (c == '[' && !isText)
                        {
                            brackPos.Push(i);
                        }
                        else if (c == '!' && !isText && !first.EndsWith("#REF") && !second.EndsWith("#REF"))
                        {
                            // the following is to handle addresses that specifies the
                            // same worksheet twice: Sheet1!A1:Sheet1:A3
                            // They will be converted to: Sheet1!A1:A3
                            if (hasSheet && second != null && second.ToLower().EndsWith(first.ToLower()))
                            {
                                second = Regex.Replace(second, $"{first}$", string.Empty);
                            }
                            if (string.IsNullOrEmpty(ws))
                            {                                
                                if (second == "")
                                {
                                    ws = first;
                                    first = "";
                                }
                                else
                                {
                                    ws = second;
                                    second = "";
                                }
                            }
                            else if(string.IsNullOrEmpty(second)==false)
                            {
                                if(!ws.Equals(second,StringComparison.OrdinalIgnoreCase))
                                {
                                    _fromRow = _toRow = _fromCol = _toCol = -1;
                                    return true;
                                }
                                second = "";
                            }
                            hasSheet = true;
                        }
                        else if (c == ',' && !isText)
                        {
                            if(_addresses==null) _addresses = new List<ExcelAddress>();
                            if(string.IsNullOrEmpty(ws))
                            {
                                first = string.IsNullOrEmpty(second) ? first : first + ":" + second;
                                second = "";
                            }
                            else
                            {
                                second = string.IsNullOrEmpty(second) ? first : first + ":" + second;
                                first = ws;
                            }
                            SetAddress(ref first, ref second, ref hasSheet);
                            ws = "";
                            hasSheet = false;
                            hasColon = false;
                        }
                        else
                        {
                            if (hasColon)
                            {
                                second += c;
                            }
                            else
                            {
                                first += c;
                            }
                        }
                    }
                }
                if (Table == null)
                {
                    if (string.IsNullOrEmpty(ws))
                    {
                        first = string.IsNullOrEmpty(second) ? first : first + ":" + second;
                        second = "";
                    }
                    else
                    {
                        second = string.IsNullOrEmpty(second) ? first : first + ":" + second;
                        first = ws;
                    }

                    SetAddress(ref first, ref second, ref hasSheet);
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void HandleBrackets(string first, string second, List<string> bracketParts)
        {
            if(!string.IsNullOrEmpty(first))
            {
                _table = new ExcelTableAddress();
                Table.Name = first;
                foreach (var s in bracketParts)
                {
                    if(s.IndexOf('[')<0)
                    {
                        switch(s.ToLower(CultureInfo.InvariantCulture))                
                        {
                            case "#all":
                                _table.IsAll = true;
                                break;
                            case "#headers":
                               _table.IsHeader = true;
                                break;
                            case "#data":
                                _table.IsData = true;
                                break;
                            case "#totals":
                                _table.IsTotals = true;
                                break;
                            case "#this row":
                                _table.IsThisRow = true;
                                break;
                            default:
                                if(string.IsNullOrEmpty(_table.ColumnSpan))
                                {
                                    _table.ColumnSpan=s;
                                }
                                else
                                {
                                    _table.ColumnSpan += ":" + s;
                                }
                                break;
                        }                
                    }
                }
            }
        }
        #region Address manipulation methods
        internal eAddressCollition Collide(ExcelAddressBase address, bool ignoreWs=false)
        {
            if (ignoreWs == false && address.WorkSheetName != WorkSheetName && 
                string.IsNullOrEmpty(address.WorkSheetName) == false && 
                string.IsNullOrEmpty(WorkSheetName) == false)
            {
                return eAddressCollition.No;
            }

            return Collide(address._fromRow, address._fromCol, address._toRow, address._toCol);
        }
        internal eAddressCollition Collide(int row, int col)
        {
            return Collide(row, col, row, col);
        }
        internal eAddressCollition Collide(int fromRow, int fromCol, int toRow, int toCol)
        {
            if (DoNotCollide(fromRow, fromCol, toRow, toCol))
            {
                return eAddressCollition.No;
            }
            else if (fromRow == _fromRow && fromCol == _fromCol &&
                    toRow == _toRow && toCol == _toCol)
            {
                return eAddressCollition.Equal;
            }
            else if (fromRow >= _fromRow && toRow <= _toRow &&
                     fromCol >= _fromCol && toCol <= _toCol)
            {
                return eAddressCollition.Inside;
            }
            else
                return eAddressCollition.Partly;
        }

        internal bool DoNotCollide(int fromRow, int fromCol, int toRow, int toCol)
        {
            return fromRow > _toRow || fromCol > _toCol
                   ||
                   _fromRow > toRow || _fromCol > toCol;
        }

        internal bool CollideFullRowOrColumn(ExcelAddressBase address)
        {
            return CollideFullRowOrColumn(address._fromRow, address._fromCol, address._toRow, address._toCol);
        }
        internal bool CollideFullRowOrColumn(int fromRow, int fromCol, int toRow, int toCol)
        {
            return (CollideFullRow(fromRow, toRow) && CollideColumn(fromCol, toCol)) || 
                   (CollideFullColumn(fromCol, toCol) && CollideRow(fromRow, toRow));
        }
        private bool CollideColumn(int fromCol, int toCol)
        {
            return fromCol  <= _toCol && toCol >= _fromCol;
        }

        internal bool CollideRow(int fromRow, int toRow)
        {
            return fromRow <= _toRow && toRow >= _fromRow;
        }
        internal bool CollideFullRow(int fromRow, int toRow)
        {
            return fromRow <= _fromRow && toRow >= _toRow;
        }
        internal bool CollideFullColumn(int fromCol, int toCol)
        {
            return fromCol <= _fromCol && toCol >= _toCol;
        }
        internal ExcelAddressBase AddRow(int row, int rows, bool setFixed=false, bool setRefOnMinMax=true, bool extendIfLastRow=false)
        {
            if (row > _toRow && (row!=_toRow+1 || extendIfLastRow==false))
            {
                return this;
            }
            var toRow = setFixed && _toRowFixed ? _toRow : _toRow + rows;
            if (toRow < 1) return null;
            if (row <= _fromRow)
            {
                var fromRow = setFixed && _fromRowFixed ? _fromRow : _fromRow + rows;
                if (fromRow > ExcelPackage.MaxRows) return null;
                return new ExcelAddressBase(GetRow(fromRow, setRefOnMinMax), _fromCol, GetRow(toRow, setRefOnMinMax), _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed, WorkSheetName, _address);
            }
            else
            {
                return new ExcelAddressBase(_fromRow, _fromCol, GetRow(toRow, setRefOnMinMax), _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed, WorkSheetName, _address);
            }
        }

        private int GetRow(int row, bool setRefOnMinMax)
        {
            if (setRefOnMinMax==false)
            {
                if (row < 1) return 1;
                if (row > ExcelPackage.MaxRows) return ExcelPackage.MaxRows;
            }

            return row;
        }
        private int GetColumn(int column, bool setRefOnMinMax)
        {
            if (setRefOnMinMax == false)
            {
                if (column < 1) return 1;
                if (column > ExcelPackage.MaxColumns) return ExcelPackage.MaxColumns;
            }

            return column;
        }

        internal ExcelAddressBase DeleteRow(int row, int rows, bool setFixed = false, bool adjustMaxRow=true)
        {
            if (row > _toRow) //After
            {
                return this;
            }
            else if (row != 0 && row <= _fromRow && row + rows > _toRow) //Inside
            {
                return null;
            }
            else if (row+rows < _fromRow || (_fromRowFixed && row < _fromRow)) //Before
            {
                var toRow = ((setFixed && _toRowFixed) || (adjustMaxRow==false && _toRow==ExcelPackage.MaxRows)) ? _toRow : _toRow - rows;
                return new ExcelAddressBase((setFixed && _fromRowFixed ? _fromRow : _fromRow - rows), _fromCol, toRow, _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed, WorkSheetName, _address);
            }
            else  //Partly
            {
                if (row <= _fromRow)
                {
                    var toRow = (setFixed && _toRowFixed) || (adjustMaxRow == false && _toRow == ExcelPackage.MaxRows) ? _toRow : _toRow - rows;

                    return new ExcelAddressBase(row, _fromCol, toRow, _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed, WorkSheetName, _address);
                }
                else
                {
                    var toRow = (setFixed && _toRowFixed) || (adjustMaxRow == false && _toRow == ExcelPackage.MaxRows) ? _toRow : _toRow - rows < row ? row - 1 : _toRow - rows;
                    return new ExcelAddressBase(_fromRow, _fromCol, toRow, _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed, WorkSheetName, _address);
                }
            }
        }
        internal ExcelAddressBase AddColumn(int col, int cols, bool setFixed = false, bool setRefOnMinMax=true)
        {
            if (col > _toCol)
            {
                return this;
            }
            var toCol = GetColumn((setFixed && _toColFixed ? _toCol : _toCol + cols), setRefOnMinMax);
            if (col <= _fromCol)
            {
                var fromCol = GetColumn((setFixed && _fromColFixed ? _fromCol : _fromCol + cols), setRefOnMinMax);
                return new ExcelAddressBase(_fromRow, fromCol, _toRow, toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed, WorkSheetName, _address);
            }
            else
            {
                return new ExcelAddressBase(_fromRow, _fromCol, _toRow, toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed, WorkSheetName, _address);
            }
        }
        internal ExcelAddressBase DeleteColumn(int col, int cols, bool setFixed = false, bool adjustMaxCol = true)
        {
            if (col > _toCol) //After
            {
                return this;
            }
            if (col!=0 && col <= _fromCol && col + cols > _toCol) //Inside
            {
                return null;
            }
            else if (col + cols < _fromCol || _fromColFixed && col < _fromCol) //Before
            {
                var toCol = ((setFixed && _toColFixed) ||(adjustMaxCol==false && _toCol==ExcelPackage.MaxColumns)) ? _toCol : _toCol - cols;
                return new ExcelAddressBase(_fromRow, (setFixed && _fromColFixed ? _fromCol : _fromCol - cols), _toRow, toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed, WorkSheetName, _address);
            }
            else  //Partly
            {
                if (col <= _fromCol)
                {
                    var toCol = ((setFixed && _toColFixed) || (adjustMaxCol == false && _toCol == ExcelPackage.MaxColumns)) ? _toCol : _toCol - cols;
                    return new ExcelAddressBase(_fromRow, col, _toRow, toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed, _ws, _address);
                }
                else
                {
                    var toCol = ((setFixed && _toColFixed) || (adjustMaxCol == false && _toCol == ExcelPackage.MaxColumns)) ? _toCol : _toCol - cols < col ? col - 1 : _toCol - cols;
                    return new ExcelAddressBase(_fromRow, _fromCol, _toRow, toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed, _ws, _address);
                }
            }
        }
        internal ExcelAddressBase Insert(ExcelAddressBase address, eShiftTypeInsert Shift)
        {
            //Before or after, no change
            if(_toRow < address._fromRow || _toCol < address._fromCol || (_fromRow > address._toRow && _fromCol > address._toCol))
            {
                return this;
            }

            int rows = address.Rows;
            int cols = address.Columns;
            string retAddress = "";
            if (Shift==eShiftTypeInsert.Right)
            {
                if (address._fromRow > _fromRow)
                {
                    retAddress = GetAddress(_fromRow, _fromCol, address._fromRow, _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
                }
                if(address._fromCol > _fromCol)
                {
                    retAddress = GetAddress(_fromRow < address._fromRow ? _fromRow : address._fromRow, _fromCol, address._fromRow, _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
                }
            }
            if (_toRow < address._fromRow)
            {
                if (_fromRow < address._fromRow)
                {

                }
                else
                {
                }
            }
            return null;
        }
        #endregion
        private void SetAddress(ref string first, ref string second, ref bool hasSheet)
        {
            string ws, address;
            if (hasSheet)
            {
                ws = first;
                address = second;
                first = "";
                second = "";
            }
            else
            {
                address = first;
                ws = "";
                first = "";
            }
            hasSheet = false;
            if (string.IsNullOrEmpty(_firstAddress))
            {
                if (string.IsNullOrEmpty(_ws) || !string.IsNullOrEmpty(ws))
                {
                    _ws = ws;                    
                }
                _firstAddress = address;
                GetRowColFromAddress(address, out _fromRow, out _fromCol, out _toRow, out  _toCol, out _fromRowFixed, out _fromColFixed, out _toRowFixed, out _toColFixed);
                _start = null;
                _end = null;
            }
            if (_addresses != null)
            {
                _addresses.Add(new ExcelAddress(_ws, address));
            }
        }
        internal enum AddressType
        {
            Invalid,
            InternalAddress,
            ExternalAddress,
            InternalName,
            ExternalName,
            Formula,
            R1C1
        }

        internal static AddressType IsValid(string Address, bool r1c1=false)
        {
            double d;
            if (Address == "#REF!")
            {
                return AddressType.Invalid;
            }
            else if(double.TryParse(Address, NumberStyles.Any, CultureInfo.InvariantCulture, out d)) //A double, no valid address
            {
                return AddressType.Invalid;
            }
            else if (IsFormula(Address))
            {
                return AddressType.Formula;
            }
            else
            {
                if (r1c1 && IsR1C1(Address))
                {
                    return AddressType.R1C1;
                }
                else
                {
                    string wb, ws, intAddress;
                    if (SplitAddress(Address, out wb, out ws, out intAddress))
                    {

                        if (intAddress.Contains("[")) //Table reference
                        {
                            return string.IsNullOrEmpty(wb) ? AddressType.InternalAddress : AddressType.ExternalAddress;
                        }
                        if (intAddress.Contains(","))
                        {
                            intAddress = intAddress.Substring(0, intAddress.IndexOf(','));
                        }
                        if (IsAddress(intAddress, true))
                        {
                            return string.IsNullOrEmpty(wb) ? AddressType.InternalAddress : AddressType.ExternalAddress;
                        }
                        else
                        {
                            return string.IsNullOrEmpty(wb) ? AddressType.InternalName : AddressType.ExternalName;
                        }
                    }
                    else
                    {
                        return AddressType.Invalid;
                    }
                }
            }
        }
        private static bool IsR1C1(string address)
        {
            var start = address.LastIndexOf("!", address.Length-1, StringComparison.OrdinalIgnoreCase);
            if (start>=0)
            {
                address = address.Substring(start + 1);
            }
            address = address.ToUpper();
            if (string.IsNullOrEmpty(address) || (address[0]!='R' && address[0]!='C'))
            {
                return false;
            }
            bool isC = false, isROrC = false;
            bool startBracket = false;
            foreach(var c in address)
            {
                switch(c)
                {
                    case 'C':
                        isC = true;
                        isROrC = true;
                        break;
                    case 'R':
                        if (isC)
                            return false;
                        isROrC = true;
                        break;
                    case '[':
                        startBracket = true;
                        break;
                    case ']':
                        if (startBracket == false) return false;
                        isROrC = false;
                        break;
                    case ':':
                        isC = false;
                        startBracket = false;
                        isROrC = false;
                        break;
                    default:
                        if((c>='0' && c<='9') ||c=='-')
                        {
                            if(isROrC==false)
                            {
                                return false;
                            }
                        }
                        else
                        {
                            return false;
                        }
                        break;
                }
            }
            return true;
        }

        private static bool IsAddress(string intAddress, bool allowRef = false)
        {
            if(string.IsNullOrEmpty(intAddress)) return false;            
            var cells = intAddress.Split(':');
            int fromRow, toRow, fromCol, toCol;

            if(!GetRowCol(cells[0], out fromRow, out fromCol, false))
            {
                return false;
            }
            if (cells.Length > 1)
            {
                if (!GetRowCol(cells[1], out toRow, out toCol, false))
                {
                    return false;
                }
            }
            else
            {
                toRow = fromRow;
                toCol = fromCol;
            }
            if (allowRef)
            {
                return
                    fromCol > -1 &&
                    toCol <= ExcelPackage.MaxColumns &&
                    fromRow > -1 &&
                    toRow <= ExcelPackage.MaxRows;
            }
            else
            {
                return 
                    fromRow <= toRow &&
                    fromCol <= toCol &&
                    fromCol > -1 &&
                    toCol <= ExcelPackage.MaxColumns &&
                    fromRow > -1 &&
                    toRow <= ExcelPackage.MaxRows;
            }
        }

        private static bool SplitAddress(string Address, out string wb, out string ws, out string intAddress)
        {
            wb = "";
            ws = "";
            intAddress = "";
            var text = "";
            bool isText = false;
            var brackPos=-1;
            for (int i = 0; i < Address.Length; i++)
            {
                if (Address[i] == '\'')
                {
                    isText = !isText;
                    if(i>0 && Address[i-1]=='\'')
                    {
                        text += "'";
                    }
                }
                else
                {
                    if(Address[i]=='!' && !isText)
                    {
                        if (text.Length>0 && text[0] == '[')
                        {
                            wb = text.Substring(1, text.IndexOf(']') - 1);
                            ws = text.Substring(text.IndexOf(']') + 1);
                        }
                        else
                        {
                            ws=text;
                        }
                        intAddress=Address.Substring(i+1);
                        return true;
                    }
                    else
                    {
                        if(Address[i]=='[' && !isText)
                        {
                            if (i > 0) //Table reference return full address;
                            {
                                intAddress=Address;
                                return true;
                            }
                            brackPos=i;
                        }
                        else if(Address[i]==']' && !isText)
                        {
                            if (brackPos > -1)
                            {
                                wb = text;
                                text = "";
                            }
                            else
                            {
                                return false;
                            }
                        }
                        else
                        {
                            text+=Address[i];
                        }
                    }
                }
            }
            intAddress = text;
            return true;
        }

        private static readonly HashSet<char> _tokens = new HashSet<char>(new char[] { '+', '-', '*', '/', '^', '&', '=', '<', '>', '(', ')', '{', '}', '%', '\"' }); //See TokenSeparatorProvider
        internal static bool IsFormula(string address)
        {
            var isText = false;
            var tableNameCount = 0;
            for (int i = 0; i < address.Length; i++)
            {
                var addressChar = address[i];
                if (addressChar == '\'')
                {
                    if(i>0 && isText==false && address.Length>i+1 && address[i - 1] == ' ' && address[i+1] != '\'')
                    {
                        return true;
                    }
                    isText = !isText;
                }
                else if (isText == false && addressChar == '[')
                    tableNameCount++;
                else if (isText == false && addressChar == ']')
                    tableNameCount--;
                else if(tableNameCount==0)
                {
                    if (isText == false && _tokens.Contains(addressChar))
                    {
                        return true;
                    }
                }
            }
            return false;
        }
        private static bool IsValidName(string address)
        {
            if (Regex.IsMatch(address, "[^0-9./*-+,½!\"@#£%&/{}()\\[\\]=?`^~':;<>|][^/*-+,½!\"@#£%&/{}()\\[\\]=?`^~':;<>|]*"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Number of rows int the address
        /// </summary>
        public int Rows 
        {
            get
            {
                return _toRow - _fromRow+1;
            }
        }
        /// <summary>
        /// Number of columns int the address
        /// </summary>
        public int Columns
        {
            get
            {
                return _toCol - _fromCol + 1;
            }
        }
        /// <summary>
        /// Returns true if the range spans a full row
        /// </summary>
        /// <returns></returns>
        public bool IsFullRow
        {
            get
            {
                return _fromCol == 1 && _toCol == ExcelPackage.MaxColumns;
            }
        }
        /// <summary>
        /// Returns true if the range spans a full column
        /// </summary>
        /// <returns></returns>
        public bool IsFullColumn
        {
            get
            {
                return _fromRow == 1 && _toRow == ExcelPackage.MaxRows;
            }
        }

        internal bool IsSingleCell
        {
            get
            {
                return (_fromRow == _toRow && _fromCol == _toCol);
            }
        }

        /// <summary>
        /// The address without the workbook or worksheet reference
        /// </summary>
        public string LocalAddress 
        { 
            get
            {                
                if (Addresses == null)
                {
                    if (_table == null)
                    {
                        return GetAddress(_fromRow, _fromCol, _toRow, _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
                    }
                    else
                    {
                        return RemoveSheetName(FirstAddress);
                    }
                }
                else
                {
                    var sb = new StringBuilder();
                    foreach (var a in Addresses)
                    {
                        if (a._table == null)
                        {
                            sb.Append(GetAddress(a._fromRow, a._fromCol, a._toRow, a._toCol, a._fromRowFixed, a._fromColFixed, a._toRowFixed, a._toColFixed));
                        }
                        else
                        {
                            sb.Append(RemoveSheetName(a.Address));
                        }
                        sb.Append(",");
                    }
                    return sb.ToString(0, sb.Length - 1);
                }
            }
        }

        private static string RemoveSheetName(string address)
        {
            var ix = address.TrimEnd().LastIndexOf('!', address.Length - 2);  //Last index can be ! if address is #REF!, so check from 
            if (ix >= 0)
            {
                address = address.Substring(ix + 1);
            }

            return address;
        }

        /// <summary>
        /// The address without the workbook reference
        /// </summary>
        internal string WorkbookLocalAddress
        {
            get
            {
                if (!_address.StartsWith("[")) return _address;
                var ix = _address.IndexOf("]",1);
                if (ix >= 0)
                {
                    return _address.Substring(ix + 1);
                }
                return _address;
            }
        }

        internal static string GetWorkbookPart(string address)
        {
            var ix = 0;
            if(address[ix]=='\'')
            {
                ix++;
            }
            if (address[ix] == '[')
            {
                var endIx = address.LastIndexOf(']');
                if (endIx > 0)
                {
                    return address.Substring(ix+1, endIx - ix - 1);
                }   
            }
            return "";
        }
        internal static string GetWorksheetPart(string address, string defaultWorkSheet)
        {
            int ix=0;
            return GetWorksheetPart(address, defaultWorkSheet, ref ix);
        }
        internal static string GetWorksheetPart(string address, string defaultWorkSheet, ref int endIx)
        {
            if(address=="") return defaultWorkSheet;
            var ix = 0;
            if (address[0] == '[' || address.StartsWith("'["))
            {
                ix = address.IndexOf(']')+1;
            }
            if (ix >= 0 && ix < address.Length)
            {
                if (address[ix] == '\'')
                {
                    var ret=GetString(address, ix+1, out endIx);
                    endIx++;
                    return ret; 
                }
                else
                {
                    endIx = address.IndexOf('!',ix)+1;
                    var subtrLen = 1;
                    if(endIx>0 && address[endIx-2]=='\'')
                    {
                        subtrLen++;
                    }
                    if(endIx > ix)
                    {
                        return address.Substring(ix, endIx - ix - subtrLen);
                    }   
                    else
                    {
                        return defaultWorkSheet;
                    }
                }
            }
            else
            {
                return defaultWorkSheet;
            }
        }
        internal static string GetAddressPart(string address)
        {
            var ix=0;
            GetWorksheetPart(address, "", ref ix);
            if(ix<address.Length)
            {
                if (address[ix] == '!')
                {
                    return address.Substring(ix + 1);
                }
                else
                {
                    return "";
                }
            }
            else
            {
                return "";
            }

        }
        internal static void SplitAddress(string fullAddress, out string wb, out string ws, out string address, string defaultWorksheet="")
        {
            wb = GetWorkbookPart(fullAddress);
            int ix=0;
            ws = GetWorksheetPart(fullAddress, defaultWorksheet, ref ix);
            if (ix < fullAddress.Length)
            {
                if (fullAddress[ix] == '!')
                {
                    address = fullAddress.Substring(ix + 1);
                }
                else
                {
                    address = fullAddress.Substring(ix);
                }
            }
            else
            {
                address="";
            }
        }
        internal static List<string[]> SplitFullAddress(string fullAddress)
        {
            var addresses = new List<string[]>();
            var currentAddress = new string[3];
            bool isInWorkbook = false;
            bool isInWorksheet = false;
            bool isInAddress = false;
            bool isInText = false;
            var prevPos = 0;
            for (int i=0;i<fullAddress.Length;i++)
            {
                if (isInWorkbook == false &&
                    isInWorksheet == false &&
                    isInAddress == false)
                {
                    if (fullAddress[i] == '[')
                    {
                        isInWorkbook = true;
                        prevPos = i + 1;
                    }
                    else if(fullAddress[i] == '\'')
                    {
                        isInWorksheet = true;
                        isInText = true;
                        prevPos = i + 1;
                    }
                    else if (fullAddress[i]=='!')
                    {
                        isInAddress = true;
                        prevPos = i + 1;
                    }
                    else
                    {
                        isInAddress = true;
                    }
                }
                else if(isInWorkbook)
                {
                    if (fullAddress[i] == ']')
                    {
                        currentAddress[0] = fullAddress.Substring(prevPos, i - prevPos);
                        isInWorkbook = false;
                    }
                }
                else if(isInWorksheet)
                {
                    if (fullAddress[i] == '\'')
                    {
                        isInText = !isInText;
                    }
                    else if (isInText==false && fullAddress[i] == '!')
                    {
                        currentAddress[1] = fullAddress.Substring(prevPos, i -prevPos - 1).Replace("''","'");
                        prevPos = i + 1;
                        isInWorksheet = false;
                    }
                }
                else if(isInAddress)
                {
                    if(fullAddress[i] == '!')
                    {
                        currentAddress[1] = fullAddress.Substring(prevPos, i - prevPos);
                        prevPos = i + 1;
                    }
                    else if (fullAddress[i]==',')
                    {
                        currentAddress[2] = fullAddress.Substring(prevPos, i - prevPos);
                        addresses.Add(currentAddress);
                        prevPos = i + 1;
                        isInAddress = false;
                        currentAddress = new string[3];
                    }
                }
            }

            if(isInWorkbook || isInWorksheet)
            {
                throw (new ArgumentException($"Invalid address {fullAddress}"));
            }
            currentAddress[2] = fullAddress.Substring(prevPos, fullAddress.Length - prevPos);
            addresses.Add(currentAddress);
            return addresses;
        }
        private static string GetString(string address, int ix, out int endIx)
        {
            var strIx = address.IndexOf("''", ix);
            var prevStrIx = ix;
            while (strIx > -1)
            {
                prevStrIx = strIx;
                strIx = address.IndexOf("''", strIx + 1);
            }
            endIx = address.IndexOf("'", prevStrIx + 1) + 1;
            return address.Substring(ix, endIx - ix - 1).Replace("''", "'");
        }

        internal bool IsValidRowCol()
        {
            return !(_fromRow > _toRow  ||
                   _fromCol > _toCol ||
                   _fromRow < 1 ||
                   _fromCol < 1 ||
                   _toRow > ExcelPackage.MaxRows ||
                   _toCol > ExcelPackage.MaxColumns);
        }
        /// <summary>
        /// Returns true if the item is equal to another item.
        /// </summary>
        /// <param name="obj">The item to compare</param>
        /// <returns>True if the items are equal</returns>
        public override bool Equals(object obj)
        {
            if (obj is ExcelAddressBase a)
            {
                if (Addresses==null || a.Addresses==null)
                {
                    if (Addresses?.Count > 1 || a.Addresses?.Count > 1) return false;
                    return IsEqual(this, a);
                }
                else
                {
                    if (Addresses.Count != a.Addresses.Count) return false;
                    for(int i=0;i<Addresses.Count;i++)
                    {
                        if (IsEqual(Addresses[i], a.Addresses[i]) == false)
                        {
                            return false;
                        }
                    }
                    return true;
                }
            }
            else
            {
                return _address == obj?.ToString();
            }
        }

        private bool IsEqual(ExcelAddressBase a1, ExcelAddressBase a2)
        {
            return a1._fromRow == a2._fromRow &&
                    a1._toRow == a2._toRow &&
                    a1._fromCol == a2._fromCol &&
                    a1._toCol == a2._toCol;
        }
        /// <summary>
        /// Returns true the address contains an external reference
        /// </summary>
        public bool IsExternal
        {
            get
            {
                return !string.IsNullOrEmpty(_wb);
            }
        }
    }
}

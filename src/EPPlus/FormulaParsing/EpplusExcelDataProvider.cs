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

namespace OfficeOpenXml.FormulaParsing
{
    public class EpplusExcelDataProvider : ExcelDataProvider
    {
        public class RangeInfo : IRangeInfo
        {
            internal ExcelWorksheet _ws;
            CellStoreEnumerator<ExcelValue> _values = null;
            int _fromRow, _toRow, _fromCol, _toCol;
            int _cellCount = 0;
            ExcelAddressBase _address;
            ICellInfo _cell;

            public RangeInfo(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
            {
                var address = new ExcelAddressBase(fromRow, fromCol, toRow, toCol);
                address._ws = ws.Name;                
                SetAddress(ws, address);
            }

            public RangeInfo(ExcelWorksheet ws, ExcelAddressBase address)
            {
                SetAddress(ws, address);
            }
            private void SetAddress(ExcelWorksheet ws, ExcelAddressBase address)
            {
                _ws = ws;
                _fromRow = address._fromRow;
                _fromCol = address._fromCol;
                _toRow = address._toRow;
                _toCol = address._toCol;
                _address = address;
                _values = new CellStoreEnumerator<ExcelValue>(_ws._values, _fromRow, _fromCol, _toRow, _toCol);
                _cell = new CellInfo(_ws, _values);
            }

            public int GetNCells()
            {
                return ((_toRow - _fromRow) + 1) * ((_toCol - _fromCol) + 1);
            }

            public bool IsEmpty
            {
                get
                {
                    if (_cellCount > 0)
                    {
                        return false;
                    }
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

            public ICellInfo Current
            {
                get { return _cell; }
            }

            public ExcelWorksheet Worksheet
            {
                get { return _ws; }
            }

            public void Dispose()
            {
                //_values = null;
                //_ws = null;
                //_cell = null;
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
                _cellCount++;
                return _values.MoveNext();
            }

            public void Reset()
            {
                _cellCount = 0;
                _values.Init();
            }


            public bool NextCell()
            {
                _cellCount++;
                return _values.MoveNext();
            }

            public IEnumerator<ICellInfo> GetEnumerator()
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
                return _ws.GetValue(row, col);
            }

            public object GetOffset(int rowOffset, int colOffset)
            {
                if (_values.Row < _fromRow || _values.Column < _fromCol)
                {
                    return _ws.GetValue(_fromRow + rowOffset, _fromCol + colOffset);
                }
                else
                {
                    return _ws.GetValue(_values.Row + rowOffset, _values.Column + colOffset);
                }
            }
        }

        public class CellInfo : ICellInfo
        {
            ExcelWorksheet _ws;
            CellStoreEnumerator<ExcelValue> _values;
            internal CellInfo(ExcelWorksheet ws, CellStoreEnumerator<ExcelValue> values)
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
                    return _ws.GetFormula(_values.Row, _values.Column);
                }
            }

            public object Value
            {
                get 
                { 
                    if(_ws._flags.GetFlagValue(_values.Row, _values.Column, CellFlags.RichText))
                    {
                        return _ws.GetRichText(_values.Row, _values.Column, null).Text;
                    }
                    else
                    {
                        return _values.Value._value;
                    }
                }
            }
            
            public double ValueDouble
            {
                get { return ConvertUtil.GetValueDouble(_values.Value._value, true); }
            }
            public double ValueDoubleLogical
            {
                get { return ConvertUtil.GetValueDouble(_values.Value._value, false); }
            }
            public bool IsHiddenRow
            {
                get 
                { 
                    var row=_ws.GetValueInner(_values.Row, 0) as RowInternal;
                    if(row != null)
                    {
                        return row.Hidden || row.Height==0;
                    }
                    else
                    {
                        return false;
                    }
                }
            }

            public bool IsExcelError
            {
                get { return ExcelErrorValue.Values.IsErrorValue(_values.Value._value); }
            }

            public IList<Token> Tokens
            {
                get 
                {
                    return _ws._formulaTokens.GetValue(_values.Row, _values.Column);
                }
            }

        }
        public class NameInfo : ExcelDataProvider.INameInfo
        {
            public ulong Id { get; set; }
            public string Worksheet { get; set; }
            public string Name { get; set; }
            public string Formula { get; set; }
            public IList<Token> Tokens { get; internal set; }
            public object Value { get; set; }
        }

        private readonly ExcelPackage _package;
        private ExcelWorksheet _currentWorksheet;
        private RangeAddressFactory _rangeAddressFactory;
        private Dictionary<ulong, INameInfo> _names=new Dictionary<ulong,INameInfo>();

        public EpplusExcelDataProvider(ExcelPackage package)
        {
            _package = package;

            _rangeAddressFactory = new RangeAddressFactory(this);
        }

        public override IEnumerable<string> GetWorksheets()
        {
            return _package.Workbook.Worksheets.Select(x => x.Name);
        }

        public override ExcelNamedRangeCollection GetWorksheetNames(string worksheet)
        {
            var ws=_package.Workbook.Worksheets[worksheet];
            if (ws != null)
            {
                return ws.Names;
            }
            else
            {
                return null;
            }
        }

        public override int GetWorksheetIndex(string worksheetName)
        {
            for (var ix = 1; ix <= _package.Workbook.Worksheets.Count; ix++)
            {
                var ws = _package.Workbook.Worksheets[ix - 1];
                if (string.Compare(worksheetName, ws.Name, true) == 0)
                {
                    return ix;
                }
            }
            return -1;
        }

        public override ExcelNamedRangeCollection GetWorkbookNameValues()
        {
            return _package.Workbook.Names;
        }

        public override IRangeInfo GetRange(string worksheet, int fromRow, int fromCol, int toRow, int toCol)
        {
            SetCurrentWorksheet(worksheet);
            var wsName = string.IsNullOrEmpty(worksheet) ? _currentWorksheet.Name : worksheet;
            var ws = _package.Workbook.Worksheets[wsName];
            return new RangeInfo(ws, fromRow, fromCol, toRow, toCol);
        }
        public override IRangeInfo GetRange(string worksheet, int row, int column, string address)
        {
            var addr = new ExcelAddress(worksheet, address);
            if (addr.Table != null)
            {
                addr = ConvertToA1C1(addr);
            }
            //SetCurrentWorksheet(addr.WorkSheet); 
            var wsName = string.IsNullOrEmpty(addr.WorkSheetName) ? _currentWorksheet.Name : addr.WorkSheetName;
            var ws = _package.Workbook.Worksheets[wsName];
            //return new CellsStoreEnumerator<object>(ws._values, addr._fromRow, addr._fromCol, addr._toRow, addr._toCol);
            return new RangeInfo(ws, addr);
        }
        public override IRangeInfo GetRange(string worksheet, string address)
        {
            var addr = new ExcelAddress(worksheet, address);
            if (addr.Table != null)
            {
                addr = ConvertToA1C1(addr);
            }
            //SetCurrentWorksheet(addr.WorkSheet); 
            var wsName = string.IsNullOrEmpty(addr.WorkSheetName) ? _currentWorksheet.Name : addr.WorkSheetName;
            var ws = _package.Workbook.Worksheets[wsName];
            //return new CellsStoreEnumerator<object>(ws._values, addr._fromRow, addr._fromCol, addr._toRow, addr._toCol);
            return new RangeInfo(ws, addr);
        }

        private ExcelAddress ConvertToA1C1(ExcelAddress addr)
        {
            //Convert the Table-style Address to an A1C1 address
            addr.SetRCFromTable(_package, addr);
            var a = new ExcelAddress(addr._fromRow, addr._fromCol, addr._toRow, addr._toCol);
            a._ws = addr._ws;            
            return a;
        }

        public override INameInfo GetName(string worksheet, string name)
        {
            ExcelNamedRange nameItem;
            ulong id;            
            ExcelWorksheet ws;
            if (string.IsNullOrEmpty(worksheet))
            {
                if(_package._workbook.Names.ContainsKey(name))
                {
                    nameItem = _package._workbook.Names[name];
                }
                else
                {
                    return null;
                }
                ws = null;
            }
            else
            {
                ws = _package._workbook.Worksheets[worksheet];
                if (ws !=null && ws.Names.ContainsKey(name))
                {
                    nameItem = ws.Names[name];
                }
                else if (_package._workbook.Names.ContainsKey(name))
                {
                    nameItem = _package._workbook.Names[name];
                }
                else
                {
                    var tbl = ws.Tables[name];
                    if (tbl != null)
                    {
                        nameItem = new ExcelNamedRange(name, ws, ws, tbl.DataRange.Address, -1);
                    }
                    else
                    {
                        var wsName = _package.Workbook.Worksheets[name];
                        if (wsName == null)
                        {
                            return null;
                        }
                        nameItem = new ExcelNamedRange(name, ws, wsName, "A:XFD", -1);
                    }
                }
            }
            id = ExcelAddressBase.GetCellID(nameItem.LocalSheetId, nameItem.Index, 0);

            if (_names.ContainsKey(id))
            {
                return _names[id];
            }
            else
            {
                var ni = new NameInfo()
                {
                    Id = id,
                    Name = name,
                    Worksheet = string.IsNullOrEmpty(worksheet) ? (nameItem.Worksheet==null ? nameItem._ws : nameItem.Worksheet.Name) : worksheet, 
                    Formula = nameItem.Formula
                };
                if (nameItem._fromRow > 0)
                {
                    ni.Value = new RangeInfo(nameItem.Worksheet ?? ws, nameItem._fromRow, nameItem._fromCol, nameItem._toRow, nameItem._toCol);
                }
                else
                {
                    ni.Value = nameItem.Value;
                }
                _names.Add(id, ni);
                return ni;
            }
        }
        public override IEnumerable<object> GetRangeValues(string address)
        {
            SetCurrentWorksheet(ExcelAddressInfo.Parse(address));
            var addr = new ExcelAddress(address);
            var wsName = string.IsNullOrEmpty(addr.WorkSheetName) ? _currentWorksheet.Name : addr.WorkSheetName;
            var ws = _package.Workbook.Worksheets[wsName];
            return (IEnumerable<object>)(new CellStoreEnumerator<ExcelValue>(ws._values, addr._fromRow, addr._fromCol, addr._toRow, addr._toCol));
        }


        public object GetValue(int row, int column)
        {
            return _currentWorksheet.GetValueInner(row, column);
        }

        public bool IsMerged(int row, int column)
        {
            //return _currentWorksheet._flags.GetFlagValue(row, column, CellFlags.Merged);
            return _currentWorksheet.MergedCells[row, column] != null;
        }

        public bool IsHidden(int row, int column)
        {
            return _currentWorksheet.Column(column).Hidden || _currentWorksheet.Column(column).Width == 0 ||
                   _currentWorksheet.Row(row).Hidden || _currentWorksheet.Row(column).Height == 0;
        }

        public override object GetCellValue(string sheetName, int row, int col)
        {
            SetCurrentWorksheet(sheetName);
            return _currentWorksheet.GetValueInner(row, col);
        }

        public override ExcelCellAddress GetDimensionEnd(string worksheet)
        {
            ExcelCellAddress address = null;
            try
            {
                address = _package.Workbook.Worksheets[worksheet].Dimension.End;
            }
            catch{}
            
            return address;
        }

        private void SetCurrentWorksheet(ExcelAddressInfo addressInfo)
        {
            if (addressInfo.WorksheetIsSpecified)
            {
                _currentWorksheet = _package.Workbook.Worksheets[addressInfo.Worksheet];
            }
            else if (_currentWorksheet == null)
            {
                _currentWorksheet = _package.Workbook.Worksheets.First();
            }
        }

        private void SetCurrentWorksheet(string worksheetName)
        {
            if (!string.IsNullOrEmpty(worksheetName))
            {
                _currentWorksheet = _package.Workbook.Worksheets[worksheetName];    
            }
            else
            {
                _currentWorksheet = _package.Workbook.Worksheets.First(); 
            }
            
        }

        //public override void SetCellValue(string address, object value)
        //{
        //    var addressInfo = ExcelAddressInfo.Parse(address);
        //    var ra = _rangeAddressFactory.Create(address);
        //    SetCurrentWorksheet(addressInfo);
        //    //var valueInfo = (ICalcEngineValueInfo)_currentWorksheet;
        //    //valueInfo.SetFormulaValue(ra.FromRow + 1, ra.FromCol + 1, value);
        //    _currentWorksheet.Cells[ra.FromRow + 1, ra.FromCol + 1].Value = value;
        //}

        public override void Dispose()
        {
            _package.Dispose();
        }

        public override int ExcelMaxColumns
        {
            get { return ExcelPackage.MaxColumns; }
        }

        public override int ExcelMaxRows
        {
            get { return ExcelPackage.MaxRows; }
        }

        public override string GetRangeFormula(string worksheetName, int row, int column)
        {
            SetCurrentWorksheet(worksheetName);
            return _currentWorksheet.GetFormula(row, column);
        }

        public override object GetRangeValue(string worksheetName, int row, int column)
        {
            SetCurrentWorksheet(worksheetName);
            return _currentWorksheet.GetValue(row, column);
        }
        public override string GetFormat(object value, string format)
        {
            var styles = _package.Workbook.Styles;
            ExcelNumberFormatXml.ExcelFormatTranslator ft=null;
            foreach(var f in styles.NumberFormats)
            {
                if(f.Format==format)
                {
                    ft=f.FormatTranslator;
                    break;
                }
            }
            if(ft==null)
            {
                ft=new ExcelNumberFormatXml.ExcelFormatTranslator(format, -1);
            }
            return ValueToTextHandler.FormatValue(value, ft,format, ft.NetFormat);
        }
        public override List<LexicalAnalysis.Token> GetRangeFormulaTokens(string worksheetName, int row, int column)
        {
            return _package.Workbook.Worksheets[worksheetName]._formulaTokens.GetValue(row, column);
        }

        public override bool IsRowHidden(string worksheetName, int row)
        {
            var b = _package.Workbook.Worksheets[worksheetName].Row(row).Height == 0 || 
                    _package.Workbook.Worksheets[worksheetName].Row(row).Hidden;

            return b;
        }

        public override void Reset()
        {
            _names = new Dictionary<ulong, INameInfo>(); //Reset name cache.            
        }

        //public override void SetToTableAddress(ExcelAddress address)
        //{
        //    address.SetRCFromTable(_package, address);
        //}
    }
}
    
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
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Table;
using System;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.ExternalReferences;
using OfficeOpenXml.FormulaParsing.Ranges;

namespace OfficeOpenXml.FormulaParsing
{
    /// <summary>
    /// EPPlus implementation of the ExcelDataProvider abstract class.
    /// </summary>
    internal class EpplusExcelDataProvider : ExcelDataProvider
    {

        private readonly ExcelPackage _package;
        private readonly ParsingContext _context;
        private ExcelWorksheet _currentWorksheet;
        private RangeAddressFactory _rangeAddressFactory;
        private Dictionary<ulong, INameInfo> _names=new Dictionary<ulong,INameInfo>();

        public EpplusExcelDataProvider(ExcelPackage package)
            : this(package, ParsingContext.Create(package))
        {

        }

        public EpplusExcelDataProvider(ExcelPackage package, ParsingContext ctx)
        {
            if (package == null) throw new ArgumentNullException(nameof(package));
            _package = package;
            _context = ctx;
            _rangeAddressFactory = new RangeAddressFactory(this, ctx);
        }

        protected ParsingContext ParsingContext => _context;

        public override IEnumerable<string> GetWorksheets()
        {
            return _package.Workbook.Worksheets.Select(x => x.Name);
        }

        public override ExcelNamedRangeCollection GetWorksheetNames(int wsIx)
        {
            var ws=_package.Workbook.Worksheets[wsIx];
            if (ws != null)
            {
                return ws.Names;
            }
            else
            {
                return null;
            }
        }
        public override ExcelNamedRangeCollection GetWorksheetNames(string worksheetName)
        {
            var ws = _package.Workbook.Worksheets[worksheetName];
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

        public override ExcelTable GetExcelTable(string name)
        {
            foreach (var ws in _package.Workbook.Worksheets)
            {
                if (ws is ExcelChartsheet) continue;
                if (ws.Tables._tableNames.ContainsKey(name))
                {
                    return ws.Tables[name];
                }
            }
            return null;
        }

        public override ExcelNamedRangeCollection GetWorkbookNameValues()
        {
            return _package.Workbook.Names;
        }

        internal override IRangeInfo GetRange(FormulaRangeAddress range)
        {
            if(range.ExternalReferenceIx > 0)
            {                 
                return new EpplusExcelExternalRangeInfo(range.ExternalReferenceIx, range.WorksheetIx, range.FromRow, range.FromCol, range.ToRow, range.ToCol, ParsingContext);
            }
            else
            {
                return new RangeInfo(range, ParsingContext);
            }
        }

        public override IRangeInfo GetRange(string worksheet, int fromRow, int fromCol, int toRow, int toCol)
        {
            SetCurrentWorksheet(worksheet);
            var wsName = string.IsNullOrEmpty(worksheet) ? _currentWorksheet.Name : worksheet;
            var ws = _package.Workbook.Worksheets[wsName];
            if (ws == null)
            {
                throw new ExcelErrorValueException(eErrorType.Ref);
            }
            else
            {
                return new RangeInfo(ws, fromRow, fromCol, toRow, toCol, ParsingContext);
            }
        }
        public override IRangeInfo GetRange(string worksheet, int row, int column, string address)
        {
            SetCurrentWorksheet(worksheet);
            var addr = new ExcelAddressBase(address, _package.Workbook, worksheet);
            if (addr.Table != null && string.IsNullOrEmpty(addr._wb))
            {
                addr = ConvertToA1C1(_package, addr, new ExcelAddressBase(row, column, row, column));
            }
            return GetRangeInternal(addr);
        }
        public override IRangeInfo GetRange(int wsIx, int row, int column)
        {
            if (wsIx < -1) wsIx = ParsingContext.CurrentCell.WorksheetIx;
            return new RangeInfo(new FormulaRangeAddress(_context) { WorksheetIx = wsIx, FromRow = row, FromCol = column, ToRow=row, ToCol=column }, _context);
        }
        public override IRangeInfo GetRange(string worksheet, string address)
        {
            SetCurrentWorksheet(worksheet);
            var addr = new ExcelAddressBase(address, _package.Workbook, worksheet);
            if (addr.Table != null)
            {
                addr = ConvertToA1C1(_package, addr, addr);
            }
            return GetRangeInternal(addr);
        }

        private IRangeInfo GetRangeInternal(ExcelAddressBase addr)
        {
            if (addr.IsExternal)
            {
                return GetExternalRangeInfo(addr, addr.WorkSheetName, _package.Workbook);
            }
            else
            {
                var wsName = string.IsNullOrEmpty(addr.WorkSheetName) ? _currentWorksheet.Name : addr.WorkSheetName;
                var ws = _package.Workbook.Worksheets[wsName];
                if (ws == null)
                {
                    throw new ExcelErrorValueException(eErrorType.Ref);
                }

                return new RangeInfo(ws, addr);
            }
        }

        private IRangeInfo GetExternalRangeInfo(ExcelAddressBase addr, string wsName, ExcelWorkbook wb)
        {
            ExcelExternalWorkbook externalWb;
            var ix = wb.ExternalLinks.GetExternalLink(addr._wb);
            if (ix >= 0)
            {
                externalWb = wb.ExternalLinks[ix].As.ExternalWorkbook;
            }
            else
            {
                throw new ExcelErrorValueException(eErrorType.Ref);
            }
            if (externalWb?.Package == null)
            {
                if(addr.Table!=null)
                {
                    throw new ExcelErrorValueException(eErrorType.Ref);
                }
                var ws = externalWb.Package.Workbook.Worksheets[wsName];
                return new EpplusExcelExternalRangeInfo(ix, ws.SheetId, addr._fromRow, addr._fromCol, addr._toRow, addr._toCol, ParsingContext);
            }
            else
            {
                addr = addr.ToInternalAddress();
                ExcelWorksheet ws;
                if (addr.Table == null)
                {
                    ws = externalWb.Package.Workbook.Worksheets[wsName];
                }
                else
                {
                    addr = ConvertToA1C1(externalWb.Package, addr, addr);
                    ws = externalWb.Package.Workbook.Worksheets[addr.WorkSheetName];
                }

                return new RangeInfo(ws, addr);
            }
        }
        private static ExcelAddress ConvertToA1C1(ExcelPackage package, ExcelAddressBase addr, ExcelAddressBase refAddress)
        {
            //Convert the Table-style Address to an A1C1 address
            addr.SetRCFromTable(package, refAddress);
            var a = new ExcelAddress(addr._fromRow, addr._fromCol, addr._toRow, addr._toCol);
            a._ws = addr._ws;            
            return a;
        }

        //public override INameInfo GetName(string worksheet, string name)
        //{
        //    if(ExcelCellBase.IsExternalAddress(name))
        //    {
        //        return GetExternalName(name, ParsingContext);
        //    }
        //    else
        //    {
        //        var ws = _package.Workbook.Worksheets[worksheet];
        //        var wsIx = ws == null ? -1 : ws.PositionId;
        //        return GetLocalName(_package, (short)wsIx, name, ParsingContext);
        //    }
        //}
        private INameInfo GetExternalName(string name, ParsingContext ctx)
        {
            var extRef = ExcelCellBase.GetWorkbookFromAddress(name);
            var ix = _package.Workbook.ExternalLinks.GetExternalLink(extRef);
            if (ix >= 0)
            {
                var externalWorkbook = _package.Workbook.ExternalLinks[ix].As.ExternalWorkbook;
                if(externalWorkbook!=null)
                {
                    if (externalWorkbook.Package == null)
                    {
                        return GetNameFromCache(externalWorkbook, name, ctx);
                    }
                    else
                    {
                        name = name.Substring(name.IndexOf("]") + 1);
                        if (name.StartsWith("!")) name = name.Substring(1);
                        return GetLocalName(externalWorkbook.Package, -1, name, ctx);
                    }
                }
            }
            return null;
        }
        private INameInfo GetExternalName(int extIx, int wsIx, string name, ParsingContext ctx)
        {
            extIx -= 1;
            if (extIx >= 0 && extIx < _package.Workbook.ExternalLinks.Count)
            {
                var externalWorkbook = _package.Workbook.ExternalLinks[extIx].As.ExternalWorkbook;
                if (externalWorkbook != null)
                {
                    if (externalWorkbook.Package == null)
                    {
                        return GetNameFromCache(externalWorkbook, wsIx, name, ctx);
                    }
                    else
                    {
                        return GetLocalName(externalWorkbook.Package, wsIx, name, ctx);
                    }
                }
                return new NameInfo()
                {
                    Name = name,
                    Value = ExcelErrorValue.Create(eErrorType.Name)
                };
            }
            return null;
        }

        private INameInfo GetLocalName(ExcelPackage package, int wsIx, string name, ParsingContext ctx)
        {
            ExcelNamedRange extName=null;
            if(wsIx==int.MinValue)
            {
                extName = package.Workbook.Names[name];
            }
            else if(wsIx != -1)
            {
                var ws = package.Workbook.Worksheets[wsIx];
                extName = ws.Names[name];
            }

            if (extName == null)
            {
                return new NameInfo()
                {
                    Name = name,
                    wsIx = -1,
                    Value = ExcelErrorValue.Create(eErrorType.Name),                    
                };
            }
            else
            {
                var ni = new NameInfo()
                {
                    Name = name,
                    wsIx = (short)(extName.Worksheet == null ? wsIx : (short)extName.Worksheet.PositionId),
                    Formula = extName.NameFormula
                };
                if (extName._fromRow > 0)
                {
                    ni.Value = new RangeInfo(extName.Worksheet ?? package.Workbook.Worksheets[extName.WorkSheetName], extName._fromRow, extName._fromCol, extName._toRow, extName._toCol, ctx, extName.ExternalReferenceIndex + 1);
                }
                else
                {
                    ni.Value = extName.Value;
                }
                return ni;
            }
        }
        private static INameInfo GetNameFromCache(ExcelExternalWorkbook externalWorkbook, int wsIx, string name, ParsingContext ctx)
        {
            ExcelExternalDefinedName nameItem=null;

            //int ix=-1;
            if (wsIx==int.MinValue)
            {
                nameItem = externalWorkbook.CachedNames[name];
            }
            else if(wsIx!=-1)
            {
                nameItem = externalWorkbook.CachedWorksheets[wsIx].CachedNames[name];
            }

            object value;
            if (!string.IsNullOrEmpty(nameItem?.RefersTo))
            {
                var nameAddress = nameItem.RefersTo.TrimStart('=');
                ExcelAddressBase address = new ExcelAddressBase(nameAddress);
                if (address.Address == "#REF!")
                {
                    value = ExcelErrorValue.Create(eErrorType.Ref);
                }
                else
                {
                    value = new EpplusExcelExternalRangeInfo(externalWorkbook.Index, externalWorkbook.CachedWorksheets.GetIndexByName(address.WorkSheetName), address._fromRow, address._fromCol, address._toRow, address._toCol, ctx);
                }
            }
            else
            {
                value = ExcelErrorValue.Create(eErrorType.Name);
            }
            return new NameInfo()
            {
                Name = name,
                Value = value
            };
        }
        private static INameInfo GetNameFromCache(ExcelExternalWorkbook externalWorkbook, string name, ParsingContext ctx)
        {
            ExcelExternalDefinedName nameItem;

            int ix = -1;
            var sheetName = ExcelAddressBase.GetWorksheetPart(name, "", ref ix);
            if (string.IsNullOrEmpty(sheetName))
            {
                if (ix > 0) name = name.Substring(ix);
                nameItem = externalWorkbook.CachedNames[name];
            }
            else
            {
                if (ix >= 0) name = name.Substring(ix);
                nameItem = externalWorkbook.CachedWorksheets[sheetName].CachedNames[name];
            }

            object value;
            if (!string.IsNullOrEmpty(nameItem.RefersTo))
            {
                var nameAddress = nameItem.RefersTo.TrimStart('=');
                ExcelAddressBase address = new ExcelAddressBase(nameAddress);
                if (address.Address == "#REF!")
                {
                    value = ExcelErrorValue.Create(eErrorType.Ref);
                }
                else
                {
                    value = new EpplusExcelExternalRangeInfo(ix, nameItem.SheetId, -1, -1, -1, -1, ctx);
                }
            }
            else
            {
                value = ExcelErrorValue.Create(eErrorType.Name);
            }
            return new NameInfo()
            {
                Name = name,
                Value = value
            };
        }
        //private INameInfo GetLocalName(ExcelPackage package, short wsIx, string name, ParsingContext ctx)
        //{
        //    ExcelNamedRange nameItem;
        //    ulong id;
        //    ExcelWorksheet ws;
        //    var ix = name.IndexOf('!');
        //    if(ix>0)
        //    {
        //        var wsName=ExcelAddressBase.GetWorksheetPart(name, "", ref ix);
        //        if(!string.IsNullOrEmpty(wsName))
        //        {
        //            name = name.Substring(ix);
        //            wsIx = (short)_package.Workbook.Worksheets[wsName].PositionId;
        //        }
        //    }
        //    if (wsIx<0)
        //    {
        //        if (package._workbook.Names.ContainsKey(name))
        //        {
        //            nameItem = package._workbook.Names[name];
        //        }
        //        else
        //        {
        //            return null;
        //        }
        //        ws = null;
        //    }
        //    else
        //    {
        //        ws = package._workbook.Worksheets[wsIx];
        //        if (ws != null && ws.Names.ContainsKey(name))
        //        {
        //            nameItem = ws.Names[name];
        //        }
        //        else if (package._workbook.Names.ContainsKey(name))
        //        {
        //            nameItem = package._workbook.Names[name];
        //        }
        //        else
        //        {
        //            var tbl = ws.Tables[name];
        //            if (tbl != null)
        //            {
        //                nameItem = new ExcelNamedRange(name, ws, ws, tbl.DataRange.Address, -1);
        //            }
        //            else
        //            {
        //                var wsName = package.Workbook.Worksheets[name];
        //                if (wsName == null)
        //                {
        //                    return null;
        //                }
        //                nameItem = new ExcelNamedRange(name, ws, wsName, "A:XFD", -1);
        //            }
        //        }
        //    }
        //    id = ExcelAddressBase.GetCellId(nameItem.LocalSheetId, nameItem.Index, 0);

        //    if (_names.ContainsKey(id))
        //    {
        //        return _names[id];
        //    }
        //    else
        //    {
        //        var ni = new NameInfo()
        //        {
        //            Id = id,
        //            Name = name,
        //            wsIx = (nameItem.Worksheet == null ? wsIx : (short)nameItem.Worksheet.PositionId),
        //            Formula = nameItem.Formula
        //        };
        //        if (nameItem._fromRow > 0)
        //        {
        //            ni.Value = new RangeInfo(nameItem.Worksheet ?? ws, nameItem._fromRow, nameItem._fromCol, nameItem._toRow, nameItem._toCol, ctx);
        //        }
        //        else
        //        {
        //            ni.Value = nameItem.Value;
        //        }
        //        _names.Add(id, ni);
        //        return ni;
        //    }
        //}

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
        public override object GetCellValue(int wsIx, int row, int col)
        {
            _currentWorksheet = _package.Workbook.Worksheets[wsIx];
            return _currentWorksheet.GetValueInner(row, col);
        }
        //public override ulong GetCellId(string sheetName, int row, int col)
        //{
        //    if (string.IsNullOrEmpty(sheetName)) return 0;
        //    var worksheet = _package.Workbook.Worksheets[sheetName];
        //    var wsIx = worksheet != null ? worksheet.IndexInList : 0;
        //    return ExcelCellBase.GetCellId(wsIx, row, col);
        //}

        public override ExcelCellAddress GetDimensionEnd(int wsIx)
        {
            if (wsIx < 0 || wsIx >= _package.Workbook.Worksheets.Count) return null;
            ExcelCellAddress address = null;
            try
            {
                address = _package.Workbook.Worksheets[wsIx].Dimension?.End;
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
        private void SetCurrentWorksheet(int wsIx)
        {
            if (wsIx < 0)
            {
                _currentWorksheet = _package.Workbook.Worksheets[wsIx];
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
            return ValueToTextHandler.FormatValue(value,false, ft, null);
        }
        public override IList<Token> GetRangeFormulaTokens(string worksheetName, int row, int column)
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

        public override bool IsExternalName(string name)
        {
            if (name[0] != '[') return false;
            var ixEnd = name.IndexOf("]");
            if(ixEnd>0)
            {
                var ix = name.Substring(1,ixEnd-1);
                var extRef=_package.Workbook.ExternalLinks.GetExternalLink(ix);
                if (extRef < 0) return false;
                var extBook = _package.Workbook.ExternalLinks[extRef].As.ExternalWorkbook;
                if(extBook==null) return false;
                var address = name.Substring(ixEnd+1);
                if (address.StartsWith("!"))
                {
                    return extBook.CachedNames.ContainsKey(address.Substring(1));
                }
                else
                {
                    int addressStart = -1;
                    var sheetName = ExcelAddressBase.GetWorksheetPart(address, "", ref addressStart);
                    if (extBook.CachedWorksheets.ContainsKey(sheetName) && addressStart>0)
                    {
                        return extBook.CachedWorksheets[sheetName].CachedNames.ContainsKey(address.Substring(addressStart));
                    }
                }
            }
            return false;
        }

        public override INameInfo GetName(int externalRef, int wsIx, string name)
        {
            if(externalRef>0)
            {
                return GetExternalName(externalRef, wsIx, name, _context);
            }
            else
            {
                var wb = _package.Workbook;                
                if (wsIx == -1)
                {
                    return new NameInfo()
                    {
                        Name = name,
                        wsIx = -1,
                        Value = ExcelErrorValue.Create(eErrorType.Name)
                    };
                }
                var workSheetIx = wsIx < 0 ? ParsingContext.CurrentCell.WorksheetIx : wsIx;
                ExcelNamedRange nameItem = null;

                if (workSheetIx >= 0 && workSheetIx < wb.Worksheets.Count)
                {
                    var ws = wb.GetWorksheetByIndexInList(workSheetIx);
                    if (ws.Names.ContainsKey(name))
                    {
                        nameItem = ws.Names[name];
                    }
                }

                if (wsIx < 0 && nameItem == null && wb.Names.ContainsKey(name))
                {
                    nameItem = wb.Names[name];
                }

                if (nameItem == null) return null;

                return GetName(nameItem);
            }
        }

        /// <summary>
        /// Gets a IName
        /// </summary>
        /// <param name="nameItem"></param>
        /// <returns></returns>
        public override INameInfo GetName(ExcelNamedRange nameItem)
        {
            var id = ExcelCellBase.GetCellId(nameItem.LocalSheetId, nameItem.Index, 0);
            var ni = new NameInfo()
            {
                Id = id,
                Name = nameItem.Name,
                wsIx = (nameItem.Worksheet == null ? int.MinValue : nameItem.Worksheet.IndexInList),
                Formula = nameItem.NameFormula
            };
            if (nameItem._fromRow > 0)
            {
                ni.Value = new RangeInfo(nameItem.Worksheet ?? ParsingContext.CurrentWorksheet, nameItem._fromRow, nameItem._fromCol, nameItem._toRow, nameItem._toCol, ParsingContext);
            }
            else
            {
                ni.Value = nameItem.Value;
            }

            return ni;
        }

        public override ulong GetCellId(int wsIx, int row, int col)
        {
            return ExcelCellBase.GetCellId(wsIx, row, col); 
        }
    }
}
    
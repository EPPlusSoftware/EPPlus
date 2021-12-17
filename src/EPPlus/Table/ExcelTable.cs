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
using System.Xml;
using System.Text.RegularExpressions;
using OfficeOpenXml.Utils;
using System.Security;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.Filter;
using OfficeOpenXml.Core.Worksheet;
using System.Data;
using OfficeOpenXml.Export.ToDataTable;
using System.IO;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Export.HtmlExport;
using System.Globalization;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Sorting;
#if !NET35 && !NET40
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Table
{
    /// <summary>
    /// An Excel Table
    /// </summary>
    public class ExcelTable : ExcelTableDxfBase, IEqualityComparer<ExcelTable>
    {
        internal ExcelTable(Packaging.ZipPackageRelationship rel, ExcelWorksheet sheet) : 
            base(sheet.NameSpaceManager)
        {
            WorkSheet = sheet;
            TableUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
            RelationshipID = rel.Id;
            var pck = sheet._package.ZipPackage;
            Part=pck.GetPart(TableUri);

            TableXml = new XmlDocument();
            LoadXmlSafe(TableXml, Part.GetStream());
            Init();
            Address = new ExcelAddressBase(GetXmlNodeString("@ref"));
            _tableStyle = GetTableStyle(StyleName);            
        }
        internal ExcelTable(ExcelWorksheet sheet, ExcelAddressBase address, string name, int tblId) : 
            base(sheet.NameSpaceManager)
	    {
            WorkSheet = sheet;
            _address = address;

            TableXml = new XmlDocument();
            LoadXmlSafe(TableXml, GetStartXml(name, tblId), Encoding.UTF8); 

            Init();

            //If the table is just one row we cannot have a header.
            if (address._fromRow == address._toRow)
            {
                ShowHeader = false;
            }
            if(AutoFilterAddress!=null)
            {
                SetAutoFilter();
            }
        }

        private void Init()
        {
            TopNode = TableXml.DocumentElement;
            SchemaNodeOrder = new string[] { "autoFilter", "sortState", "tableColumns", "tableStyleInfo" };
            InitDxf(WorkSheet.Workbook.Styles, this, null);
            TableBorderStyle = new ExcelDxfBorderBase(WorkSheet.Workbook.Styles, null);
            HeaderRowBorderStyle = new ExcelDxfBorderBase(WorkSheet.Workbook.Styles, null);
            _tableSorter = new TableSorter(this);
            HtmlExporter = new TableExporter(this);
        }

        private string GetStartXml(string name, int tblId)
        {
            name = ConvertUtil.ExcelEscapeString(name);
            string xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?>";
            xml += string.Format("<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"{0}\" name=\"{1}\" displayName=\"{2}\" ref=\"{3}\" headerRowCount=\"1\">",
            tblId,
            name,
            ExcelAddressUtil.GetValidName(name),
            Address.Address);
            xml += string.Format("<autoFilter ref=\"{0}\" />", Address.Address);

            int cols=Address._toCol-Address._fromCol+1;
            xml += string.Format("<tableColumns count=\"{0}\">",cols);
            var names = new HashSet<string>();            
            for(int i=1;i<=cols;i++)
            {
                var cell = WorkSheet.Cells[Address._fromRow, Address._fromCol+i-1];
                string colName= SecurityElement.Escape(cell.Value?.ToString());
                if (cell.Value == null || names.Contains(colName))
                {
                    //Get an unique name
                    int a=i;
                    do
                    {
                        colName = string.Format("Column{0}", a++);
                    }
                    while (names.Contains(colName));
                }
                names.Add(colName);
                xml += string.Format("<tableColumn id=\"{0}\" name=\"{1}\" />", i,colName);
            }
            xml += "</tableColumns>";
            xml += "<tableStyleInfo name=\"TableStyleMedium9\" showFirstColumn=\"0\" showLastColumn=\"0\" showRowStripes=\"1\" showColumnStripes=\"0\" /> ";
            xml += "</table>";

            return xml;
        }
        internal static string CleanDisplayName(string name) 
        {
            return Regex.Replace(name, @"[^\w\.-_]", "_");
        }
        internal Packaging.ZipPackagePart Part
        {
            get;
            set;
        }
        /// <summary>
        /// Provides access to the XML data representing the table in the package.
        /// </summary>
        public XmlDocument TableXml
        {
            get;
            set;
        }
        /// <summary>
        /// The package internal URI to the Table Xml Document.
        /// </summary>
        public Uri TableUri
        {
            get;
            internal set;
        }
        internal string RelationshipID
        {
            get;
            set;
        }
        const string ID_PATH = "@id";
        internal int Id 
        {
            get
            {
                return GetXmlNodeInt(ID_PATH);
            }
            set
            {
                SetXmlNodeString(ID_PATH, value.ToString());
            }
        }
        const string NAME_PATH = "@name";
        const string DISPLAY_NAME_PATH = "@displayName";
        /// <summary>
        /// The name of the table object in Excel
        /// </summary>
        public string Name
        {
            get
            {
                return GetXmlNodeString(NAME_PATH);
            }
            set 
            {
                if(Name.Equals(value, StringComparison.CurrentCultureIgnoreCase)==false && WorkSheet.Workbook.ExistsTableName(value))
                {
                    throw (new ArgumentException("Tablename is not unique"));
                }
                string prevName = Name;
                if (WorkSheet.Tables._tableNames.ContainsKey(prevName))
                {
                    int ix=WorkSheet.Tables._tableNames[prevName];
                    WorkSheet.Tables._tableNames.Remove(prevName);
                    WorkSheet.Tables._tableNames.Add(value,ix);
                }
                var ta = new TableAdjustFormula(this);
                ta.AdjustFormulas(prevName, value);
                SetXmlNodeString(NAME_PATH, value);
                SetXmlNodeString(DISPLAY_NAME_PATH, ExcelAddressUtil.GetValidName(value));
            }
        }
        
        internal void DeleteMe()
        {
            if (RelationshipID != null)
            {
                WorkSheet.DeleteNode($"d:tableParts/d:tablePart[@r:id='{RelationshipID}']");
            }
            if (TableUri != null && WorkSheet._package.ZipPackage.PartExists(TableUri))
            {
                WorkSheet._package.ZipPackage.DeletePart(TableUri);
            }
        }

        /// <summary>
        /// The worksheet of the table
        /// </summary>
        public ExcelWorksheet WorkSheet
        {
            get;
            set;
        }

        private ExcelAddressBase _address = null;
        /// <summary>
        /// The address of the table
        /// </summary>
        public ExcelAddressBase Address
        {
            get
            {
                return _address;
            }
            internal set
            {
                _address = value;
                if(value!=null)
                {
                    SetXmlNodeString("@ref", value.Address);
                    WriteAutoFilter(ShowTotal);
                }
            }
        }
        /// <summary>
        /// The table range
        /// </summary>
        public ExcelRangeBase Range
        {
            get
            {
                return WorkSheet.Cells[_address._fromRow, _address._fromCol, _address._toRow, _address._toCol];
            }
        }
        internal ExcelRangeBase DataRange
        {
            get
            {
                int fromRow = ShowHeader ? _address._fromRow + 1 : _address._fromRow;
                int toRow = ShowTotal ? _address._toRow - 1: _address._toRow;

                return WorkSheet.Cells[fromRow, _address._fromCol, toRow, _address._toCol];
            }
        }

        #region Export table data

        /// <summary>
        /// Converts the table range to CSV format
        /// </summary>
        /// <returns></returns>
        /// <seealso cref="ExcelRangeBase.ToText()"/>
        public string ToText()
        {
            return Range.ToText();
        }

        public TableExporter HtmlExporter { get; private set; }

        /// <summary>
        /// Converts the table range to CSV format
        /// </summary>
        /// <param name="format">Parameters/options for conversion to text</param>
        /// <returns></returns>
        /// <seealso cref="ExcelRangeBase.ToText(ExcelOutputTextFormat)"/>
        public string ToText(ExcelOutputTextFormat format)
        {
            return Range.ToText(format);
        }

#if !NET35 && !NET40
        /// <summary>
        /// Converts the table range to CSV format
        /// </summary>
        /// <returns></returns>
        /// <seealso cref="ExcelRangeBase.ToTextAsync()"/>
        public Task<string> ToTextAsync()
        {
            return Range.ToTextAsync();
        }

        /// <summary>
        /// Converts the table range to CSV format
        /// </summary>
        /// <returns></returns>
        /// <seealso cref="ExcelRangeBase.ToText(ExcelOutputTextFormat)"/>
        public Task<string> ToTextAsync(ExcelOutputTextFormat format)
        {
            return Range.ToTextAsync(format);
        }
#endif

        /// <summary>
        /// Exports the table to a file
        /// </summary>
        /// <param name="file">The export file</param>
        /// <param name="format">Export options</param>
        /// <seealso cref="ExcelRangeBase.SaveToText(FileInfo, ExcelOutputTextFormat)"></seealso>
        public void SaveToText(FileInfo file, ExcelOutputTextFormat format)
        {
            Range.SaveToText(file, format);
        }

        /// <summary>
        /// Exports the table to a <see cref="Stream"/>
        /// </summary>
        /// <param name="stream">Data will be exported to this stream</param>
        /// <param name="format">Export options</param>
        /// <seealso cref="ExcelRangeBase.SaveToText(Stream, ExcelOutputTextFormat)"></seealso>
        public void SaveToText(Stream stream, ExcelOutputTextFormat format)
        {
            Range.SaveToText(stream, format);
        }
#if !NET35 && !NET40
        /// <summary>
        /// Exports the table to a <see cref="Stream"/>
        /// </summary>
        /// <param name="stream">Data will be exported to this stream</param>
        /// <param name="format">Export options</param>
        /// <seealso cref="ExcelRangeBase.SaveToText(Stream, ExcelOutputTextFormat)"></seealso>
        public async Task SaveToTextAsync(Stream stream, ExcelOutputTextFormat format)
        {
            await Range.SaveToTextAsync(stream, format);
        }

        /// <summary>
        /// Exports the table to a file
        /// </summary>
        /// <param name="file">Data will be exported to this stream</param>
        /// <param name="format">Export options</param>
        /// <seealso cref="ExcelRangeBase.SaveToTextAsync(FileInfo, ExcelOutputTextFormat)"/>
        public async Task SaveToTextAsync(FileInfo file, ExcelOutputTextFormat format)
        {
            await Range.SaveToTextAsync(file, format);
        }

#endif

        /// <summary>
        /// Exports the table to a <see cref="System.Data.DataTable"/>
        /// </summary>
        /// <returns>A <see cref="System.Data.DataTable"/> containing the data in the table range</returns>
        /// <seealso cref="ExcelRangeBase.ToDataTable()"/>
        public DataTable ToDataTable()
        {
            return Range.ToDataTable();
        }

        /// <summary>
        /// Exports the table to a <see cref="System.Data.DataTable"/>
        /// </summary>
        /// <returns>A <see cref="System.Data.DataTable"/> containing the data in the table range</returns>
        /// <seealso cref="ExcelRangeBase.ToDataTable(ToDataTableOptions)"/>
        public DataTable ToDataTable(ToDataTableOptions options)
        {
            return Range.ToDataTable(options);
        }

        /// <summary>
        /// Exports the table to a <see cref="System.Data.DataTable"/>
        /// </summary>
        /// <returns>A <see cref="System.Data.DataTable"/> containing the data in the table range</returns>
        /// <seealso cref="ExcelRangeBase.ToDataTable(Action{ToDataTableOptions})"/>
        public DataTable ToDataTable(Action<ToDataTableOptions> configHandler)
        {
            return Range.ToDataTable(configHandler);
        }

        #endregion

        internal ExcelTableColumnCollection _cols = null;
        /// <summary>
        /// Collection of the columns in the table
        /// </summary>
        public ExcelTableColumnCollection Columns
        {
            get
            {
                if(_cols==null)
                {
                    _cols = new ExcelTableColumnCollection(this);
                }
                return _cols;
            }
        }
        TableStyles _tableStyle = TableStyles.Medium6;
        /// <summary>
        /// The table style. If this property is custom, the style from the StyleName propery is used.
        /// </summary>
        public TableStyles TableStyle
        {
            get
            {                
                return _tableStyle;
            }
            set
            {
                _tableStyle=value;
                if (value != TableStyles.Custom)
                {
                    SetXmlNodeString(STYLENAME_PATH, "TableStyle" + value.ToString());
                }
            }
        }
        const string HEADERROWCOUNT_PATH = "@headerRowCount";
        const string AUTOFILTER_PATH = "d:autoFilter";
        const string AUTOFILTER_ADDRESS_PATH = AUTOFILTER_PATH + "/@ref";
        /// <summary>
        /// If the header row is visible or not
        /// </summary>
        public bool ShowHeader
        {
            get
            {
                return GetXmlNodeInt(HEADERROWCOUNT_PATH)!=0;
            }
            set
            {
                if (Address._toRow - Address._fromRow < 0 && value ||
                    Address._toRow - Address._fromRow == 1 && value && ShowTotal)
                {
                    throw (new Exception("Cant set ShowHeader-property. Table has too few rows"));
                }

                if(value)
                {
                    DeleteNode(HEADERROWCOUNT_PATH);
                    WriteAutoFilter(ShowTotal);
                    for (int i = 0; i < Columns.Count; i++)
                    {
                        var v = WorkSheet.GetValue<string>(Address._fromRow, Address._fromCol + i);
                        if(string.IsNullOrEmpty(v))
                        {
                            WorkSheet.SetValue(Address._fromRow, Address._fromCol + i, _cols[i].Name);
                        }
                        else if (v != _cols[i].Name)
                        {
                            _cols[i].Name = v;
                        }
                    }
                    HeaderRowStyle.SetStyle();
                    foreach (var c in Columns)
                    {
                        c.HeaderRowStyle.SetStyle();
                    }
                }
                else
                {
                    SetXmlNodeString(HEADERROWCOUNT_PATH, "0");
                    DeleteAllNode(AUTOFILTER_ADDRESS_PATH);
                    DataStyle.SetStyle();
                }
            }
        }
        internal ExcelAddressBase AutoFilterAddress
        {
            get
            {
                string a=GetXmlNodeString(AUTOFILTER_ADDRESS_PATH);
                if (a == "")
                {
                    return null;
                }
                else
                {
                    return new ExcelAddressBase(a);
                }
            }
        }
        ExcelAutoFilter _autoFilter=null;
        /// <summary>
        /// Autofilter settings for the table
        /// </summary>
        public ExcelAutoFilter AutoFilter
        {
            get
            {
                if(ShowFilter)
                {
                    return _autoFilter;
                }
                else
                {
                    return null;
                }
            }
        }
        private void WriteAutoFilter(bool showTotal)
        {
            string autofilterAddress;
            if (ShowHeader)
            {
                if (showTotal)
                {
                    autofilterAddress = ExcelCellBase.GetAddress(Address._fromRow, Address._fromCol, Address._toRow - 1, Address._toCol);
                }
                else
                {
                    autofilterAddress = Address.Address;
                }
                SetXmlNodeString(AUTOFILTER_ADDRESS_PATH, autofilterAddress);
                SetAutoFilter();
            }
        }

        private void SetAutoFilter()
        {
            if (_autoFilter == null)
            {
                var node = TopNode.SelectSingleNode(AUTOFILTER_PATH, NameSpaceManager);
                _autoFilter = new ExcelAutoFilter(NameSpaceManager, node, this);
                _autoFilter.Address = AutoFilterAddress;
            }

        }

        /// <summary>
        /// If the header row has an autofilter
        /// </summary>
        public bool ShowFilter 
        { 
            get
            {
                return ShowHeader && AutoFilterAddress != null;
            }
            set
            {
                if (ShowHeader)
                {
                    if (value)
                    {
                        WriteAutoFilter(ShowTotal);
                    }
                    else 
                    {
                        DeleteAllNode(AUTOFILTER_PATH);
                        _autoFilter = null;
                    }
                }
                else if(value)
                {
                    throw(new InvalidOperationException("Filter can only be applied when ShowHeader is set to true"));
                }
            }
        }
        const string TOTALSROWCOUNT_PATH = "@totalsRowCount";
        const string TOTALSROWSHOWN_PATH = "@totalsRowShown";
        /// <summary>
        /// If the total row is visible or not
        /// </summary>
        public bool ShowTotal
        {
            get
            {
                return GetXmlNodeInt(TOTALSROWCOUNT_PATH) == 1;
            }
            set
            {
                if (value != ShowTotal)
                {
                    if (value)
                    {
                        Address=new ExcelAddress(WorkSheet.Name, ExcelAddressBase.GetAddress(Address.Start.Row, Address.Start.Column, Address.End.Row+1, Address.End.Column));
                    }
                    else
                    {
                        Address = new ExcelAddress(WorkSheet.Name, ExcelAddressBase.GetAddress(Address.Start.Row, Address.Start.Column, Address.End.Row - 1, Address.End.Column));
                    }
                    SetXmlNodeString("@ref", Address.Address);
                    if (value)
                    {
                        SetXmlNodeString(TOTALSROWCOUNT_PATH, "1");
                        TotalsRowStyle.SetStyle();
                        foreach (var c in Columns)
                        {
                            c.TotalsRowStyle.SetStyle();
                        }
                    }
                    else
                    {
                        DeleteNode(TOTALSROWCOUNT_PATH);
                        DataStyle.SetStyle();
                    }
                    WriteAutoFilter(value);
                }
            }
        }
        const string STYLENAME_PATH = "d:tableStyleInfo/@name";
        /// <summary>
        /// The style name for custum styles
        /// </summary>
        public string StyleName
        {
            get
            {
                return GetXmlNodeString(STYLENAME_PATH);
            }
            set
            {
                _tableStyle = GetTableStyle(value);
                if(_tableStyle==TableStyles.None)
                {
                    DeleteAllNode(STYLENAME_PATH);
                }
                else
                {
                    SetXmlNodeString(STYLENAME_PATH, value);
                }
            }
        }

        private TableStyles GetTableStyle(string value)
        {
            if (value.StartsWith("TableStyle"))
            {
                try
                {
                    return (TableStyles)Enum.Parse(typeof(TableStyles), value.Substring(10, value.Length - 10), true);
                }
                catch
                {
                    return TableStyles.Custom;
                }
            }
            else if (value == "None")
            {
                return TableStyles.None;
            }
            else
            {
                return TableStyles.Custom;
            }
        }

        const string SHOWFIRSTCOLUMN_PATH = "d:tableStyleInfo/@showFirstColumn";
        /// <summary>
        /// Display special formatting for the first row
        /// </summary>
        public bool ShowFirstColumn
        {
            get
            {
                return GetXmlNodeBool(SHOWFIRSTCOLUMN_PATH);
            }
            set
            {
                SetXmlNodeBool(SHOWFIRSTCOLUMN_PATH, value, false);
            }   
        }
        const string SHOWLASTCOLUMN_PATH = "d:tableStyleInfo/@showLastColumn";
        /// <summary>
        /// Display special formatting for the last row
        /// </summary>
        public bool ShowLastColumn
        {
            get
            {
                return GetXmlNodeBool(SHOWLASTCOLUMN_PATH);
            }
            set
            {
                SetXmlNodeBool(SHOWLASTCOLUMN_PATH, value, false);
            }
        }
        const string SHOWROWSTRIPES_PATH = "d:tableStyleInfo/@showRowStripes";
        /// <summary>
        /// Display banded rows
        /// </summary>
        public bool ShowRowStripes
        {
            get
            {
                return GetXmlNodeBool(SHOWROWSTRIPES_PATH);
            }
            set
            {
                SetXmlNodeBool(SHOWROWSTRIPES_PATH, value, false);
            }
        }
        const string SHOWCOLUMNSTRIPES_PATH = "d:tableStyleInfo/@showColumnStripes";
        /// <summary>
        /// Display banded columns
        /// </summary>
        public bool ShowColumnStripes
        {
            get
            {
                return GetXmlNodeBool(SHOWCOLUMNSTRIPES_PATH);
            }
            set
            {
                SetXmlNodeBool(SHOWCOLUMNSTRIPES_PATH, value, false);
            }
        }

        const string TOTALSROWCELLSTYLE_PATH = "@totalsRowCellStyle";
        /// <summary>
        /// Named style used for the total row
        /// </summary>
        public string TotalsRowCellStyle
        {
            get
            {
                return GetXmlNodeString(TOTALSROWCELLSTYLE_PATH);
            }
            set
            {
                if (WorkSheet.Workbook.Styles.NamedStyles.FindIndexById(value) < 0)
                {
                    throw (new Exception(string.Format("Named style {0} does not exist.", value)));
                }
                SetXmlNodeString(TopNode, TOTALSROWCELLSTYLE_PATH, value, true);

                if (ShowTotal)
                {
                    WorkSheet.Cells[Address._toRow, Address._fromCol, Address._toRow, Address._toCol].StyleName = value;
                }
            }
        }
        const string DATACELLSTYLE_PATH = "@dataCellStyle";
        /// <summary>
        /// Named style used for the data cells
        /// </summary>
        public string DataCellStyleName
        {
            get
            {
                return GetXmlNodeString(DATACELLSTYLE_PATH);
            }
            set
            {
                if (WorkSheet.Workbook.Styles.NamedStyles.FindIndexById(value) < 0)
                {
                    throw (new Exception(string.Format("Named style {0} does not exist.", value)));
                }
                SetXmlNodeString(TopNode, DATACELLSTYLE_PATH, value, true);

                int fromRow = Address._fromRow + (ShowHeader ? 1 : 0),
                    toRow = Address._toRow - (ShowTotal ? 1 : 0);

                if (fromRow < toRow)
                {
                    WorkSheet.Cells[fromRow, Address._fromCol, toRow, Address._toCol].StyleName = value;
                }
            }
        }
        const string HEADERROWCELLSTYLE_PATH = "@headerRowCellStyle";
        /// <summary>
        /// Named style used for the header row
        /// </summary>
        public string HeaderRowCellStyle
        {
            get
            {
                return GetXmlNodeString(HEADERROWCELLSTYLE_PATH);
            }
            set
            {
                if (WorkSheet.Workbook.Styles.NamedStyles.FindIndexById(value) < 0)
                {
                    throw (new Exception(string.Format("Named style {0} does not exist.", value)));
                }
                SetXmlNodeString(TopNode, HEADERROWCELLSTYLE_PATH, value, true);

                if (ShowHeader)
                {
                    WorkSheet.Cells[Address._fromRow, Address._fromCol, Address._fromRow, Address._toCol].StyleName = value;
                }

            }
        }
        /// <summary>
        /// Checkes if two tables are the same
        /// </summary>
        /// <param name="x">Table 1</param>
        /// <param name="y">Table 2</param>
        /// <returns></returns>
        public bool Equals(ExcelTable x, ExcelTable y)
        {
            return x.WorkSheet == y.WorkSheet && x.Id == y.Id && x.TableXml.OuterXml == y.TableXml.OuterXml;
        }
        /// <summary>
        /// Returns a hashcode generated from the TableXml
        /// </summary>
        /// <param name="obj">The table</param>
        /// <returns>The hashcode</returns>
        public int GetHashCode(ExcelTable obj)
        {
            return obj.TableXml.OuterXml.GetHashCode();
        }
        /// <summary>
        /// Adds new rows to the table. 
        /// </summary>
        /// <param name="rows">Number of rows to add to the table. Default is 1</param>
        /// <returns></returns>
        public ExcelRangeBase AddRow(int rows = 1)
        {
            return InsertRow(int.MaxValue, rows);
        }
        /// <summary>
        /// Inserts one or more rows before the specified position in the table.
        /// </summary>
        /// <param name="position">The position in the table where the row will be inserted. Default is in the end of the table. 0 will insert the row at the top. Any value larger than the number of rows in the table will insert a row at the bottom of the table.</param>
        /// <param name="rows">Number of rows to insert.</param>
        /// <param name="copyStyles">Copy styles from the row above. If inserting a row at position 0, the first row will be used as a template.</param>
        /// <returns>The inserted range</returns>
        public ExcelRangeBase InsertRow(int position, int rows=1, bool copyStyles=true)
        {
            if(position < 0)
            {
                throw new ArgumentException("position", "position can't be negative");
            }
            if (rows < 0)
            {
                throw new ArgumentException("position", "rows can't be negative");
            }
            var isFirstRow = position == 0;
            var subtact = ShowTotal ? 2 : 1;
            if (position>=ExcelPackage.MaxRows || position > _address._fromRow + position + rows - subtact)
            {
                position = _address.Rows - subtact;
            }
            if (_address._fromRow+position+rows>ExcelPackage.MaxRows)
            {
                throw new InvalidOperationException("Insert will exceed the maximum number of rows in the worksheet");
            }
            if(ShowHeader) position++;
            var address = ExcelCellBase.GetAddress(_address._fromRow + position, _address._fromCol, _address._fromRow + position + rows - 1, _address._toCol);
            var range = new ExcelRangeBase(WorkSheet, address);

            WorksheetRangeInsertHelper.Insert(range,eShiftTypeInsert.Down, false);

            ExtendCalculatedFormulas(range);

            if (copyStyles)
            {
                int copyFromRow = isFirstRow ? DataRange._fromRow + rows + 1 : _address._fromRow + position - 1;
                if (range._toRow > _address._toRow)
                {
                    Address = _address.AddRow(_address._toRow, rows);
                }
                CopyStylesFromRow(address, copyFromRow);    //Separate copy instead of using Insert paramter 3 as the first row should not copy the styles from the header row.
            }

            return range;
        }

        private void ExtendCalculatedFormulas(ExcelRangeBase range)
        {
            foreach(var c in Columns)
            {
                if(!string.IsNullOrEmpty(c.CalculatedColumnFormula))
                {
                    c.SetFormulaCells(range._fromRow, range._toRow, range._fromCol + c.Position);
                }
            }
        }

        private void CopyStylesFromRow(string address, int copyRow)
        {
            var range = WorkSheet.Cells[address];
            for (var col = range._fromCol; col <= range._toCol; col++)
            {
                var styleId = WorkSheet.Cells[copyRow, col].StyleID;
                if (styleId != 0)
                {
                    for (int row = range._fromRow; row <= range._toRow; row++)
                    {

                        WorkSheet.SetStyleInner(row, col, styleId);
                    }
                }
            }
        }
        private void CopyStylesFromColumn(string address, int copyColumn)
        {
            var range = WorkSheet.Cells[address];
            for (int row = range._fromRow; row <= range._toRow; row++)
            {
                var styleId = WorkSheet.Cells[row, copyColumn].StyleID;
                if (styleId != 0)
                {
                for (var col = range._fromCol; col <= range._toCol; col++)
                {

                        WorkSheet.SetStyleInner(row, col, styleId);
                    }
                }
            }
        }

        /// <summary>
        /// Deletes one or more rows at the specified position in the table.
        /// </summary>
        /// <param name="position">The position in the table where the row will be deleted. 0 will delete the first row. </param>
        /// <param name="rows">Number of rows to delete.</param>
        /// <returns></returns>
        public ExcelRangeBase DeleteRow(int position, int rows = 1)
        {
            if (position < 0)
            {
                throw new ArgumentException("position", "position can't be negative");
            }
            if (rows < 0)
            {
                throw new ArgumentException("position", "rows can't be negative");
            }
            if (_address._fromRow + position + rows > _address._toRow)
            {
                throw new InvalidOperationException("Delete will exceed the number of rows in the table");
            }
            var subtract = ShowTotal ? 2 : 1;
            if(position==0 && rows+subtract >=_address.Rows)
            {
                throw new InvalidOperationException("Can't delete all table rows. A table must have at least one row.");
            }
            position++; //Header row should not be deleted.
            var address = ExcelCellBase.GetAddress(_address._fromRow + position, _address._fromCol, _address._fromRow + position + rows - 1, _address._toCol);
            var range = new ExcelRangeBase(WorkSheet, address);
            range.Delete(eShiftTypeDelete.Up);
            return range;
        }
        /// <summary>
        /// Inserts one or more columns before the specified position in the table.
        /// </summary>
        /// <param name="position">The position in the table where the column will be inserted. 0 will insert the column at the leftmost. Any value larger than the number of rows in the table will insert a row at the bottom of the table.</param>
        /// <param name="columns">Number of rows to insert.</param>
        /// <param name="copyStyles">Copy styles from the column to the left.</param>
        /// <returns>The inserted range</returns>
        internal ExcelRangeBase InsertColumn(int position, int columns, bool copyStyles=false)
        {
            if (position < 0)
            {
                throw new ArgumentException("position", "position can't be negative");
            }
            if (columns < 0)
            {
                throw new ArgumentException("columns", "columns can't be negative");
            }
            var isFirstColumn = position == 0;
            if (position >= ExcelPackage.MaxColumns || position > _address._fromCol + position + columns - 1)
            {
                position = _address.Columns;
            }

            if (_address._fromCol + position + columns - 1 > ExcelPackage.MaxColumns)
            {
                throw new InvalidOperationException("Insert will exceed the maximum number of columns in the worksheet");
            }

            var address = ExcelCellBase.GetAddress(_address._fromRow, _address._fromCol + position, _address._toRow, _address._fromCol + position + columns - 1);
            var range = new ExcelRangeBase(WorkSheet, address);

            WorksheetRangeInsertHelper.Insert(range, eShiftTypeInsert.Right, true);

            if (position == 0)
            {
                Address = new ExcelAddressBase(_address._fromRow, _address._fromCol - columns, _address._toRow, _address._toCol);
            }
            else if (range._toCol > _address._toCol)
            {
                Address = new ExcelAddressBase(_address._fromRow, _address._fromCol, _address._toRow, _address._toCol+columns);
            }
            
            if(copyStyles && isFirstColumn==false)
            {
                var copyFromCol = _address._fromCol + position - 1;
                CopyStylesFromColumn(address, copyFromCol);
            }

            return range;
        }
        /// <summary>
        /// Deletes one or more columns at the specified position in the table.
        /// </summary>
        /// <param name="position">The position in the table where the column will be deleted.</param>
        /// <param name="columns">Number of rows to delete.</param>
        /// <returns>The deleted range</returns>
        internal ExcelRangeBase DeleteColumn(int position, int columns)
        {
            if (position < 0)
            {
                throw new ArgumentException("position", "position can't be negative");
            }
            if (columns < 0)
            {
                throw new ArgumentException("columns", "columns can't be negative");
            }

            if (_address._toCol < _address._fromCol + position + columns - 1)
            {
                throw new InvalidOperationException("Delete will exceed the number of columns in the table");
            }

            var address = ExcelCellBase.GetAddress(_address._fromRow, _address._fromCol + position, _address._toRow, _address._fromCol + position + columns - 1);
            var range = new ExcelRangeBase(WorkSheet, address);

            WorksheetRangeDeleteHelper.Delete(range, eShiftTypeDelete.Left);

            return range;
        }
        internal int? HeaderRowBorderDxfId
        {
            get
            {
                return GetXmlNodeIntNull("@headerRowBorderDxfId");
            }
            set
            {
                SetXmlNodeInt("@headerRowBorderDxfId", value);
            }
        }
        public ExcelDxfBorderBase HeaderRowBorderStyle { get; set; }
        internal int? TableBorderDxfId
        {
            get
            {
                return GetXmlNodeIntNull("@tableBorderDxfId");
            }
            set
            {
                SetXmlNodeInt("@tableBorderDxfId", value);
            }
        }
        public ExcelDxfBorderBase TableBorderStyle { get; set; }

        #region Sorting
        private TableSorter _tableSorter = null;
        const string SortStatePath = "d:sortState";
        SortState _sortState = null;

        public SortState SortState
        {
            get
            {
                if (_sortState == null)
                {
                    var node = TableXml.SelectSingleNode($"//{SortStatePath}", NameSpaceManager);
                    if (node == null) return null;
                    _sortState = new SortState(NameSpaceManager, node);
                }
                return _sortState;
            }
        }

        internal void SetTableSortState(int[] columns, bool[] descending, CompareOptions compareOptions, Dictionary<int, string[]> customLists)
        {
            //Set sort state
            var sortState = new SortState(Range.Worksheet.NameSpaceManager, this);
            var dataRange = DataRange;
            sortState.Ref = dataRange.Address;
            sortState.CaseSensitive = (compareOptions == CompareOptions.IgnoreCase || compareOptions == CompareOptions.OrdinalIgnoreCase);
            for (var ix = 0; ix < columns.Length; ix++)
            {
                bool? desc = null;
                if (descending.Length > ix && descending[ix])
                {
                    desc = true;
                }
                var adr = ExcelCellBase.GetAddress(dataRange._fromRow, dataRange._fromCol + columns[ix], dataRange._toRow, dataRange._fromCol + columns[ix]);
                if(customLists.ContainsKey(columns[ix]))
                {
                    sortState.SortConditions.Add(adr, desc, customLists[columns[ix]]);
                }
                else
                {
                    sortState.SortConditions.Add(adr, desc);
                }
            }
        }

        /// <summary>
        /// Sorts the data in the table according to the supplied <see cref="RangeSortOptions"/>
        /// </summary>
        /// <param name="options"></param>
        /// <example> 
        /// <code>
        /// var options = new SortOptions();
        /// options.SortBy.Column(0).ThenSortBy.Column(1, eSortDirection.Descending);
        /// </code>
        /// </example>
        public void Sort(TableSortOptions options)
        {
            _tableSorter.Sort(options);
        }

        /// <summary>
        /// Sorts the data in the table according to the supplied action of <see cref="RangeSortOptions"/>
        /// </summary>
        /// <example> 
        /// <code>
        /// table.Sort(x =&gt; x.SortBy.Column(0).ThenSortBy.Column(1, eSortDirection.Descending);
        /// </code>
        /// </example>
        /// <param name="configuration">An action with parameters for sorting</param>
        public void Sort(Action<TableSortOptions> configuration)
        {
            _tableSorter.Sort(configuration);
        }

        #endregion
    }
}

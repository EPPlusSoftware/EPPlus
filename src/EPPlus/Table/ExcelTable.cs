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

namespace OfficeOpenXml.Table
{
    /// <summary>
    /// An Excel Table
    /// </summary>
    public class ExcelTable : XmlHelper, IEqualityComparer<ExcelTable>
    {
        internal ExcelTable(Packaging.ZipPackageRelationship rel, ExcelWorksheet sheet) : 
            base(sheet.NameSpaceManager)
        {
            WorkSheet = sheet;
            TableUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
            RelationshipID = rel.Id;
            var pck = sheet._package.Package;
            Part=pck.GetPart(TableUri);

            TableXml = new XmlDocument();
            LoadXmlSafe(TableXml, Part.GetStream());
            init();
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
            TopNode = TableXml.DocumentElement;

            init();

            //If the table is just one row we can not have a header.
            if (address._fromRow == address._toRow)
            {
                ShowHeader = false;
            }
            if(AutoFilterAddress!=null)
            {
                SetAutoFilter();
            }
        }

        private void init()
        {
            TopNode = TableXml.DocumentElement;
            SchemaNodeOrder = new string[] { "autoFilter", "tableColumns", "tableStyleInfo" };
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
                SetXmlNodeString(NAME_PATH, value);
                SetXmlNodeString(DISPLAY_NAME_PATH, ExcelAddressUtil.GetValidName(value));
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
                SetXmlNodeString("@ref",value.Address);
                WriteAutoFilter(ShowTotal);
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
                }
                else
                {
                    SetXmlNodeString(HEADERROWCOUNT_PATH, "0");
                    DeleteAllNode(AUTOFILTER_ADDRESS_PATH);
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
                    }
                    else
                    {
                        DeleteNode(TOTALSROWCOUNT_PATH);
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
                if (WorkSheet.Workbook.Styles.NamedStyles.FindIndexByID(value) < 0)
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
                if (WorkSheet.Workbook.Styles.NamedStyles.FindIndexByID(value) < 0)
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
                if (WorkSheet.Workbook.Styles.NamedStyles.FindIndexByID(value) < 0)
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
    }
}

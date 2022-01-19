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
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Constants;
using OfficeOpenXml.Drawing.Slicer;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Table
{

    /// <summary>
    /// A table column
    /// </summary>
    public class ExcelTableColumn : ExcelTableDxfBase
    {
        internal ExcelTable _tbl;
        internal ExcelTableColumn(XmlNamespaceManager ns, XmlNode topNode, ExcelTable tbl, int pos) :
            base(ns, topNode)
        {
            _tbl = tbl;
            InitDxf(tbl.WorkSheet.Workbook.Styles, null, this);
            Position = pos;
        }
        /// <summary>
        /// The column id
        /// </summary>
        public int Id 
        {
            get
            {
                return GetXmlNodeInt("@id");
            }
            set
            {
                SetXmlNodeString("@id", value.ToString());
            }
        }
        /// <summary>
        /// The position of the column
        /// </summary>
        public int Position
        {
            get;
            internal set;
        }
        /// <summary>
        /// The name of the column
        /// </summary>
        public string Name
        {
            get
            {
                var n=GetXmlNodeString("@name");
                if (string.IsNullOrEmpty(n))
                {
                    if (_tbl.ShowHeader)
                    {
                        n = ConvertUtil.ExcelDecodeString(_tbl.WorkSheet.GetValue<string>(_tbl.Address._fromRow, _tbl.Address._fromCol + this.Position));
                    }
                    else
                    {
                        n = "Column" + (this.Position+1).ToString();
                    }
                }
                return n;
            }
            set
            {
                var v = ConvertUtil.ExcelEncodeString(value);
                SetXmlNodeString("@name", v);
                if (_tbl.ShowHeader)
                {
                    var cellValue = _tbl.WorkSheet.GetValue(_tbl.Address._fromRow, _tbl.Address._fromCol + Position);
                    if (v.Equals(cellValue?.ToString(),StringComparison.CurrentCultureIgnoreCase)==false)
                    {
                        _tbl.WorkSheet.SetValue(_tbl.Address._fromRow, _tbl.Address._fromCol + Position, value);
                    }
                }
                _tbl.WorkSheet.SetTableTotalFunction(_tbl, this);
            }
        }
        /// <summary>
        /// A string text in the total row
        /// </summary>
        public string TotalsRowLabel
        {
            get
            {
                return GetXmlNodeString("@totalsRowLabel");
            }
            set
            {
                SetXmlNodeString("@totalsRowLabel", value);
                _tbl.WorkSheet.SetValueInner(_tbl.Address._toRow, _tbl.Address._fromCol+Position, value);
            }
        }
        /// <summary>
        /// Build-in total row functions.
        /// To set a custom Total row formula use the TotalsRowFormula property
        /// <seealso cref="TotalsRowFormula"/>
        /// </summary>
        public RowFunctions TotalsRowFunction
        {
            get
            {
                if (GetXmlNodeString("@totalsRowFunction") == "")
                {
                    return RowFunctions.None;
                }
                else
                {
                    return (RowFunctions)Enum.Parse(typeof(RowFunctions), GetXmlNodeString("@totalsRowFunction"), true);
                }
            }
            set
            {
                if (value == RowFunctions.Custom)
                {
                    throw(new Exception("Use the TotalsRowFormula-property to set a custom table formula"));
                }
                string s = value.ToString();
                s = s.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + s.Substring(1, s.Length - 1);
                SetXmlNodeString("@totalsRowFunction", s);
                _tbl.WorkSheet.SetTableTotalFunction(_tbl, this);
            }
        }
        const string TOTALSROWFORMULA_PATH = "d:totalsRowFormula";
        /// <summary>
        /// Sets a custom Totals row Formula.
        /// Be carefull with this property since it is not validated. 
        /// <example>
        /// tbl.Columns[9].TotalsRowFormula = string.Format("SUM([{0}])",tbl.Columns[9].Name);
        /// </example>
        /// </summary>
        public string TotalsRowFormula
        {
            get
            {
                return GetXmlNodeString(TOTALSROWFORMULA_PATH);
            }
            set
            {
                if(!string.IsNullOrEmpty(value))
                {
                    if (value.StartsWith("=")) value = value.Substring(1, value.Length - 1);
                }
                SetXmlNodeString("@totalsRowFunction", "custom");                
                SetXmlNodeString(TOTALSROWFORMULA_PATH, value);
                _tbl.WorkSheet.SetTableTotalFunction(_tbl, this);
            }
        }
        const string DATACELLSTYLE_PATH = "@dataCellStyle";
        /// <summary>
        /// The named style for datacells in the column
        /// </summary>
        public string DataCellStyleName
        {
            get
            {
                return GetXmlNodeString(DATACELLSTYLE_PATH);
            }
            set
            {
                if(_tbl.WorkSheet.Workbook.Styles.NamedStyles.FindIndexById(value)<0)
                {
                    throw(new Exception(string.Format("Named style {0} does not exist.",value)));
                }
                SetXmlNodeString(TopNode, DATACELLSTYLE_PATH, value,true);
               
                int fromRow=_tbl.Address._fromRow + (_tbl.ShowHeader?1:0),
                    toRow=_tbl.Address._toRow - (_tbl.ShowTotal?1:0),
                    col=_tbl.Address._fromCol+Position;

                if (fromRow <= toRow)
                {
                    _tbl.WorkSheet.Cells[fromRow, col, toRow, col].StyleName = value;
                }
            }
        }
  		const string CALCULATEDCOLUMNFORMULA_PATH = "d:calculatedColumnFormula";

        ExcelTableSlicer _slicer = null;
        /// <summary>
        /// Returns the slicer attached to a column.
        /// If the column has multiple slicers, the first is returned.
        /// </summary>
        public ExcelTableSlicer Slicer 
        {
            get
            {
                if (_slicer == null)
                {
                    var wb = _tbl.WorkSheet.Workbook;
                    if (wb.ExistsNode($"d:extLst/d:ext[@uri='{ExtLstUris.WorkbookSlicerTableUri}']"))
                    {
                        foreach (var ws in wb.Worksheets)
                        {
                            foreach (var d in ws.Drawings)
                            {
                                if (d is ExcelTableSlicer s && s.TableColumn == this)
                                {
                                    _slicer = s;
                                    return _slicer;
                                }
                            }
                        }
                    }
                }
                return _slicer;
            }
            internal set
            {
                _slicer = value;
            }
        }
        public ExcelTableSlicer AddSlicer()
        {            
            return _tbl.WorkSheet.Drawings.AddTableSlicer(this);
        }
        /// <summary>
        /// Sets a calculated column Formula.
        /// Be carefull with this property since it is not validated. 
        /// <example>
        /// tbl.Columns[9].CalculatedColumnFormula = string.Format("SUM(MyDataTable[[#This Row],[{0}]])",tbl.Columns[9].Name);  //Reference within the current row
        /// tbl.Columns[9].CalculatedColumnFormula = string.Format("MyDataTable[[#Headers],[{0}]]",tbl.Columns[9].Name);  //Reference to a column header
        /// tbl.Columns[9].CalculatedColumnFormula = string.Format("MyDataTable[[#Totals],[{0}]]",tbl.Columns[9].Name);  //Reference to a column total        
        /// </example>
        /// </summary>
        public string CalculatedColumnFormula
 		{
 			get
 			{
 				return GetXmlNodeString(CALCULATEDCOLUMNFORMULA_PATH);
 			}
 			set
 			{
 				if (value.StartsWith("=")) value = value.Substring(1, value.Length - 1);
 				SetXmlNodeString(CALCULATEDCOLUMNFORMULA_PATH, value);

                SetTableFormula();
 			}
 		}
        public ExcelTable Table
        {
            get
            {
                return _tbl;
            }
        }
        private void SetTableFormula()
        {
            int fromRow = _tbl.ShowHeader ? _tbl.Address._fromRow + 1 : _tbl.Address._fromRow;
            int toRow = _tbl.ShowTotal ? _tbl.Address._toRow - 1 : _tbl.Address._toRow;
            var colNum = _tbl.Address._fromCol + Position;
            SetFormulaCells(fromRow, toRow, colNum);
        }

        internal void SetFormulaCells(int fromRow, int toRow, int colNum)
        {
            string r1c1Formula = ExcelCellBase.TranslateToR1C1(CalculatedColumnFormula, _tbl.ShowHeader ? _tbl.Address._fromRow + 1 : _tbl.Address._fromRow, colNum);
            bool needsTranslation = r1c1Formula != CalculatedColumnFormula;

            for (int row = fromRow; row <= toRow; row++)
            {
                _tbl.WorkSheet.SetFormula(row, colNum, needsTranslation ? ExcelCellBase.TranslateFromR1C1(r1c1Formula, row, colNum) : r1c1Formula);
            }
        }
    }
}

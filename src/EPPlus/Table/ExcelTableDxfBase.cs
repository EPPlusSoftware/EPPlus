using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using System;
using System.Drawing;
using System.Xml;

namespace OfficeOpenXml.Table
{
    public class ExcelTableDxfBase : XmlHelper
    {
        private ExcelTable _table=null;
        private ExcelTableColumn _tableColumn = null;
        internal ExcelTableDxfBase(XmlNamespaceManager nsm) : base(nsm)
        {
        }
        internal ExcelTableDxfBase(XmlNamespaceManager nsm, XmlNode topNode) : base(nsm, topNode)
        {
        }
        internal void InitDxf(ExcelStyles styles, ExcelTable table, ExcelTableColumn tableColumn)
        {
            _table = table;
            _tableColumn = tableColumn;
            HeaderRowStyle = styles.GetDxf(HeaderRowDxfId, SetHeaderStyle);
            DataStyle = styles.GetDxf(DataDxfId, SetDataStyle);
            TotalsRowStyle = styles.GetDxf(TotalsRowDxfId, SetTotalsStyle);
        }
        internal int? HeaderRowDxfId
        {
            get
            {
                return GetXmlNodeIntNull("@headerRowDxfId");
            }
            set
            {
                SetXmlNodeInt("@headerRowDxfId", value);
            }
        }
        internal string HeaderRowStyleName
        {
            get
            {
                return GetXmlNodeString("@headerRowCellStyle");
            }
            set
            {
                SetXmlNodeString("@headerRowCellStyle",value);
            }
        }

        public ExcelDxfStyle HeaderRowStyle { get; internal set; }
        internal int? DataDxfId
        {
            get
            {
                return GetXmlNodeIntNull("@dataDxfId");
            }
            set
            {
                SetXmlNodeInt("@dataDxfId", value);
            }
        }
        public ExcelDxfStyle DataStyle { get; internal set; }
        public ExcelDxfStyle TotalsRowStyle { get;  internal set; }
        internal int? TotalsRowDxfId
        {
            get
            {
                return GetXmlNodeIntNull("@totalsRowDxfId");
            }
            set
            {
                SetXmlNodeInt("@totalsRowDxfId", value);
            }
        }

        void SetHeaderStyle(eStyleClass styleClass, eStyleProperty styleProperty, object value)
        {
            if ((_table ?? _tableColumn.Table).ShowHeader == false || value==null) return;
            ExcelRangeBase headerRange;
            if(_tableColumn==null)
            {
                headerRange = _table.Range.Offset(0, 0, 1, _table.Range.Columns);
            }
            else
            {
                var tbl = _tableColumn.Table;
                headerRange = tbl.Range.Offset(0, _tableColumn.Position, 1, 1);
            }
            SetStyle(headerRange, styleClass, styleProperty, value);
        }

        void SetDataStyle(eStyleClass styleClass, eStyleProperty styleProperty, object value)
        {
            if (value == null) return;
            ExcelRangeBase range;
            if (_tableColumn == null)
            {
                range = _table.DataRange;
            }
            else
            {
                var tbl = _tableColumn.Table;
                range = tbl.DataRange.Offset(0, _tableColumn.Position, tbl.DataRange.Rows, 1);
            }
            SetStyle(range, styleClass, styleProperty, value);
        }
        void SetTotalsStyle(eStyleClass styleClass, eStyleProperty styleProperty, object value)
        {
            if ((_table??_tableColumn.Table).ShowTotal == false || value==null) return;
            ExcelRangeBase totalRange;
            if (_tableColumn == null)
            {
                totalRange = _table.Range.Offset(_table.Range.Rows-1, 0, 1, _table.Range.Columns);
            }
            else
            {
                var tbl = _tableColumn.Table;
                totalRange = tbl.Range.Offset(tbl.Range.Rows-1, _tableColumn.Position, 1, 1);
            }
            SetStyle(totalRange, styleClass, styleProperty, value);
        }
        private void SetStyle(ExcelRangeBase headerRange, eStyleClass styleClass, eStyleProperty styleProperty, object value)
        {
            switch(styleClass)
            {
                case eStyleClass.Fill:
                    SetStyleFill(headerRange, styleProperty, value);
                    break;
                case eStyleClass.FillPatternColor:
                    SetStyleColor(headerRange.Style.Fill.PatternColor, styleProperty, value);
                    break;
                case eStyleClass.FillBackgroundColor:
                    SetStyleColor(headerRange.Style.Fill.BackgroundColor, styleProperty, value);
                    break;
                case eStyleClass.GradientFill:
                    SetStyleGradient(headerRange, styleProperty, value);
                    break;
                case eStyleClass.FillGradientColor1:
                    SetStyleColor(headerRange.Style.Fill.Gradient.Color1, styleProperty, value);
                    break;
                case eStyleClass.FillGradientColor2:
                    SetStyleColor(headerRange.Style.Fill.Gradient.Color2, styleProperty, value);
                    break;
                case eStyleClass.BorderTop:
                    SetStyleBorder(headerRange.Style.Border.Top, styleProperty, value);
                    break;
                case eStyleClass.BorderBottom:
                    SetStyleBorder(headerRange.Style.Border.Bottom, styleProperty, value);
                    break;
                case eStyleClass.BorderLeft:
                    SetStyleBorder(headerRange.Style.Border.Left, styleProperty, value);
                    break;
                case eStyleClass.BorderRight:
                    SetStyleBorder(headerRange.Style.Border.Right, styleProperty, value);
                    break;
                case eStyleClass.Font:
                    SetStyleFont(headerRange, styleProperty, value);
                    break;
                case eStyleClass.Numberformat:
                    SetStyleNumberFormat(headerRange, styleProperty, value);
                    break;
            }
        }
        private void SetStyleNumberFormat(ExcelRangeBase range, eStyleProperty styleProperty, object value)
        {
            switch (styleProperty)
            {
                case eStyleProperty.Format:
                    if(value is int n)
                    {
                        range.Style.Numberformat.Format = ExcelNumberFormat.GetFromBuildInFromID(n);
                    }
                    else
                    {
                        range.Style.Numberformat.Format = value.ToString();
                    }
                    break;
            }
        }
        private void SetStyleFont(ExcelRangeBase headerRange, eStyleProperty styleProperty, object value)
        {
            switch(styleProperty)
            {
                case eStyleProperty.Name:
                    headerRange.Style.Font.Name = value.ToString();
                    break;
                case eStyleProperty.Bold:
                    headerRange.Style.Font.Bold = (bool)value;
                    break;
                case eStyleProperty.Italic:
                    headerRange.Style.Font.Italic = (bool)value;
                    break;
                case eStyleProperty.UnderlineType:
                    headerRange.Style.Font.UnderLineType = (ExcelUnderLineType)value;
                    break;
                case eStyleProperty.Strike:
                    headerRange.Style.Font.Strike = (bool)value;
                    break;
                case eStyleProperty.AutoColor:
                case eStyleProperty.Color:
                case eStyleProperty.Theme:
                case eStyleProperty.IndexedColor:
                    SetStyleColor(headerRange.Style.Font.Color, styleProperty, value);
                    break;
                case eStyleProperty.Size:
                    headerRange.Style.Font.Size = (float)value;
                    break;
                case eStyleProperty.Family:
                    headerRange.Style.Font.Family = (int)value;
                    break;
                case eStyleProperty.Charset:
                    headerRange.Style.Font.Charset = (int)value;
                    break;
                case eStyleProperty.VerticalAlign:
                    headerRange.Style.Font.VerticalAlign = (ExcelVerticalAlignmentFont)value;
                    break;
            }
        }

        private void SetStyleBorder(ExcelBorderItem border, eStyleProperty styleProperty, object value)
        {
            switch (styleProperty)
            {
                case eStyleProperty.AutoColor:
                case eStyleProperty.Color:
                case eStyleProperty.Theme:
                case eStyleProperty.IndexedColor:
                    SetStyleColor(border.Color, styleProperty, value);
                    break;
                case eStyleProperty.Style:
                    border.Style = (ExcelBorderStyle)value;
                    break;
            }
        }

            private void SetStyleColor(ExcelColor color, eStyleProperty styleProperty, object value)
        {            
            switch (styleProperty)
            {
                case eStyleProperty.AutoColor:
                    color.SetAuto();
                    break;
                case eStyleProperty.Color:
                    color.SetColor((Color)value);
                    break;
                case eStyleProperty.Theme:
                    color.SetColor((eThemeSchemeColor)value);
                    break;
                case eStyleProperty.IndexedColor:
                    color.SetColor((ExcelIndexedColor)value);
                    break;
                case eStyleProperty.Tint:
                    color.Tint = (decimal)value;
                    break;
            }
        }
        private void SetStyleFill(ExcelRangeBase headerRange, eStyleProperty styleProperty, object value)
        {
            switch(styleProperty)
            {
                case eStyleProperty.PatternType:
                    headerRange.Style.Fill.PatternType = (ExcelFillStyle)value;                    
                    break;
            }
        }
        private void SetStyleGradient(ExcelRangeBase headerRange, eStyleProperty styleProperty, object value)
        {
            switch (styleProperty)
            {
                case eStyleProperty.GradientDegree:
                    headerRange.Style.Fill.Gradient.Degree = (double)(value ?? 0D);
                    break;
                case eStyleProperty.GradientTop:
                    headerRange.Style.Fill.Gradient.Top = (double)(value ?? 0D);
                    break;
                case eStyleProperty.GradientBottom:
                    headerRange.Style.Fill.Gradient.Bottom = (double)(value ?? 0D);
                    break;
                case eStyleProperty.GradientLeft:
                    headerRange.Style.Fill.Gradient.Left = (double)(value ?? 0D);
                    break;
                case eStyleProperty.GradientRight:
                    headerRange.Style.Fill.Gradient.Right = (double)(value ?? 0D);
                    break;
            }
        }
    }
}

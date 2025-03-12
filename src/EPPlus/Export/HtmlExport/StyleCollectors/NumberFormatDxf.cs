using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;

using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors
{
    internal class NumberFormatDxf : INumberFormat
    {
        ExcelDxfNumberFormat NumberFormat;
        ExcelFormatTranslator FormatTranslator;
        int StyleId;

        internal NumberFormatDxf(ExcelDxfNumberFormat format, int styleId)
        {
            NumberFormat = format;
            FormatTranslator = ValueToTextHandler.GetDxfNumberFormat(styleId, format._styles).FormatTranslator;
            StyleId = styleId;
        }

        public string NumberFormatString => NumberFormat.Format;

        public int NumberFormatID => NumberFormat.NumFmtID;

        public ExcelIndexedColor? ColorId => FormatTranslator.NumFtColor;

        public bool HasValue
        {
            get
            {
                return !string.IsNullOrEmpty(NumberFormat.Id);
            }
        }
    }
}

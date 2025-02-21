using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts
{
    internal interface INumberFormat
    {
        string NumberFormatString { get; }
        int NumberFormatID { get; }
        public ExcelIndexedColor? ColorId { get; }
        internal bool HasValue { get; }
    }
}

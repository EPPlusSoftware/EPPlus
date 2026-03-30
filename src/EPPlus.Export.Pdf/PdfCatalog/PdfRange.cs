using OfficeOpenXml
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.PdfCatalog
{
    internal struct PdfRange
    {
        public ExcelRangeBase Range { get; set; }
        public bool ExtendColumns { get; set; }
    }
}

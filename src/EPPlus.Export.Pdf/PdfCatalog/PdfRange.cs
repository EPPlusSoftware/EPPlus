using OfficeOpenXml;
using System.Collections.Generic;

namespace EPPlus.Export.Pdf.PdfCatalog
{
    internal struct PdfRange
    {
        public ExcelRangeBase Range { get; set; }
        public bool ExtendColumns { get; set; }
        public PdfCellCollection Map { get; set; }
        public List<double> RowHeights = new List<double>();
        public List<double> ColWidths = new List<double>();
        public double TotalHeight;
        public double TotalWidth;
        public double AdditionalHeight;
        public double AdditionalWidth;

        public PdfRange(ExcelRangeBase range, bool extendColumns)
        {
            Range = range;
            ExtendColumns = extendColumns;
        }
    }
}

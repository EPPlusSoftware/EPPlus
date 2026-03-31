using OfficeOpenXml;

namespace EPPlus.Export.Pdf.PdfCatalog
{
    internal struct PdfRange
    {
        public ExcelRangeBase Range { get; set; }
        public bool ExtendColumns { get; set; }
        public PdfCellCollection Map { get; set; }

        public PdfRange(ExcelRangeBase range, bool extendColumns)
        {
            Range = range;
            ExtendColumns = extendColumns;
        }
    }
}

using OfficeOpenXml.Style;

namespace EPPlus.Export.Pdf.PdfLayout
{
    internal enum TableCellStyle
    {
        Header,
        OddRow,
        EvenRow,
        TotalRow,
        FirstColumn,
        LastColumn,
        OddColumn,
        EvenColumn,
        WholeTable
    }
    internal class PdfTableLayout
    {
        internal TableCellStyle TableCellStyleType { get; set; }
        internal ExcelTableStyleElement Style;
    }
}

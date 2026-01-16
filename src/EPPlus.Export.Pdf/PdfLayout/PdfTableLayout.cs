using System;
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

    [Flags]
    public enum TableBorderStyle
    {
        None = 0,
        Top = 1 << 0,
        Bottom = 1 << 1,
        Left = 1 << 2,
        Right = 1 << 3,
        Horizontal = 1 << 4,
        Vertical = 1 << 5
    }

    internal class PdfTableLayout
    {
        internal TableCellStyle TableCellStyleType { get; set; }
        internal ExcelTableStyleElement MainStyle;
        internal ExcelTableStyleElement WholeStyle;
        internal TableBorderStyle borderStyleType = TableBorderStyle.None;
        //var to hold border styles or check when setting borders?
    }
}

using System;
using EPPlus.Export.Pdf.Pdfhelpers;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;

namespace EPPlus.Export.Pdf.PdfLayout
{
    /// <summary>
    /// Holds the styles for the cell.
    /// xf styles is the base style and dxf styles is an override(sort of) that can come from tables or conditional formatting.
    /// They are used in the following priority:
    /// 1. xf style that is not default
    /// 2. dxf style if it exsist
    /// 3. deault xf style
    /// </summary>
    internal class PdfCellStyle
    {
        //Fill
        internal ExcelFill xfFill { get; set; }
        internal ExcelDxfFill dxfFill { get; set; }

        //Borders
        internal ExcelBorderItem xfTop { get; set; }
        internal ExcelBorderItem xfBottom { get; set; }
        internal ExcelBorderItem xfLeft { get; set; }
        internal ExcelBorderItem xfRight { get; set; }
        internal bool DiagonalUp { get; set; }
        internal bool DiagonalDown { get; set; }
        internal ExcelBorderItem Diagonal { get; set; }
        internal ExcelDxfBorderItem dxfTop { get; set; }
        internal ExcelDxfBorderItem dxfBottom { get; set; }
        internal ExcelDxfBorderItem dxfLeft { get; set; }
        internal ExcelDxfBorderItem dxfRight { get; set; }
        internal ExcelDxfBorderItem dxfHorizontal { get; set; }
        internal ExcelDxfBorderItem dxfVertical { get; set; }

        //Fonts
        internal ExcelFont xfFont { get; set; }
        internal ExcelDxfFontBase dxfFont { get; set; }
    }
}

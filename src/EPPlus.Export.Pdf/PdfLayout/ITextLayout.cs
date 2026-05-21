using EPPlus.Fonts.OpenType.Integration;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System.Collections.Generic;

namespace EPPlus.Export.Pdf.PdfLayout
{
    internal interface ITextLayout
    {
        public  List<PdfTextFormat> TextFormats { get; set; }
        public double TextLength { get; set; }
        public double TextHeight { get; set; }
        public TextLayoutEngine TextLayoutEngine{ get; set; }

        //public abstract void CalculateTextSpill(double Width, double Rotation);
        //public abstract Vector2 CalculateAlignmentPositionAndTextOffsets(ExcelRangeBase cell, double x, double y, double width, double height);
        //public abstract void CheckClipping(ExcelRangeBase cell, double x, double y, double width, double height);
    }
}

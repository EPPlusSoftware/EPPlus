using EPPlus.Fonts.OpenType.Integration;
using System.Collections.Generic;

namespace EPPlus.Export.Pdf.PdfLayout
{
    internal interface ITextLayout
    {
        public  List<PdfTextFormat> TextFormats { get; set; }
        public double TextLength { get; set; }
        public double TextHeight { get; set; }
        public TextLayoutEngine TextLayoutEngine{ get; set; }
    }
}

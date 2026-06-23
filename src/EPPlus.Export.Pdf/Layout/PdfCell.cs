using EPPlus.Fonts.OpenType.Integration;
using System.Collections.Generic;

namespace EPPlus.Export.Pdf.Layout
{
    internal class PdfCellBase
    {
        public string Name { get; set; }
        public bool Hidden;
        public PdfCellAlignmentData ContentAligmnet;
        public List<TextFragment> TextFragments { get; set; }
        public List<PdfShapedText> ShapedTexts { get; set; }
        public TextLineCollection TextLines { get; set; }
        public string Text { get; set; }

        public double TotalTextLength { get; set; }
        public double ColumnWidth { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }

        public TextLayoutEngine TextLayoutEngine { get; set; }

        public bool IsPrintTitleRow;
        public bool IsPrintTitleCol;
    }
}

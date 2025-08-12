using OfficeOpenXml.PDF.Math;
using OfficeOpenXml.PDF.PdfGraphics;
using OfficeOpenXml.Style;
using System.Data;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal struct PdfCellFontData
    {
        public string FontName = "Aptos Narrow";
        public int FontFamily = 0;
        public double FontSize = 11;
        public bool Bold = false;
        public bool Italic = false;
        public bool Strike = false;
        public bool SubScript = false;
        public bool SuperScript = false;
        public bool Underline = false;
        public ExcelUnderLineType UnderlineType = ExcelUnderLineType.None;
        public PdfColor FontColor = PdfColor.Black;
        public string Text = "";
        //public NumberFormatting;
        //public ExcelRichTextCollection RichText;

        public PdfCellFontData()
        {
        }
    }

    internal struct PdfCellFillData
    {
        public PdfColor BackgroundColor = PdfColor.None;
        public ExcelFillStyle PattenStyle = ExcelFillStyle.None;
        public PdfColor PatternColor = PdfColor.Black;
        //Fill Effects

        public PdfCellFillData()
        {
        }
    }

    internal struct PdfCellBordersData
    {
        public PdfCellBorderData Top;
        public PdfCellBorderData Bottom;
        public PdfCellBorderData Left;
        public PdfCellBorderData Right;
        public PdfCellBorderData DiagonalUp;
        public PdfCellBorderData DiagonalDown;
    }

    internal struct PdfCellBorderData
    {
        public ExcelBorderStyle BorderStyle = ExcelBorderStyle.None;
        public PdfColor BorderColor = PdfColor.Black;

        public PdfCellBorderData()
        {
        }
    }

    internal struct PdfCellAlignmentData
    {
        public ExcelHorizontalAlignment HorizontalAlignment = ExcelHorizontalAlignment.General;
        public ExcelVerticalAlignment VerticalAlignment = ExcelVerticalAlignment.Bottom;
        public int Indent = 0;
        public bool WrapText = false;
        public bool ShrinkToFit = false;
        public int TextRotation = 0;
        public ExcelReadingOrder TextDirection = ExcelReadingOrder.ContextDependent;

        public PdfCellAlignmentData()
        {
        }
    }

}

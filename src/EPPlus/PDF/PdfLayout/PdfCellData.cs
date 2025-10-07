using OfficeOpenXml.PDF.Math;
using OfficeOpenXml.PDF.PdfGraphics;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal struct PdfCellFontData
    {
        public string FontName = "Aptos Narrow";
        public int FontFamily = 0;
        public string SubFamily = "Regular";
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
        public string Font
        { get { return FontName + " " + SubFamily; } }

        public PdfCellFontData() { }
    }

    internal class PdfCellGradientFillData
    {
        public ExcelFillGradientType GradientType;
        public PdfColor Color1;
        public PdfColor Color2;
        public PdfColor Color3;
        public double Degree;
        public double Top;
        public double Bottom;
        public double Left;
        public double Right;
        public double[] matrix;
        public double[] coords;

        public override string ToString()
        {
            return GradientType.ToString() + Color1.ToHexString() + Color2.ToHexString() + Degree + Top + Bottom + Left + Right;
        }
    }

    internal struct PdfCellFillData
    {
        public string id;
        public PdfColor BackgroundColor = PdfColor.None;
        public ExcelFillStyle PattenStyle = ExcelFillStyle.None;
        public PdfColor PatternColor = PdfColor.Black;
        //Fill Effects
        public PdfCellGradientFillData GradientFillData = null;

        public PdfCellFillData() { }
    }

    internal class PdfCellBordersData
    {
        public PdfCellBorderData Top = new PdfCellBorderData(2, -2);
        public PdfCellBorderData Bottom = new PdfCellBorderData(2, 2);
        public PdfCellBorderData Left = new PdfCellBorderData(2, -2);
        public PdfCellBorderData Right = new PdfCellBorderData(-2, -2);
        public PdfCellBorderData DiagonalUp = new PdfCellBorderData(2, 2);
        public PdfCellBorderData DiagonalDown = new PdfCellBorderData(2, 2);

        public PdfCellBordersData() { }
    }

    internal class PdfCellBorderData
    {
        public ExcelBorderStyle BorderStyle = ExcelBorderStyle.None;
        public PdfColor BorderColor = PdfColor.Black;
        public Vector2 DoubleBorderOffsets;

        public PdfCellBorderData(double x, double y)
        {
            DoubleBorderOffsets = new Vector2(x, y);
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

        public PdfCellAlignmentData() { }
    }

    internal struct GridLine
    {
        public static double Width = 0.125d;
        public static double HalfWidth = Width / 2d;
        public static double FourthWidth = Width / 4d;
        public double X1;
        public double Y1;
        public double X2;
        public double Y2;

        public GridLine(double x1, double y1, double x2, double y2)
        {
            X1 = x1; Y1 = y1; X2 = x2; Y2 = y2;
        }
    }

}

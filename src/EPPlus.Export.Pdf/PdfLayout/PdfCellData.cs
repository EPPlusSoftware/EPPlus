/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using EPPlus.Graphics.Math;
using System.Drawing;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using EPPlus.Export.Pdf.Pdfhelpers;

namespace EPPlus.Export.Pdf.PdfLayout
{
    internal class PdfTextLine
    {
        public string Text;
        public double TextLength;
        public double LineHeight;
        public double FontHeight;
        public double Offset;
    }

    internal class PdfCellFontData
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
        public Color FontColor = Color.Black;
        public List<PdfTextLine> Lines = new List<PdfTextLine>();
        public double TextLength = 0d;
        public double LineHeight = 0d;
        public double FontHeight = 0d;
        //public NumberFormatting;
        //public ExcelRichTextCollection RichText;
        public string FullFontName
        { get { return FontName + " " + SubFamily; } }

        public PdfCellFontData() { }
    }

    internal class PdfCellGradientFillData
    {
        public ExcelFillGradientType GradientType;
        public Color Color1;
        public Color Color2;
        public Color Color3;
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

    internal class PdfCellFillData
    {
        public string id;
        public Color BackgroundColor = Color.Empty;
        public ExcelFillStyle PattenStyle = ExcelFillStyle.None;
        public Color PatternColor = Color.Black;
        //Fill Effects
        public PdfCellGradientFillData GradientFillData = null;
        public bool enhanceGridLine = false;
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
        public Color BorderColor = Color.Black;
        public Vector2 DoubleBorderOffsets;

        public PdfCellBorderData(double x, double y)
        {
            DoubleBorderOffsets = new Vector2(x, y);
        }
    }

    internal class PdfCellAlignmentData
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

    internal class GridLine
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

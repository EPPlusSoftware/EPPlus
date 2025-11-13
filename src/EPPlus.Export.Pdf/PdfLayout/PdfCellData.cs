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
using EPPlus.Export.Pdf.Math;
using EPPlus.Export.Pdf.PdfGraphics;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.Style;
using System.Collections.Generic;

namespace EPPlus.Export.Pdf.PdfLayout
{
    internal class PdfCellTextLine
    {
        public List<PdfCellTextItem> TextItemCollection = new List<PdfCellTextItem>();
        public bool IsRichText = false;
        public string Text;
        public double Offset;
        public double TextLength
        {
            get
            {
                double val = 0;
                foreach (var tp in TextItemCollection)
                {
                    val += tp.TextLength;
                }
                return val;
            }
        }
        public double LineHeight
        {
            get
            {
                double val = 0;
                foreach (var tp in TextItemCollection)
                {
                    val = tp.LineHeight > val ? tp.LineHeight : val;
                }
                return val;
            }
        }
        public double FontHeight
        {
            get
            {
                double val = 0;
                foreach (var tp in TextItemCollection)
                {
                    val = tp.FontHeight > val ? tp.FontHeight : val;
                }
                return val;
            }
        }
    }

    internal class PdfCellTextItem
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
        public string Text;
        public double TextLength = 0d;
        public double LineHeight = 0d;
        public double FontHeight = 0d;

        public string FullFontName
        { get { return FontName + " " + SubFamily; } }

        public PdfCellTextItem() { }
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

    internal class PdfCellFillData
    {
        public string id;
        public PdfColor BackgroundColor = PdfColor.None;
        public ExcelFillStyle PattenStyle = ExcelFillStyle.None;
        public PdfColor PatternColor = PdfColor.Black;
        //Fill Effects
        public PdfCellGradientFillData GradientFillData = null;
        public bool enhanceGridLine = false;
        public PdfCellFillData() { }
    }

    internal class PdfCellBordersData
    {
        public PdfCellBorderData Top = new PdfCellBorderData(LineType.Top);
        public PdfCellBorderData Bottom = new PdfCellBorderData(LineType.Bottom);
        public PdfCellBorderData Left = new PdfCellBorderData(LineType.Left);
        public PdfCellBorderData Right = new PdfCellBorderData(LineType.Right);
        public PdfCellBorderData DiagonalUp = new PdfCellBorderData(LineType.DiagonalUp);
        public PdfCellBorderData DiagonalDown = new PdfCellBorderData(LineType.DiagonalDown);
        public bool[] NeighbourBorder = new bool[8];

        public PdfCellBordersData() { }
    }

    internal enum LineType
    {
        Top = 0,
        Bottom,
        Left,
        Right,
        DiagonalUp,
        DiagonalDown
    }

    internal class PdfCellBorderData
    {
        internal const double Hair = 0.5d;
        internal const double Thin = 0.85d;
        internal const double Small = 1.1d;
        internal const double Medium = 1.5d;
        internal const double Thick = 2.0d;
        internal const string NoDash = "[] 0 d";
        internal const string Dotted = "[0 2] 0 d";
        internal const string DashDot = "[4 2 1 2] 0 d";
        internal const string DashDotDot = "[4 2 1 2 1 2] 0 d";
        internal const string Dashed = "[4 3] 0 d";
        internal const string MediumDashDot = "[6 3 2 3] 0 d";
        internal const string MediumDashDotDot = "[6 3 2 3 2 3] 0 d";
        internal const string MediumDashed = "[6 4] 0 d";

        public ExcelBorderStyle BorderStyle = ExcelBorderStyle.None;
        public PdfColor BorderColor = PdfColor.Black;
        public readonly LineType LineType;

        public PdfCellBorderData(LineType LineType)
        {
            this.LineType = LineType;
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

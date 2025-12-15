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
using EPPlus.Fonts.OpenType;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using EPPlus.Export.Pdf.Pdfhelpers;
using EPPlus.Graphics;

namespace EPPlus.Export.Pdf.PdfLayout
{
    public enum PdfWritingMode
    {
        HorizontalLtr,
        HorizontalRtl,
        VerticalTtb, // top-to-bottom
        VerticalBtt, // bottom-to-top
    }

    class TextToken
    {
        public bool IsWhitespace;
        public PdfCellTextItem Item;
    }

    internal class PdfCellTextLine
    {
        public List<PdfCellTextItem> TextItemCollection = new List<PdfCellTextItem>();
        public bool IsRichText = false;
        public string Text;
        public double Offset;
        public double Advance;
        public double CrossSize;
        public PdfWritingMode WritingMode { get; set; } = PdfWritingMode.HorizontalLtr;
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

        public object Clone() => this.MemberwiseClone();
    }

    internal class PdfCellTextItem
    {
        public string FontName { get; set; }
        public int FontFamily { get; set; }
        public string SubFamily { get; set; }
        public double FontSize { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Strike { get; set; }
        public bool SubScript { get; set; }
        public bool SuperScript { get; set; }
        public bool Underline { get; set; }
        public string FontName = "Aptos Narrow";
        public int FontFamily = 0;
        public FontSubFamily SubFamily = FontSubFamily.Regular;
        public double FontSize = 11;
        public bool Bold = false;
        public bool Italic = false;
        public bool Strike = false;
        public bool SubScript = false;
        public bool SuperScript = false;
        public bool Underline = false;
        public ExcelUnderLineType UnderlineType = ExcelUnderLineType.None;
        public string Text { get; set; }
        public Color FontColor { get; set; }
        public double TextLength = 0d;
        public double LineHeight = 0d;
        public double FontHeight = 0d;
        public double Advance = 0;
        public double CrossSize = 0;
        public double Ascent = 0;
        public double Descent = 0;
        public Rect GlyphBox = new Rect();
        public Dictionary<char, Vector2> characterOffset = new Dictionary<char, Vector2>();
        public List<GlyphPosition> GlyphPositions;
        public string FullFontName
        { get { return FontName + " " + SubFamily; } }

        public PdfCellTextItem() { }

        public object Clone() => this.MemberwiseClone();
    }

    internal class GlyphPosition
    {
        public char Character;
        public double AdvanceX;
        public double AdvanceY;
        public double OffsetX;
        public double OffsetY;
        public Rect GlyphBox;
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
        public PdfCellBorderData Top = new PdfCellBorderData(LineType.Top);
        public PdfCellBorderData Bottom = new PdfCellBorderData(LineType.Bottom);
        public PdfCellBorderData Left = new PdfCellBorderData(LineType.Left);
        public PdfCellBorderData Right = new PdfCellBorderData(LineType.Right);
        public PdfCellBorderData DiagonalUp = new PdfCellBorderData(LineType.DiagonalUp);
        public PdfCellBorderData DiagonalDown = new PdfCellBorderData(LineType.DiagonalDown);

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
        public readonly LineType LineType;
        public Color BorderColor = Color.Black;
        public double MergedDiagonalWidth = 0d;
        public double MergedDiagonalHeight = 0d;
        public double Width = 0;
        public double Height = 0;
        public double X = 0;
        public double Y = 0;

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
        public bool IsVertical = false;

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

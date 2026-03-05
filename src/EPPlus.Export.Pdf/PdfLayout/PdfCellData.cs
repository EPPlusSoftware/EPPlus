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
using EPPlus.Export.Pdf.Pdfhelpers;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Graphics;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.Drawing;

/*
 * This file could probably be trimmed a lot and remove unnessacry stuff. 
 */
namespace EPPlus.Export.Pdf.PdfLayout
{
    public enum PdfWritingMode
    {
        HorizontalLtr,
        HorizontalRtl,
        VerticalTtb, // top-to-bottom
        VerticalBtt, // bottom-to-top
    }

    internal class PdfCellLines
    {
        public bool IsRichText = false;
        public PdfWritingMode WritingMode { get; set; } = PdfWritingMode.HorizontalLtr;

        public List<PdfCellLine> Lines = new List<PdfCellLine>();
        private string _text = null;
        public string Text
        {
            get
            {
                _text = string.Empty;
                foreach (var t in Lines)
                {
                    _text += t.Text;
                }
                return _text;
            }
        }

        public double Height
        {
            get
            {
                double val = 0;
                foreach (var l in Lines)
                {
                    val += l.LineHeight;
                }
                return val;
            }
        }

        public double TextLength
        {
            get
            {
                double val = 0;
                foreach (var tp in Lines)
                {
                    val = tp.TextLength > val ? tp.TextLength : val;
                }
                return val;
            }
        }
        public double TextHeight
        {
            get
            {
                double val = 0;
                foreach (var tp in Lines)
                {
                    val = tp.TextHeight > val ? tp.TextHeight : val;
                }
                return val;
            }
        }
        public double LineHeight
        {
            get
            {
                double val = 0;
                foreach (var tp in Lines)
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
                foreach (var tp in Lines)
                {
                    val = tp.FontHeight > val ? tp.FontHeight : val;
                }
                return val;
            }
        }
    }

    internal class PdfCellLine
    {
        public double Offset = 0d;
        public List<PdfCellWord> Words = new List<PdfCellWord>();
        private string _text = null;
        public string Text
        {
            get
            {
                _text = string.Empty;
                foreach (var t in Words)
                {
                    _text += t.Text;
                }
                return _text;
            }
        }

        public double TextLength
        {
            get
            {
                double val = 0;
                foreach (var w in Words)
                {
                    val += w.TextLength;
                }
                return val;
            }
        }
        public double TextHeight
        {
            get
            {
                double val = 0;

                int first = 0;
                int last = Words.Count - 1;

                while (first <= last && PdfString.IsNullOrWhiteSpace(Words[first].Text))
                    first++;

                while (last >= first && PdfString.IsNullOrWhiteSpace(Words[last].Text))
                    last--;

                // sum heights between them
                for (int i = first; i <= last; i++)
                {
                    val += Words[i].TextHeight;
                }

                return val;
            }
        }
        public double LineHeight
        {
            get
            {
                double val = 0;
                foreach (var w in Words)
                {
                    val = w.LineHeight > val ? w.LineHeight : val;
                }
                return val;
            }
        }
        public double FontHeight
        {
            get
            {
                double val = 0;
                foreach (var w in Words)
                {
                    val = w.FontHeight > val ? w.FontHeight : val;
                }
                return val;
            }
        }
    }

    internal class PdfCellWord
    {
        public List<PdfTextFormat> Characters = new List<PdfTextFormat>();
        private string _text = null;
        public string Text
        {
            get
            {
                _text = string.Empty;
                foreach (var t in Characters)
                {
                    _text += t.Text;
                }
                return _text;
            } 
        }

        public double TextLength
        {
            get
            {
                double val = 0;
                foreach (var c in Characters)
                {
                    val += c.TextLength;
                }
                return val;
            }
        }
        public double TextHeight
        {
            get
            {
                double val = 0;
                foreach (var c in Characters)
                {
                    val += c.LineHeight;
                }
                return val;
            }
        }
        public double LineHeight
        {
            get
            {
                double val = 0;
                foreach (var c in Characters)
                {
                    val = c.LineHeight > val ? c.LineHeight : val;
                }
                return val;
            }
        }
        public double FontHeight
        {
            get
            {
                double val = 0;
                foreach (var c in Characters)
                {
                    val = c.FontHeight > val ? c.FontHeight : val;
                }
                return val;
            }
        }
    }

    internal struct PdfTextFormat
    {
        public IFontProvider FontProvider;
        //sometimes a font can have other fonts for certain characters. Key as the glyph id, Value is the font label.
        public Dictionary<byte, string> FontIDLabel;
        public Dictionary<byte, string> FontIdMap;
        public List<OpenTypeFont> UsedFonts;

        public ShapedText ShapedText;
        public string FontName;
        public int FontFamily;
        public FontSubFamily SubFamily;
        public double FontSize;
        public bool Bold;
        public bool Italic;
        public bool Strike;
        public bool SubScript;
        public bool SuperScript;
        public bool Underline;
        public ExcelUnderLineType UnderlineType;
        public string Text;
        public Color FontColor;
        public double TextLength;
        public double LineHeight;
        public double FontHeight;
        public Rect GlyphBox;
        public double characterOffset;
        public List<GlyphPosition> GlyphPositions;
        public string FullFontName
        {
            get
            {
                string subfam = " " + SubFamily.ToString();
                if (SubFamily == FontSubFamily.Regular)
                    subfam = "";
                else if (SubFamily == FontSubFamily.BoldItalic)
                    subfam = " Bold Italic";
                return FontName + subfam;
            }
        }

        //Compares stylings.
        public bool Equals(PdfTextFormat other)
        {
            if (!string.Equals(FontName, other.FontName))
                return false;

            if (FontFamily != other.FontFamily)
                return false;

            if (SubFamily != other.SubFamily)
                return false;

            if (FontSize != other.FontSize)
                return false;

            if (Bold != other.Bold ||
                Italic != other.Italic ||
                Strike != other.Strike ||
                SubScript != other.SubScript ||
                SuperScript != other.SuperScript ||
                Underline != other.Underline)
                return false;

            if (UnderlineType != other.UnderlineType)
                return false;

            if (!FontColor.Equals(other.FontColor))
                return false;

            return true;
        }
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
        public ExcelFillStyle PatternStyle = ExcelFillStyle.None;
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

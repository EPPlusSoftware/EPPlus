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
using System.Drawing;
using EPPlus.Export.Pdf.Enums;

namespace EPPlus.Export.Pdf.Layout
{
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
        internal const double OuterGridLine = 1d;
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
        public bool IsHeading = false;

        internal const double DoubleWidth = 0.75d;  // weight of each of the two lines
        internal const double DoubleOffset = 0.85d;  // offset from the gridline; also the corner miter amount

        public bool PerpAtStart = false;   // Top/Bottom: left end · Left/Right: bottom end
        public bool PerpAtEnd = false;   // Top/Bottom: right end · Left/Right: top end

        public bool ContAtStart = false;   // NEW: the border continues collinearly past the start vertex
        public bool ContAtEnd = false;   // NEW: the border continues collinearly past the end vertex
                                         //
                                         // (DoubleWidth = 0.75 and DoubleOffset = 0.85 are already present — keep them.)

        public PdfCellBorderData(LineType LineType)
        {
            this.LineType = LineType;
        }
    }
}

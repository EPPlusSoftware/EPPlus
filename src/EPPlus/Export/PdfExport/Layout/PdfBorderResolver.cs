using EPPlus.Export.Pdf.Helpers;
using EPPlus.Export.Pdf.Layout;
using OfficeOpenXml.Export.PdfExport.Data;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.PdfExport.Layout
{
    internal class PdfBorderResolver
    {
        internal static bool HasBorder(PdfCellStyle cellStyle)
        {
            if (cellStyle == null) return false;
            bool SideHas(bool suppress, ExcelBorderItem xf, ExcelDxfBorderItem dxf) =>
                !suppress && (xf.Style != ExcelBorderStyle.None || (dxf?.HasValue ?? false));
            return SideHas(cellStyle.SuppressTop, cellStyle.xfTop, cellStyle.dxfTop)
                || SideHas(cellStyle.SuppressBottom, cellStyle.xfBottom, cellStyle.dxfBottom)
                || SideHas(cellStyle.SuppressLeft, cellStyle.xfLeft, cellStyle.dxfLeft)
                || SideHas(cellStyle.SuppressRight, cellStyle.xfRight, cellStyle.dxfRight)
                || (cellStyle.Diagonal != null && cellStyle.Diagonal.Style != ExcelBorderStyle.None);
        }

        internal static bool HasDiagonalBorder(PdfCellStyle style) => style?.Diagonal != null && style.Diagonal.Style != ExcelBorderStyle.None;

        internal static void SetBorderStyle(PdfCellStyle style, PdfCellBorderLayout border, bool isBlackAndWhite)
        {
            var topStyle = style.xfTop.Style == ExcelBorderStyle.None ? ((style.dxfTop != null && style.dxfTop.HasValue) ? (ExcelBorderStyle)style.dxfTop.Style : ExcelBorderStyle.None) : style.xfTop.Style;
            var topColor = (style.xfTop.Style == ExcelBorderStyle.None && style.dxfTop != null) ? PdfColor.SetColorFromHex(style.dxfTop.Color.LookupColor(style.dxfTop)) : PdfColor.SetColorFromHex(style.xfTop.Color.LookupColor(style.xfTop));
            topColor = isBlackAndWhite ? Color.Black : topColor;

            var bottomStyle = style.xfBottom.Style == ExcelBorderStyle.None ? ((style.dxfBottom != null && style.dxfBottom.HasValue) ? (ExcelBorderStyle)style.dxfBottom.Style : ExcelBorderStyle.None) : style.xfBottom.Style;
            var bottomColor = (style.xfBottom.Style == ExcelBorderStyle.None && style.dxfBottom != null) ? PdfColor.SetColorFromHex(style.dxfBottom.Color.LookupColor(style.dxfBottom)) : PdfColor.SetColorFromHex(style.xfBottom.Color.LookupColor(style.xfBottom));
            bottomColor = isBlackAndWhite ? Color.Black : bottomColor;

            var leftStyle = style.xfLeft.Style == ExcelBorderStyle.None ? ((style.dxfLeft != null && style.dxfLeft.HasValue) ? (ExcelBorderStyle)style.dxfLeft.Style : ExcelBorderStyle.None) : style.xfLeft.Style;
            var leftColor = (style.xfLeft.Style == ExcelBorderStyle.None && style.dxfLeft != null) ? PdfColor.SetColorFromHex(style.dxfLeft.Color.LookupColor(style.dxfLeft)) : PdfColor.SetColorFromHex(style.xfLeft.Color.LookupColor(style.xfLeft));
            leftColor = isBlackAndWhite ? Color.Black : leftColor;

            var rightStyle = style.xfRight.Style == ExcelBorderStyle.None ? ((style.dxfRight != null && style.dxfRight.HasValue) ? (ExcelBorderStyle)style.dxfRight.Style : ExcelBorderStyle.None) : style.xfRight.Style;
            var rightColor = (style.xfRight.Style == ExcelBorderStyle.None && style.dxfRight != null) ? PdfColor.SetColorFromHex(style.dxfRight.Color.LookupColor(style.dxfRight)) : PdfColor.SetColorFromHex(style.xfRight.Color.LookupColor(style.xfRight));
            rightColor = isBlackAndWhite ? Color.Black : rightColor;

            var diagUpStyle = style.DiagonalUp ? style.Diagonal.Style : ExcelBorderStyle.None;
            var diagUpColor = style.DiagonalUp ? PdfColor.SetColorFromHex(style.Diagonal.Color.LookupColor(style.Diagonal)) : Color.Transparent;
            diagUpColor = isBlackAndWhite ? Color.Black : diagUpColor;

            var diagDownStyle = style.DiagonalDown ? style.Diagonal.Style : ExcelBorderStyle.None;
            var diagDownColor = style.DiagonalDown ? PdfColor.SetColorFromHex(style.Diagonal.Color.LookupColor(style.Diagonal)) : Color.Transparent;
            diagDownColor = isBlackAndWhite ? Color.Black : diagDownColor;

            if (style.SuppressTop) topStyle = ExcelBorderStyle.None;
            if (style.SuppressBottom) bottomStyle = ExcelBorderStyle.None;
            if (style.SuppressLeft) leftStyle = ExcelBorderStyle.None;
            if (style.SuppressRight) rightStyle = ExcelBorderStyle.None;

            border.SetStyle((EPPlus.Export.Pdf.Enums.ExcelBorderStyle)topStyle, topColor,
                            (EPPlus.Export.Pdf.Enums.ExcelBorderStyle)bottomStyle, bottomColor,
                            (EPPlus.Export.Pdf.Enums.ExcelBorderStyle)leftStyle, leftColor,
                            (EPPlus.Export.Pdf.Enums.ExcelBorderStyle)rightStyle, rightColor,
                            (EPPlus.Export.Pdf.Enums.ExcelBorderStyle)diagUpStyle, diagUpColor,
                            (EPPlus.Export.Pdf.Enums.ExcelBorderStyle)diagDownStyle, diagDownColor);
        }

        private static PdfCell CellAt(Page page, int row, int col)
        {
            if (row < page.FromRow || row > page.ToRow || col < page.FromColumn || col > page.ToColumn) return null;
            return page.Map[row, col];
        }

        private static bool VBorderAt(Page page, int row, int col)
            => CellHasLeftBorder(CellAt(page, row, col)) || CellHasRightBorder(CellAt(page, row, col - 1));

        private static bool HBorderBelow(Page page, int row, int col)
            => CellHasBottomBorder(CellAt(page, row, col)) || CellHasTopBorder(CellAt(page, row + 1, col));

        private static bool HBorderAbove(Page page, int row, int col)
            => CellHasTopBorder(CellAt(page, row, col)) || CellHasBottomBorder(CellAt(page, row - 1, col));

        internal static void SetDoubleBorderMiterFlags(Page page, int row, int col, PdfCellBorderLayout border)
        {
            var b = border.BorderData;
            b.Top.PerpAtStart = VBorderAt(page, row, col) || VBorderAt(page, row - 1, col);
            b.Top.PerpAtEnd = VBorderAt(page, row, col + 1) || VBorderAt(page, row - 1, col + 1);
            b.Bottom.PerpAtStart = VBorderAt(page, row, col) || VBorderAt(page, row + 1, col);
            b.Bottom.PerpAtEnd = VBorderAt(page, row, col + 1) || VBorderAt(page, row + 1, col + 1);
            b.Left.PerpAtStart = HBorderBelow(page, row, col) || HBorderBelow(page, row, col - 1);
            b.Left.PerpAtEnd = HBorderAbove(page, row, col) || HBorderAbove(page, row, col - 1);
            b.Right.PerpAtStart = HBorderBelow(page, row, col) || HBorderBelow(page, row, col + 1);
            b.Right.PerpAtEnd = HBorderAbove(page, row, col) || HBorderAbove(page, row, col + 1);
            b.Top.NeighborDouble = IsDoubleBottom(CellAt(page, row - 1, col));
            b.Bottom.NeighborDouble = IsDoubleTop(CellAt(page, row + 1, col));
            b.Left.NeighborDouble = IsDoubleRight(CellAt(page, row, col - 1));
            b.Right.NeighborDouble = IsDoubleLeft(CellAt(page, row, col + 1));
            var nAbove = CellAt(page, row - 1, col);
            var nBelow = CellAt(page, row + 1, col);
            var nLeft = CellAt(page, row, col - 1);
            var nRight = CellAt(page, row, col + 1);
            b.Top.NeighborDiagAtStart = HasDiagUp(nAbove);
            b.Top.NeighborDiagAtEnd = HasDiagDown(nAbove);
            b.Bottom.NeighborDiagAtStart = HasDiagDown(nBelow);
            b.Bottom.NeighborDiagAtEnd = HasDiagUp(nBelow);
            b.Left.NeighborDiagAtStart = HasDiagDown(nLeft);
            b.Left.NeighborDiagAtEnd = HasDiagUp(nLeft);
            b.Right.NeighborDiagAtStart = HasDiagUp(nRight);
            b.Right.NeighborDiagAtEnd = HasDiagDown(nRight);
            var dTL = CellAt(page, row - 1, col - 1);
            var dTR = CellAt(page, row - 1, col + 1);
            var dBL = CellAt(page, row + 1, col - 1);
            var dBR = CellAt(page, row + 1, col + 1);
            bool cutTL = CellHasBottomBorder(dTL) && CellHasRightBorder(dTL);
            bool cutTR = CellHasBottomBorder(dTR) && CellHasLeftBorder(dTR);
            bool cutBL = CellHasTopBorder(dBL) && CellHasRightBorder(dBL);
            bool cutBR = CellHasTopBorder(dBR) && CellHasLeftBorder(dBR);
            b.Top.CutOuterAtStart = cutTL; b.Top.CutOuterAtEnd = cutTR;
            b.Bottom.CutOuterAtStart = cutBL; b.Bottom.CutOuterAtEnd = cutBR;
            b.Left.CutOuterAtStart = cutBL; b.Left.CutOuterAtEnd = cutTL;
            b.Right.CutOuterAtStart = cutBR; b.Right.CutOuterAtEnd = cutTR;
        }

        private static bool HasDiagUp(PdfCell cell)
        {
            var cs = cell?.CellStyle; if (cs == null) return false;
            return cs.DiagonalUp && cs.Diagonal != null && cs.Diagonal.Style != ExcelBorderStyle.None;
        }
        private static bool HasDiagDown(PdfCell cell)
        {
            var cs = cell?.CellStyle; if (cs == null) return false;
            return cs.DiagonalDown && cs.Diagonal != null && cs.Diagonal.Style != ExcelBorderStyle.None;
        }

        private static bool IsDoubleTop(PdfCell cell)
        {
            var cs = cell?.CellStyle; if (cs == null) return false;
            var s = cs.xfTop.Style != ExcelBorderStyle.None ? cs.xfTop.Style
                  : ((cs.dxfTop != null && cs.dxfTop.HasValue) ? (ExcelBorderStyle)cs.dxfTop.Style : ExcelBorderStyle.None);
            return s == ExcelBorderStyle.Double;
        }
        private static bool IsDoubleBottom(PdfCell cell)
        {
            var cs = cell?.CellStyle; if (cs == null) return false;
            var s = cs.xfBottom.Style != ExcelBorderStyle.None ? cs.xfBottom.Style
                  : ((cs.dxfBottom != null && cs.dxfBottom.HasValue) ? (ExcelBorderStyle)cs.dxfBottom.Style : ExcelBorderStyle.None);
            return s == ExcelBorderStyle.Double;
        }
        private static bool IsDoubleLeft(PdfCell cell)
        {
            var cs = cell?.CellStyle; if (cs == null) return false;
            var s = cs.xfLeft.Style != ExcelBorderStyle.None ? cs.xfLeft.Style
                  : ((cs.dxfLeft != null && cs.dxfLeft.HasValue) ? (ExcelBorderStyle)cs.dxfLeft.Style : ExcelBorderStyle.None);
            return s == ExcelBorderStyle.Double;
        }
        private static bool IsDoubleRight(PdfCell cell)
        {
            var cs = cell?.CellStyle; if (cs == null) return false;
            var s = cs.xfRight.Style != ExcelBorderStyle.None ? cs.xfRight.Style
                  : ((cs.dxfRight != null && cs.dxfRight.HasValue) ? (ExcelBorderStyle)cs.dxfRight.Style : ExcelBorderStyle.None);
            return s == ExcelBorderStyle.Double;
        }

        private static bool CellHasRightBorder(PdfCell cell)
        {
            var cs = cell?.CellStyle; if (cs == null) return false;
            return (cs.xfRight != null && cs.xfRight.Style != ExcelBorderStyle.None) || (cs.dxfRight?.HasValue ?? false);
        }

        private static bool CellHasLeftBorder(PdfCell cell)
        {
            var cs = cell?.CellStyle; if (cs == null) return false;
            return (cs.xfLeft != null && cs.xfLeft.Style != ExcelBorderStyle.None) || (cs.dxfLeft?.HasValue ?? false);
        }

        private static bool CellHasTopBorder(PdfCell cell)
        {
            var cs = cell?.CellStyle; if (cs == null) return false;
            return (cs.xfTop != null && cs.xfTop.Style != ExcelBorderStyle.None) || (cs.dxfTop?.HasValue ?? false);
        }

        private static bool CellHasBottomBorder(PdfCell cell)
        {
            var cs = cell?.CellStyle; if (cs == null) return false;
            return (cs.xfBottom != null && cs.xfBottom.Style != ExcelBorderStyle.None) || (cs.dxfBottom?.HasValue ?? false);
        }
    }
}

using EPPlus.Export.Pdf.Layout;
using EPPlus.Export.Pdf.Resources;
using EPPlus.Export.Pdf.Settings;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.Integration.DataHolders;
using OfficeOpenXml.Export.PdfExport.Data;
using OfficeOpenXml.Export.PdfExport.TextShaping;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.PdfExport.Layout
{
    internal class PdfHeadingLayout
    {
        internal static Pages PrecomputePrintTitleCells(PdfPageSettings pageSettings, PdfWorksheet pdfSheet, PdfRange range, Pages pdfPages)
        {
            for (int i = 0; i < pdfPages.Page.Length; i++)
                pdfPages.Page[i] = PrecomputePagePrintTitleCells(pageSettings, pdfSheet, range, pdfPages.Page[i]);
            return pdfPages;
        }

        private static Page PrecomputePagePrintTitleCells(PdfPageSettings pageSettings, PdfWorksheet pdfSheet, PdfRange range, Page page)
        {
            page.PrintTitleCells = new List<PrintTitleCellDraw>();
            page.PrintTitleBorders = new List<PrintTitleCellDraw>();

            bool topBand = page.PrintTitleHeight > 0d && pdfSheet.PrintTitleRowFrom >= 0 && page.FromRow > pdfSheet.PrintTitleRowTo;
            bool leftBand = page.PrintTitleWidth > 0d && pdfSheet.PrintTitleColFrom >= 0 && page.FromColumn > pdfSheet.PrintTitleColTo;
            if (!topBand && !leftBand)
                return page;
            var contentColX = new Dictionary<int, double>();
            double cx = PdfLayout.GetOriginX(pageSettings, page) + page.HeadingWidth + page.PrintTitleWidth;
            for (int c = page.FromColumn; c <= page.ToColumn; c++)
            {
                contentColX[c] = cx;
                cx += PdfRangeGeometry.RangeColWidth(range, page.FromRow, c);
            }
            var contentRowY = new Dictionary<int, double>();
            double cy = PdfLayout.GetOriginY(pageSettings, page) - page.HeadingHeight - page.PrintTitleHeight;
            for (int r = page.FromRow; r <= page.ToRow; r++)
            {
                contentRowY[r] = cy;
                cy -= PdfRangeGeometry.RangeRowHeight(range, r);
            }
            var titleColX = new Dictionary<int, double>();
            if (leftBand)
            {
                double tx = PdfLayout.GetOriginX(pageSettings, page) + page.HeadingWidth;
                for (int c = pdfSheet.PrintTitleColFrom; c <= pdfSheet.PrintTitleColTo; c++)
                {
                    titleColX[c] = tx;
                    tx += PdfRangeGeometry.RangeColWidth(range, page.FromRow, c);
                }
            }
            var titleRowY = new Dictionary<int, double>();
            if (topBand)
            {
                double ty = PdfLayout.GetOriginY(pageSettings, page) - page.HeadingHeight;
                for (int r = pdfSheet.PrintTitleRowFrom; r <= pdfSheet.PrintTitleRowTo; r++)
                {
                    titleRowY[r] = ty;
                    ty -= PdfRangeGeometry.RangeRowHeight(range, r);
                }
            }
            if (topBand)
                PdfBandRegion.ProcessBandRegionCells(page, range, pdfSheet.PrintTitleRowFrom, pdfSheet.PrintTitleRowTo, page.FromColumn, page.ToColumn, contentColX, titleRowY);
            if (leftBand)
                PdfBandRegion.ProcessBandRegionCells(page, range, page.FromRow, page.ToRow, pdfSheet.PrintTitleColFrom, pdfSheet.PrintTitleColTo, titleColX, contentRowY);
            if (topBand && leftBand)
                PdfBandRegion.ProcessBandRegionCells(page, range, pdfSheet.PrintTitleRowFrom, pdfSheet.PrintTitleRowTo, pdfSheet.PrintTitleColFrom, pdfSheet.PrintTitleColTo, titleColX, titleRowY);

            page.PrintTitleGridLines = new List<GridLine>();
            if (pageSettings.ShowGridLines)
            {
                if (topBand)
                    PdfBandRegion.EmitBandGrid(page.PrintTitleGridLines, range, pdfSheet.PrintTitleRowFrom, pdfSheet.PrintTitleRowTo, page.FromColumn, page.ToColumn,
                        PdfBandRegion.BandEdgesX(contentColX, page.FromColumn, page.ToColumn, range, page.FromRow),
                        PdfBandRegion.BandEdgesY(titleRowY, pdfSheet.PrintTitleRowFrom, pdfSheet.PrintTitleRowTo, range));
                if (leftBand)
                    PdfBandRegion.EmitBandGrid(page.PrintTitleGridLines, range, page.FromRow, page.ToRow, pdfSheet.PrintTitleColFrom, pdfSheet.PrintTitleColTo,
                       PdfBandRegion.BandEdgesX(titleColX, pdfSheet.PrintTitleColFrom, pdfSheet.PrintTitleColTo, range, page.FromRow),
                       PdfBandRegion.BandEdgesY(contentRowY, page.FromRow, page.ToRow, range));
                if (topBand && leftBand)
                    PdfBandRegion.EmitBandGrid(page.PrintTitleGridLines, range, pdfSheet.PrintTitleRowFrom, pdfSheet.PrintTitleRowTo, pdfSheet.PrintTitleColFrom, pdfSheet.PrintTitleColTo,
                        PdfBandRegion.BandEdgesX(titleColX, pdfSheet.PrintTitleColFrom, pdfSheet.PrintTitleColTo, range, page.FromRow),
                        PdfBandRegion.BandEdgesY(titleRowY, pdfSheet.PrintTitleRowFrom, pdfSheet.PrintTitleRowTo, range));
            }
            page.PrintTitleHeadings = new List<PrintTitleHeadingDraw>();
            if (pageSettings.ShowHeadings)
            {
                if (topBand)
                {
                    for (int r = pdfSheet.PrintTitleRowFrom; r <= pdfSheet.PrintTitleRowTo; r++)
                    {
                        double h = PdfRangeGeometry.RangeRowHeight(range, r);
                        if (h <= 0d) continue;
                        page.PrintTitleHeadings.Add(new PrintTitleHeadingDraw
                        {
                            IsRow = true,
                            Index = r,
                            X = PdfLayout.GetOriginX(pageSettings, page),
                            Y = titleRowY[r],
                            Width = page.HeadingWidth,
                            Height = h
                        });
                    }
                }
                if (leftBand)
                {
                    for (int c = pdfSheet.PrintTitleColFrom; c <= pdfSheet.PrintTitleColTo; c++)
                    {
                        double w = PdfRangeGeometry.RangeColWidth(range, page.FromRow, c);
                        if (w <= 0d) continue;
                        page.PrintTitleHeadings.Add(new PrintTitleHeadingDraw
                        {
                            IsRow = false,
                            Index = c,
                            X = titleColX[c],
                            Y = PdfLayout.GetOriginY(pageSettings, page),
                            Width = w,
                            Height = page.HeadingHeight
                        });
                    }
                }
            }
            if (topBand)
            {
                AddIncomingSpill(page, range, pdfSheet.PrintTitleRowFrom, pdfSheet.PrintTitleRowTo, page.FromColumn, page.ToColumn,
                                    GetOriginX(pageSettings, page) + page.HeadingWidth + page.PrintTitleWidth,
                                    GetOriginY(pageSettings, page) - page.HeadingHeight,
                                    isPrintTitle: true);
            }
            if (leftBand)
            {
                AddIncomingSpill(page, range, page.FromRow, page.ToRow,
                                    pdfSheet.PrintTitleColFrom, pdfSheet.PrintTitleColTo,
                                    GetOriginX(pageSettings, page) + page.HeadingWidth,
                                    GetOriginY(pageSettings, page) - page.HeadingHeight - page.PrintTitleHeight,
                                    isPrintTitle: true);
            }
            if (topBand && leftBand)
            {
                AddIncomingSpill(page, range, pdfSheet.PrintTitleRowFrom, pdfSheet.PrintTitleRowTo,
                                    pdfSheet.PrintTitleColFrom, pdfSheet.PrintTitleColTo,
                                    GetOriginX(pageSettings, page) + page.HeadingWidth,
                                    GetOriginY(pageSettings, page) - page.HeadingHeight,
                                    isPrintTitle: true);
            }
            return page;
        }

        internal static void AddPrintTitleCells(PdfPageSettings pageSettings, PdfDictionaries dictionaries, Page page, PdfPageLayout pageLayout)
        {
            if (page.PrintTitleCells == null) return;
            pageLayout.PrintTitleGridLines = page.PrintTitleGridLines ?? new List<GridLine>();
            foreach (var t in page.PrintTitleCells)
            {
                var map = t.Cell;
                var fill = new PdfCellLayout(t.X, t.Y, t.Width, t.Height);
                SetFill(dictionaries, map.CellStyle, map.Text, fill);
                fill.Name = map.Name;
                fill.IsPrintTitle = true;
                fill.UpdateShadingPositionMatrix(pageSettings);
                pageLayout.AddChild(fill);

                if (map.TextLines != null && map.TextLines.Count > 0)
                {
                    var info = new MergedCellDrawInfo { X = t.X, Y = t.Y, Width = t.Width, Height = t.Height };
                    var text = new PdfCellContentLayout(pageSettings, dictionaries, map, t.X, t.Y, t.Width, t.Height);
                    text.Name = map.Name;
                    text.IsPrintTitle = true;
                    text.GidsAndCharMap(dictionaries);
                    text.SetupClipping(t.ClipX, t.ClipY, t.ClipWidth, t.ClipHeight);
                    pageLayout.AddChild(text);
                }
                if (!map.Merged && HasBorder(map.CellStyle))
                {
                    var border = new PdfCellBorderLayout(false, MergedCellCorners.All, new MergedCellDrawInfo(), t.X, t.Y, t.Width, t.Height);
                    SetBorderStyle(map.CellStyle, border, pageSettings.BlackAndWhite);
                    border.IsPrintTitle = true;
                    border.Name = "PrintTitleBorder_" + map.Name;
                    pageLayout.AddChild(border);
                }
                foreach (var b in page.PrintTitleBorders)
                {
                    var sub = b.Cell;
                    if (!HasBorder(sub.CellStyle)) continue;
                    var border = new PdfCellBorderLayout(false, MergedCellCorners.All, new MergedCellDrawInfo(), b.X, b.Y, b.Width, b.Height);
                    SetBorderStyle(sub.CellStyle, border, pageSettings.BlackAndWhite);
                    border.IsPrintTitle = true;
                    border.Name = "PrintTitleMergeBorder_" + sub.Name;
                    pageLayout.AddChild(border);
                }
            }
        }

        internal static void AddHeadingCell(PdfPageSettings pageSettings, PdfDictionaries dictionaries, PdfPageLayout pageLayout, PdfCellStyle headingStyle, string label, double x, double y, double width, double height, string fontName, float fontSize, string namePrefix)
        {
            if (width == 0d || height == 0d) return;
            var fill = new PdfCellLayout(x, y, width, height);
            PdfFillResolver.SetFill(dictionaries, headingStyle, label, fill);
            fill.Name = namePrefix;
            fill.IsHeading = true;
            fill.UpdateShadingPositionMatrix(pageSettings);
            pageLayout.AddChild(fill);
            var cell = CreateHeadingPdfCell(pageSettings, dictionaries, label, ExcelHorizontalAlignment.Center, width, height, fontName, fontSize);
            if (cell.TextLines != null && cell.TextLines.Count > 0)
            {
                var info = new MergedCellDrawInfo { X = x, Y = y, Width = width, Height = height };
                var text = new PdfCellContentLayout(pageSettings, dictionaries, cell, x, y, width, height);
                text.Name = namePrefix + "_Text";
                text.IsHeading = true;
                text.GidsAndCharMap(dictionaries);
                pageLayout.AddChild(text);
            }
            if (!pageSettings.Draft)
                AddHeadingCellBorder(pageLayout, x, y, width, height);
        }

        internal static void AddHeadingCells(PdfPageSettings pageSettings, PdfDictionaries dictionaries, Page page, PdfPageLayout pageLayout, double contentStartX, double contentStartY, double headingWidth, double headingHeight, string fontName, float fontSize, ExcelFill fill)
        {
            var headingStyle = new PdfCellStyle();
            headingStyle.xfFill = fill;
            var cornerFill = new PdfCellLayout(GetOriginX(pageSettings, page), GetOriginY(pageSettings, page), headingWidth, headingHeight);
            PdfFillResolver.SetFill(dictionaries, headingStyle, "", cornerFill);
            cornerFill.Name = "Heading_Corner";
            cornerFill.UpdateShadingPositionMatrix(pageSettings);
            pageLayout.AddChild(cornerFill);
            double x = contentStartX;
            for (int col = page.FromColumn; col <= page.ToColumn; col++)
            {
                double colWidth = page.Map[page.FromRow, col]?.ColumnWidth ?? 0d;
                if (colWidth == 0d) { x += colWidth; continue; }
                string colLetter = ExcelCellBase.GetColumnLetter(col);
                AddHeadingCell(pageSettings, dictionaries, pageLayout, headingStyle, colLetter,
                    x, GetOriginY(pageSettings, page), colWidth, headingHeight, fontName, fontSize, "Heading_Col_" + colLetter);
                x += colWidth;
            }
            if (!pageSettings.Draft)
            {
                double y = contentStartY;
                for (int row = page.FromRow; row <= page.ToRow; row++)
                {
                    double rowHeight = page.RowHeights[row - page.FromRow];
                    if (rowHeight == 0d) { y -= rowHeight; continue; }
                    string rowNum = row.ToString();
                    AddHeadingCell(pageSettings, dictionaries, pageLayout, headingStyle, rowNum,
                        GetOriginX(pageSettings, page), y, headingWidth, rowHeight, fontName, fontSize, "Heading_Row_" + rowNum);
                    y -= rowHeight;
                }
            }
        }

        internal static void AddPrintTitleHeadings(PdfPageSettings pageSettings, PdfDictionaries dictionaries, Page page, PdfPageLayout pageLayout, string fontName, float fontSize, ExcelFill fill)
        {
            if (page.PrintTitleHeadings == null || page.PrintTitleHeadings.Count == 0) return;

            var headingStyle = new PdfCellStyle();
            headingStyle.xfFill = fill;

            foreach (var h in page.PrintTitleHeadings)
            {
                string label = h.IsRow ? h.Index.ToString() : ExcelCellBase.GetColumnLetter(h.Index);
                string namePrefix = (h.IsRow ? "Heading_Row_" : "Heading_Col_") + label;
                AddHeadingCell(pageSettings, dictionaries, pageLayout, headingStyle, label, h.X, h.Y, h.Width, h.Height, fontName, fontSize, namePrefix);
            }
        }

        private static void AddHeadingCellBorder(PdfPageLayout pageLayout, double x, double y, double width, double height)
        {
            double right = x + width;
            double bottom = y - height;
            pageLayout.BorderLines.Add(new GridLine(x, y, right, y));
            pageLayout.BorderLines.Add(new GridLine(x, bottom, right, bottom));
            pageLayout.BorderLines.Add(new GridLine(x, y, x, bottom));
            pageLayout.BorderLines.Add(new GridLine(right, y, right, bottom));
        }

        private static PdfCell CreateHeadingPdfCell(PdfPageSettings pageSettings, PdfDictionaries dictionaries, string text, ExcelHorizontalAlignment hAlign, double width, double height, string fontName, float fontSize)
        {
            var cell = new PdfCell();
            cell.ColumnWidth = width;
            cell.Width = width;
            cell.Height = height;
            cell.Text = text;
            cell.CellStyle = new PdfCellStyle();
            cell.ContentAligmnet = new PdfCellAlignmentData();
            cell.ContentAligmnet.HorizontalAlignment = (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)hAlign;
            cell.ContentAligmnet.VerticalAlignment = (EPPlus.Export.Pdf.Enums.ExcelVerticalAlignment)ExcelVerticalAlignment.Bottom;
            cell.ContentAligmnet.WrapText = false;
            if (!string.IsNullOrEmpty(text))
            {
                var tf = new TextFragment();
                tf.Font = new RichTextFormatSimple();
                tf.Font.Family = fontName;
                tf.Font.SubFamily = FontSubFamily.Regular;
                tf.Font.Size = fontSize;
                tf.Text = text;
                tf.RichTextOptions.Bold = false;
                tf.RichTextOptions.Italic = false;
                tf.RichTextOptions.UnderlineType = 12;
                tf.RichTextOptions.StrikeType = 1;
                cell.TextFragments = new List<TextFragment> { tf };
                PdfTextShaper.ShapeText(pageSettings, dictionaries, cell);
            }
            return cell;
        }
    }
}

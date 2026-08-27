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
using EPPlus.Export.Pdf.Helpers;
using EPPlus.Export.Pdf.Layout;
using EPPlus.Export.Pdf.Resources;
using EPPlus.Export.Pdf.Settings;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.Integration.DataHolders;
using EPPlus.Graphics;
using OfficeOpenXml.Export.PdfExport.Data;
using OfficeOpenXml.Export.PdfExport.TextShaping;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace OfficeOpenXml.Export.PdfExport.Layout
{ 
    internal struct PrintTitleCellDraw
    {
        public PdfCell Cell;
        public double X;
        public double Y;
        public double Width;
        public double Height;
        public double ClipX;       // text clip (spill bound)
        public double ClipY;
        public double ClipWidth;
        public double ClipHeight;
    }

    internal struct PrintTitleHeadingDraw
    {
        public bool IsRow;     // true = row-number heading (left strip); false = column-letter (top strip)
        public int Index;      // original absolute row/column index — the label source
        public double X;
        public double Y;
        public double Width;
        public double Height;
    }

    internal struct SpillCellDraw
    {
        public PdfCell Cell;
        public double X, Y, Width, Height;                  // source cell's true position (off-window)
        public double ClipX, ClipY, ClipWidth, ClipHeight;  // visible slice on this page
        public bool IsPrintTitle;
    }

    internal class PdfLayout
    {
        private const double rowHeadingWith1CharWidth = 18d;

        public static Transform GetLayout(PdfPageSettings[] sheetSettings, PdfDictionaries dictionaries, PdfWorksheet[] pdfSheets)
        {
            var PagesCollection = GetPages(sheetSettings, pdfSheets);
            // Page numbering is document-global; sheet 1 supplies it, as before.
            var Catalog = GetCatalog(sheetSettings[0].FirstPageNumber, dictionaries, PagesCollection);
            return Catalog;
        }

        internal static Transform GetCatalog(int firstPageNumber, PdfDictionaries dictionaries, List<Pages> pdfPages)
        {
            Transform Catalog = new Transform(0d, 0d, 0d, 0d);
            int totalPages = GetTotalPages(pdfPages);
            for (int i = 0; i < pdfPages.Count; i++)
            {
                var pageSettings = pdfPages[i].Settings;


                var pages = pdfPages[i].Page;
                int pageNumber = pageSettings.FirstPageNumber;
                for (int j = 0; j < pages.Length; j++)
                {
                    var page = pages[j];
                    PdfPageLayout pageLayout = new PdfPageLayout(0d, 0d, 0d, 0d);
                    pageLayout.Settings = pageSettings;             
                    pageLayout.isCommentsPage = pdfPages[i].IsCommentsPage;
                    pageLayout.HeadingWidth = page.HeadingWidth;
                    pageLayout.HeadingHeight = page.HeadingHeight;
                    pageLayout.PrintTitleWidth = page.PrintTitleWidth;
                    pageLayout.PrintTitleHeight = page.PrintTitleHeight;
                    var drawnMergedCells = new HashSet<string>();
                    double contentStartX = pageSettings.ContentBounds.Left + page.HeadingWidth + page.PrintTitleWidth;
                    double contentStartY = pageSettings.ContentBounds.Top - page.HeadingHeight - page.PrintTitleHeight;
                    if (pageSettings.ShowHeadings && !pdfPages[i].IsCommentsPage)
                    {
                        AddHeadingCells(pageSettings, dictionaries, page, pageLayout, contentStartX, contentStartY, page.HeadingWidth, page.HeadingHeight, pdfPages[i].HeadingFontName, pdfPages[i].HeadingFontSize, pdfPages[i].HeadingFill);
                        AddPrintTitleHeadings(pageSettings, dictionaries, page, pageLayout, pdfPages[i].HeadingFontName, pdfPages[i].HeadingFontSize, pdfPages[i].HeadingFill);
                        AddSpillCells(pageSettings, dictionaries, page, pageLayout);
                    }
                    AddPrintTitleCells(pageSettings, dictionaries, page, pageLayout);
                    double y = contentStartY;
                    double x = contentStartX;
                    for (int row = pages[j].FromRow; row <= pages[j].ToRow; row++)
                    {
                        double rowHeight = pages[j].RowHeights[row - pages[j].FromRow];
                        for (int col = pages[j].FromColumn; col <= pages[j].ToColumn; col++)
                        {
                            var map = pages[j].Map[row, col];
                            MergedCellDrawInfo info = new MergedCellDrawInfo();
                            //Merged Cell
                            if (map.Merged)
                            {
                                string key = map.MergedAddress.Address;
                                if (!drawnMergedCells.Contains(key) &&
                                    pages[j].MergedCells.TryGetValue(key, out info))
                                {
                                    //Fill
                                    var cellStyle = map.Main?.CellStyle ?? map.CellStyle;
                                    var fill = new PdfCellLayout(info.X, info.Y, info.Width, info.Height);
                                    SetFill(dictionaries, cellStyle, map.Text, fill);
                                    fill.Name = map.Name;
                                    fill.UpdateShadingPositionMatrix(pageSettings);
                                    pageLayout.AddChild(fill);
                                    //Text
                                    var sourceMap = (map.TextLines != null && map.TextLines.Count > 0) ? map : (map.Main != null && map.Main.TextLines != null && map.Main.TextLines.Count > 0) ? map.Main : null;
                                    if (sourceMap != null)
                                    {
                                        var text = new PdfCellContentLayout(pageSettings, dictionaries, sourceMap, info, info.X, info.Y, info.Width, info.Height);
                                        text.Name = map.Name;
                                        text.GidsAndCharMap(dictionaries);
                                        text.SetupClipping(info.X, info.Y, info.Width, info.Height);
                                        pageLayout.AddChild(text);
                                    }
                                    if (map.Main != null) // map.Main != null → this is NOT the top-left cell
                                    {
                                        var mergeMainStyle = map.Main.CellStyle;
                                        if (HasDiagonalBorder(mergeMainStyle))
                                        {
                                            var diagBorder = new PdfCellBorderLayout(
                                                isMerged: false,            // use X/Y/W/H path in renderer, not info.*
                                                corners: MergedCellCorners.All,
                                                info: info,
                                                x: info.X,           // virtual full-merge top-left X
                                                y: info.Y,           // virtual full-merge top Y
                                                width: info.Width,       // full merge width
                                                height: info.Height);     // full merge height
                                            SetBorderStyle(mergeMainStyle, diagBorder);
                                            diagBorder.Name = map.Name;
                                            // Suppress edge borders — this layout exists only for the diagonal
                                            diagBorder.BorderData.Top.BorderStyle = (EPPlus.Export.Pdf.Enums.ExcelBorderStyle)ExcelBorderStyle.None;
                                            diagBorder.BorderData.Bottom.BorderStyle = (EPPlus.Export.Pdf.Enums.ExcelBorderStyle)ExcelBorderStyle.None;
                                            diagBorder.BorderData.Left.BorderStyle = (EPPlus.Export.Pdf.Enums.ExcelBorderStyle)ExcelBorderStyle.None;
                                            diagBorder.BorderData.Right.BorderStyle = (EPPlus.Export.Pdf.Enums.ExcelBorderStyle)ExcelBorderStyle.None;
                                            pageLayout.AddChild(diagBorder);
                                        }
                                    }
                                    drawnMergedCells.Add(key);
                                }
                            }
                            else
                            {
                                //Fill
                                var fill = new PdfCellLayout(x, y, map.ColumnWidth, rowHeight);
                                SetFill(dictionaries, map.CellStyle, map.Text, fill);
                                fill.UpdateShadingPositionMatrix(pageSettings);
                                fill.Name = map.Name;
                                pageLayout.AddChild(fill);
                                //Text
                                if (map.TextLines != null && map.TextLines.Count > 0)
                                {
                                    var text = new PdfCellContentLayout(pageSettings, dictionaries, map, info, x, y, map.ColumnWidth, rowHeight);
                                    text.Name = map.Name;
                                    text.GidsAndCharMap(dictionaries);
                                    if (NeedsClipping(map, pages[j], row, col))
                                        text.SetupClipping(x, y, map.ColumnWidth, rowHeight);
                                    pageLayout.AddChild(text);
                                }
                            }
                            //Border
                            var borderStyle = (map.Merged && map.Main != null) ? map.Main.CellStyle : map.CellStyle;
                            if (HasBorder(map.CellStyle))
                            {
                                var border = new PdfCellBorderLayout(map.Merged, GetCorners(map.MergedAddress, row, col), info, x, y, map.ColumnWidth, rowHeight);
                                SetBorderStyle(map.CellStyle, border);
                                border.Name = map.Name;
                                if (map.Merged && map.MergedAddress != null)
                                {
                                    var addr = map.MergedAddress;
                                    if (row != addr.Start.Row) border.BorderData.Top.BorderStyle = (EPPlus.Export.Pdf.Enums.ExcelBorderStyle)ExcelBorderStyle.None;
                                    if (row != addr.End.Row) border.BorderData.Bottom.BorderStyle = (EPPlus.Export.Pdf.Enums.ExcelBorderStyle)ExcelBorderStyle.None;
                                    if (col != addr.Start.Column) border.BorderData.Left.BorderStyle = (EPPlus.Export.Pdf.Enums.ExcelBorderStyle)ExcelBorderStyle.None;
                                    if (col != addr.End.Column) border.BorderData.Right.BorderStyle = (EPPlus.Export.Pdf.Enums.ExcelBorderStyle)ExcelBorderStyle.None;
                                }
                                pageLayout.AddChild(border);
                            }
                            x += map.ColumnWidth;
                        }
                        y -= rowHeight;
                        x = contentStartX; //pageSettings.ContentBounds.Left;
                    }
                    if (page.HeaderFooters != null)
                    {
                        bool isVeryFirstPage = (i == 0 && j == 0);
                        var hfType = isVeryFirstPage ? HeaderFooterType.First : (pageNumber % 2 == 0 ? HeaderFooterType.Even : HeaderFooterType.Odd);
                        var leftH = page.HeaderFooters.Get(hfType, HeaderFooterSection.Header, HeaderFooterAlignment.Left);
                        if (leftH != null)
                        {
                            SubstitutePageNumbers(pageSettings, dictionaries, leftH, pageNumber, totalPages);
                            var ascent = leftH.Content.TextLines[0].LargestAscent;
                            var hfx = pageSettings.Margins.LeftPu;
                            var hfy = pageSettings.PageSize.HeightPu - pageSettings.Margins.HeaderPu - ascent;
                            var text = new PdfCellContentLayout(pageSettings, dictionaries, leftH, hfx, hfy, 0, 0);
                            text.Name = "LeftHeader";
                            text.IsHeaderFooter = true;
                            text.GidsAndCharMap(dictionaries);
                            pageLayout.AddChild(text);
                        }
                        var centerH = page.HeaderFooters.Get(hfType, HeaderFooterSection.Header, HeaderFooterAlignment.Center);
                        if (centerH != null)
                        {
                            SubstitutePageNumbers(pageSettings, dictionaries, centerH, pageNumber, totalPages);
                            var ascent = centerH.Content.TextLines[0].LargestAscent;
                            var hfx = pageSettings.Margins.LeftPu;
                            var hfy = pageSettings.PageSize.HeightPu - pageSettings.Margins.HeaderPu - ascent;
                            var hfWidth = pageSettings.PageSize.WidthPu - pageSettings.Margins.LeftPu - pageSettings.Margins.RightPu;
                            var text = new PdfCellContentLayout(pageSettings, dictionaries, centerH, hfx, hfy, hfWidth, 0);
                            text.Name = "CenterHeader";
                            text.IsHeaderFooter = true;
                            text.GidsAndCharMap(dictionaries);
                            pageLayout.AddChild(text);
                        }
                        var rightH = page.HeaderFooters.Get(hfType, HeaderFooterSection.Header, HeaderFooterAlignment.Right);
                        if (rightH != null)
                        {
                            SubstitutePageNumbers(pageSettings, dictionaries, rightH, pageNumber, totalPages);
                            var ascent = rightH.Content.TextLines[0].LargestAscent;
                            var hfx = pageSettings.PageSize.WidthPu - pageSettings.Margins.RightPu;
                            var hfy = pageSettings.PageSize.HeightPu - pageSettings.Margins.HeaderPu - ascent;
                            var text = new PdfCellContentLayout(pageSettings, dictionaries, rightH, hfx, hfy, 0, 0);
                            text.Name = "RightHeader";
                            text.IsHeaderFooter = true;
                            text.GidsAndCharMap(dictionaries);
                            pageLayout.AddChild(text);
                        }
                        var leftF = page.HeaderFooters.Get(hfType, HeaderFooterSection.Footer, HeaderFooterAlignment.Left);
                        if (leftF != null)
                        {
                            SubstitutePageNumbers(pageSettings, dictionaries, leftF, pageNumber, totalPages);
                            int last = leftF.Content.TextLines.Count - 1;
                            var descent = leftF.Content.TextLines[last].LargestDescent;
                            var hfx = pageSettings.Margins.LeftPu;
                            var hfy = pageSettings.Margins.FooterPu + descent;
                            var text = new PdfCellContentLayout(pageSettings, dictionaries, leftF, hfx, hfy, 0, 0);
                            text.Name = "LeftFooter";
                            text.IsHeaderFooter = true;
                            text.GidsAndCharMap(dictionaries);
                            pageLayout.AddChild(text);
                        }
                        var centerF = page.HeaderFooters.Get(hfType, HeaderFooterSection.Footer, HeaderFooterAlignment.Center);
                        if (centerF != null)
                        {
                            SubstitutePageNumbers(pageSettings, dictionaries, centerF, pageNumber, totalPages);
                            int last = centerF.Content.TextLines.Count - 1;
                            var descent = centerF.Content.TextLines[last].LargestDescent;
                            var hfx = pageSettings.PageSize.WidthPu / 2d;
                            var hfy = pageSettings.Margins.FooterPu + descent;
                            var text = new PdfCellContentLayout(pageSettings, dictionaries, centerF, hfx, hfy, 0, 0);
                            text.Name = "CenterFooter";
                            text.IsHeaderFooter = true;
                            text.GidsAndCharMap(dictionaries);
                            pageLayout.AddChild(text);
                        }
                        var rightF = page.HeaderFooters.Get(hfType, HeaderFooterSection.Footer, HeaderFooterAlignment.Right);
                        if (rightF != null)
                        {
                            SubstitutePageNumbers(pageSettings, dictionaries, rightF, pageNumber, totalPages);
                            int last = rightF.Content.TextLines.Count - 1;
                            var descent = rightF.Content.TextLines[last].LargestDescent;
                            var hfx = pageSettings.PageSize.WidthPu - pageSettings.Margins.RightPu;
                            var hfy = pageSettings.Margins.FooterPu + descent;
                            var text = new PdfCellContentLayout(pageSettings, dictionaries, rightF, hfx, hfy, 0, 0);
                            text.Name = "RightFooter";
                            text.IsHeaderFooter = true;
                            text.GidsAndCharMap(dictionaries);
                            pageLayout.AddChild(text);
                        }
                    }
                    PdfGridlinesLayout.AddGridLines(pageSettings, pages[j], pageLayout, borderOnly: !pageSettings.ShowGridLines || pdfPages[i].IsCommentsPage);
                    pageLayout.ChildObjects.Sort((a, b) =>
                    {
                        int cmp = a.Z.CompareTo(b.Z);
                        if (cmp == 0)
                            return string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase);
                        return cmp;
                    });
                    pageNumber++;
                    Catalog.AddChild(pageLayout);
                }
            }
            return Catalog;
        }

        private static void SetFill(PdfDictionaries dictionaries, PdfCellStyle cellStyle, string text, PdfCellLayout fill)
        {
            var xfFill = cellStyle.xfFill;
            var dxfFill = cellStyle.dxfFill;
            if (dxfFill != null && xfFill.IsEmpty())
            {
                var patternStyle = dxfFill.PatternType != null ? (ExcelFillStyle)dxfFill.PatternType : ExcelFillStyle.Solid;
                if (patternStyle == ExcelFillStyle.Solid)
                {
                    fill.SetFill( PdfColor.SetColorFromHex(dxfFill.BackgroundColor.LookupColor()));
                }
                else if (patternStyle != ExcelFillStyle.None)
                {
                    var bkgc = PdfColor.SetColorFromHex(dxfFill.PatternColor.Color == null ? "#FFFFFFFF" : dxfFill.PatternColor.LookupColor());
                    var patc = PdfColor.SetColorFromHex(dxfFill.BackgroundColor.LookupColor());
                    fill.SetPattern(dictionaries, (EPPlus.Export.Pdf.Enums.ExcelFillStyle)patternStyle, bkgc, patc);
                }
                else if (dxfFill.Gradient != null)
                {
                    var gradientType = dxfFill.Gradient.GradientType == null ? ExcelFillGradientType.None : (ExcelFillGradientType)dxfFill.Gradient.GradientType;
                    var color1 = PdfColor.SetColorFromHex(dxfFill.Gradient.Colors[0].Color.LookupColor());
                    var color2 = PdfColor.SetColorFromHex(dxfFill.Gradient.Colors[1].Color.LookupColor());
                    var color3 = PdfColor.SetColorFromHex(dxfFill.Gradient.Colors[2].Color.LookupColor());
                    var degree = dxfFill.Gradient.Degree == null ? 0 : (double)dxfFill.Gradient.Degree;
                    var top = dxfFill.Gradient.Top == null ? 0 : (double)dxfFill.Gradient.Top;
                    var bottom = dxfFill.Gradient.Bottom == null ? 0 : (double)dxfFill.Gradient.Bottom;
                    var left = dxfFill.Gradient.Left == null ? 0 : (double)dxfFill.Gradient.Left;
                    var right  = dxfFill.Gradient.Right == null ? 0 : (double)dxfFill.Gradient.Right;
                    fill.SetGradient(dictionaries, (EPPlus.Export.Pdf.Enums.ExcelFillGradientType)gradientType, color1, color2, color3, degree, top, bottom, left, right);
                }
            }
            else
            {
                if (xfFill.PatternType == ExcelFillStyle.Solid)
                {
                    var bkgc = xfFill.BackgroundColor;
                    var patternStyle = xfFill.PatternType;
                    if (string.IsNullOrEmpty(bkgc.LookupColor()) && !string.IsNullOrEmpty(text))
                    {
                        fill.SetFill(Color.Empty);
                    }
                    else
                    {
                        fill.SetFill(PdfColor.SetColorFromHex(bkgc.LookupColor()));
                    }
                }
                else if (xfFill.PatternType != ExcelFillStyle.None)
                {
                    var  patternStyle = xfFill.PatternType;
                    var bkgc = PdfColor.SetColorFromHex(xfFill.PatternColor.Rgb == null ? "#FFFFFFFF" : xfFill.PatternColor.LookupColor());
                    var patc = PdfColor.SetColorFromHex(xfFill.BackgroundColor.LookupColor());
                    fill.SetPattern(dictionaries, (EPPlus.Export.Pdf.Enums.ExcelFillStyle)patternStyle, bkgc, patc);
                }
                else if (xfFill.HasGradient)
                {
                    var gradientType = xfFill.Gradient.Type;
                    var color1 = PdfColor.SetColorFromHex(xfFill.Gradient.Color1.LookupColor());
                    var color2 = PdfColor.SetColorFromHex(xfFill.Gradient.Color2.LookupColor());
                    var color3 = PdfColor.SetColorFromHex(xfFill.Gradient.Color3.LookupColor());
                    var degree = xfFill.Gradient.Degree;
                    var top = double.IsNaN(xfFill.Gradient.Top) ? 0 : xfFill.Gradient.Top;
                    var bottom = double.IsNaN(xfFill.Gradient.Bottom) ? 0 : xfFill.Gradient.Bottom;
                    var left = double.IsNaN(xfFill.Gradient.Left) ? 0 : xfFill.Gradient.Left;
                    var right = double.IsNaN(xfFill.Gradient.Right) ? 0 : xfFill.Gradient.Right;
                    fill.SetGradient(dictionaries, (EPPlus.Export.Pdf.Enums.ExcelFillGradientType)gradientType, color1, color2, color3, degree, top, bottom, left, right);
                }
            }
        }

        static MergedCellCorners GetCorners(ExcelAddressBase addr, int row, int col)
        {
            if (addr == null) return MergedCellCorners.All;
            MergedCellCorners result = MergedCellCorners.None;

            if (row == addr.Start.Row && col == addr.Start.Column)
                result |= MergedCellCorners.TopLeft;

            if (row == addr.Start.Row && col == addr.End.Column)
                result |= MergedCellCorners.TopRight;

            if (row == addr.End.Row && col == addr.Start.Column)
                result |= MergedCellCorners.BottomLeft;

            if (row == addr.End.Row && col == addr.End.Column)
                result |= MergedCellCorners.BottomRight;
            return result;
        }

        private static bool NeedsClipping(PdfCell map, Page page, int row, int col)
        {
            if (map.ContentAligmnet == null) return false;
            // Fill alignment always clips; WrapText is already wrapped but clip for safety.
            if (map.ContentAligmnet.HorizontalAlignment == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Fill || map.ContentAligmnet.WrapText)
                return true;
            if (map.TotalTextLength <= map.ColumnWidth) return false;
            var halign = map.ContentAligmnet.HorizontalAlignment;
            if (halign == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Left || halign == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.General)
            {
                // Text spills right — clip if the right neighbour has content or we're at the page edge.
                if (col >= page.ToColumn) return true;
                var right = page.Map[row, col + 1];
                return right != null && !string.IsNullOrEmpty(right.Text);
            }
            else if (halign == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Right)
            {
                // Text spills left — clip if the left neighbour has content or we're at the page edge.
                if (col <= page.FromColumn) return true;
                var left = page.Map[row, col - 1];
                return left != null && !string.IsNullOrEmpty(left.Text);
            }
            else if (halign == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Center)
            {
                // Text spills both ways — clip if either neighbour blocks or we're at an edge.
                bool rightBlocked = col >= page.ToColumn || (page.Map[row, col + 1] != null && !string.IsNullOrEmpty(page.Map[row, col + 1].Text));
                bool leftBlocked = col <= page.FromColumn || (page.Map[row, col - 1] != null && !string.IsNullOrEmpty(page.Map[row, col - 1].Text));
                return rightBlocked || leftBlocked;
            }
            return false;
        }
        private static bool HasBorder(PdfCellStyle cellStyle)
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

        private static bool HasDiagonalBorder(PdfCellStyle style) => style?.Diagonal != null && style.Diagonal.Style != ExcelBorderStyle.None;

        private static void AddHeadingCell(PdfPageSettings pageSettings, PdfDictionaries dictionaries, PdfPageLayout pageLayout, PdfCellStyle headingStyle, string label, double x, double y, double width, double height, string fontName, float fontSize, string namePrefix)
        {
            if (width == 0d || height == 0d) return;
            var fill = new PdfCellLayout(x, y, width, height);
            SetFill(dictionaries, headingStyle, label, fill);
            fill.Name = namePrefix;
            fill.IsHeading = true;
            fill.UpdateShadingPositionMatrix(pageSettings);
            pageLayout.AddChild(fill);
            var cell = CreateHeadingPdfCell(pageSettings, dictionaries, label, ExcelHorizontalAlignment.Center, width, height, fontName, fontSize);
            if (cell.TextLines != null && cell.TextLines.Count > 0)
            {
                var info = new MergedCellDrawInfo { X = x, Y = y, Width = width, Height = height };
                var text = new PdfCellContentLayout(pageSettings, dictionaries, cell, info, x, y, width, height);
                text.Name = namePrefix + "_Text";
                text.IsHeading = true;
                text.GidsAndCharMap(dictionaries);
                pageLayout.AddChild(text);
            }
            AddHeadingCellBorder(pageLayout, x, y, width, height);
        }

        private static void AddHeadingCells(PdfPageSettings pageSettings, PdfDictionaries dictionaries, Page page, PdfPageLayout pageLayout, double contentStartX, double contentStartY, double headingWidth, double headingHeight, string fontName, float fontSize, ExcelFill fill)
        {
            var headingStyle = new PdfCellStyle();
            headingStyle.xfFill = fill;
            var cornerFill = new PdfCellLayout(pageSettings.ContentBounds.Left, pageSettings.ContentBounds.Top, headingWidth, headingHeight);
            SetFill(dictionaries, headingStyle, "", cornerFill);
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
                    x, pageSettings.ContentBounds.Top, colWidth, headingHeight, fontName, fontSize, "Heading_Col_" + colLetter);
                x += colWidth;
            }
            double y = contentStartY;
            for (int row = page.FromRow; row <= page.ToRow; row++)
            {
                double rowHeight = page.RowHeights[row - page.FromRow];
                if (rowHeight == 0d) { y -= rowHeight; continue; }
                string rowNum = row.ToString();
                AddHeadingCell(pageSettings, dictionaries, pageLayout, headingStyle, rowNum,
                    pageSettings.ContentBounds.Left, y, headingWidth, rowHeight, fontName, fontSize, "Heading_Row_" + rowNum);
                y -= rowHeight;
            }
        }

        private static void AddPrintTitleHeadings(PdfPageSettings pageSettings, PdfDictionaries dictionaries, Page page, PdfPageLayout pageLayout, string fontName, float fontSize, ExcelFill fill)
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
            pageLayout.BorderLines.Add(new GridLine(x, y, right, y));       // top
            pageLayout.BorderLines.Add(new GridLine(x, bottom, right, bottom));  // bottom
            pageLayout.BorderLines.Add(new GridLine(x, y, x, bottom));  // left
            pageLayout.BorderLines.Add(new GridLine(right, y, right, bottom));  // right
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
                tf.RichTextOptions.UnderlineType = 12;  // none
                tf.RichTextOptions.StrikeType = 1;   // none
                cell.TextFragments = new List<TextFragment> { tf };
                PdfTextShaper.ShapeText(pageSettings, dictionaries, cell);
            }
            return cell;
        }

        private static void AddSpillCells(PdfPageSettings pageSettings, PdfDictionaries dictionaries, Page page, PdfPageLayout pageLayout)
        {
            if (page.SpillCells == null) return;
            foreach (var s in page.SpillCells)
            {
                var map = s.Cell;
                if (map.TextLines == null || map.TextLines.Count == 0) continue;
                if (s.ClipWidth <= 0d || s.ClipHeight <= 0d) continue;

                var text = new PdfCellContentLayout(pageSettings, dictionaries, map, new MergedCellDrawInfo(), s.X, s.Y, s.Width, s.Height);
                text.Name = "Spill_" + map.Name;
                text.IsPrintTitle = s.IsPrintTitle;       // false → clipped content group; true → outside-clip (band)
                text.GidsAndCharMap(dictionaries);
                text.SetupClipping(s.ClipX, s.ClipY, s.ClipWidth, s.ClipHeight);
                pageLayout.AddChild(text);
            }
        }

        private static void AddPrintTitleCells(PdfPageSettings pageSettings, PdfDictionaries dictionaries, Page page, PdfPageLayout pageLayout)
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
                    var text = new PdfCellContentLayout(pageSettings, dictionaries, map, info, t.X, t.Y, t.Width, t.Height);
                    text.Name = map.Name;
                    text.IsPrintTitle = true;
                    text.GidsAndCharMap(dictionaries);
                    text.SetupClipping(t.ClipX, t.ClipY, t.ClipWidth, t.ClipHeight);
                    pageLayout.AddChild(text);
                }
                if (!map.Merged && HasBorder(map.CellStyle))   // was: if (HasBorder(map.CellStyle))
                {
                    var border = new PdfCellBorderLayout(false, MergedCellCorners.All, new MergedCellDrawInfo(), t.X, t.Y, t.Width, t.Height);
                    SetBorderStyle(map.CellStyle, border);
                    border.IsPrintTitle = true;
                    border.Name = "PrintTitleBorder_" + map.Name;
                    pageLayout.AddChild(border);
                }
                // Per-cell borders for merged band cells — outside the margin clip like the rest of the band.
                foreach (var b in page.PrintTitleBorders)
                {
                    var sub = b.Cell;
                    if (!HasBorder(sub.CellStyle)) continue;
                    var border = new PdfCellBorderLayout(false, MergedCellCorners.All, new MergedCellDrawInfo(), b.X, b.Y, b.Width, b.Height);
                    SetBorderStyle(sub.CellStyle, border);
                    border.IsPrintTitle = true;
                    border.Name = "PrintTitleMergeBorder_" + sub.Name;
                    pageLayout.AddChild(border);
                }
            }
        }

        private static void SetBorderStyle(PdfCellStyle style, PdfCellBorderLayout border)
        {
            var topStyle = style.xfTop.Style == ExcelBorderStyle.None ? ((style.dxfTop != null && style.dxfTop.HasValue) ? (ExcelBorderStyle)style.dxfTop.Style : ExcelBorderStyle.None) : style.xfTop.Style;
            var topColor = (style.xfTop.Style == ExcelBorderStyle.None && style.dxfTop != null) ? PdfColor.SetColorFromHex(style.dxfTop.Color.LookupColor(style.dxfTop)) : PdfColor.SetColorFromHex(style.xfTop.Color.LookupColor(style.xfTop));

            var bottomStyle = style.xfBottom.Style == ExcelBorderStyle.None ? ((style.dxfBottom != null && style.dxfBottom.HasValue) ? (ExcelBorderStyle)style.dxfBottom.Style : ExcelBorderStyle.None) : style.xfBottom.Style;
            var bottomColor = (style.xfBottom.Style == ExcelBorderStyle.None && style.dxfBottom != null) ? PdfColor.SetColorFromHex(style.dxfBottom.Color.LookupColor(style.dxfBottom)) : PdfColor.SetColorFromHex(style.xfBottom.Color.LookupColor(style.xfBottom));

            var leftStyle = style.xfLeft.Style == ExcelBorderStyle.None ? ((style.dxfLeft != null && style.dxfLeft.HasValue) ? (ExcelBorderStyle)style.dxfLeft.Style : ExcelBorderStyle.None) : style.xfLeft.Style;
            var leftColor = (style.xfLeft.Style == ExcelBorderStyle.None && style.dxfLeft != null) ? PdfColor.SetColorFromHex(style.dxfLeft.Color.LookupColor(style.dxfLeft)) : PdfColor.SetColorFromHex(style.xfLeft.Color.LookupColor(style.xfLeft));

            var rightStyle = style.xfRight.Style == ExcelBorderStyle.None ? ((style.dxfRight != null && style.dxfRight.HasValue) ? (ExcelBorderStyle)style.dxfRight.Style : ExcelBorderStyle.None) : style.xfRight.Style;
            var rightColor = (style.xfRight.Style == ExcelBorderStyle.None && style.dxfRight != null) ? PdfColor.SetColorFromHex(style.dxfRight.Color.LookupColor(style.dxfRight)) : PdfColor.SetColorFromHex(style.xfRight.Color.LookupColor(style.xfRight));

            var diagUpStyle = style.DiagonalUp ? style.Diagonal.Style : ExcelBorderStyle.None;
            var diagUpColor = style.DiagonalUp ? PdfColor.SetColorFromHex(style.Diagonal.Color.LookupColor(style.Diagonal)) : Color.Transparent;

            var diagDownStyle = style.DiagonalDown ? style.Diagonal.Style : ExcelBorderStyle.None;
            var diagDownColor = style.DiagonalDown ? PdfColor.SetColorFromHex(style.Diagonal.Color.LookupColor(style.Diagonal)) : Color.Transparent;

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
        private static int GetTotalPages(List<Pages> pdfPages)
        {
            int totalPages = 0;
            for (int i = 0; i < pdfPages.Count; i++)
            {
                totalPages += pdfPages[i].Page.Length;
            }
            return totalPages;
        }

        internal static List<Pages> GetPages(PdfPageSettings[] sheetSettings, PdfWorksheet[] pdfSheets)
        {
            List<Pages> PagesCollection = new List<Pages>();
            for (int si = 0; si < pdfSheets.Length; si++)
            {
                var pdfSheet= pdfSheets[si];
                var pageSettings = sheetSettings[si];            

                for (int ri = 0; ri < pdfSheet.Ranges.Count; ri++)
                {
                    var range = pdfSheet.Ranges[ri];
                    var pages = GetNumberOfPages(pageSettings, pdfSheet, ref range);
                    pages = AssignRangeToPages(pageSettings, range, pages);
                    pages = MapPage(range, pages);
                    pages = GetHeaderFooter(range, pages, pdfSheet);
                    pages = PrecomputeMergedCells(pageSettings, range, pages);
                    pages = PrecomputeSpillCells(pageSettings, range, pages);
                    pages = PrecomputePrintTitleCells(pageSettings, pdfSheet, range, pages);
                    pages.HeadingFontName = pdfSheet.NormalStyle.Style.Font.Name;
                    pages.HeadingFontSize = pdfSheet.NormalStyle.Style.Font.Size;
                    pages.HeadingFill = pdfSheet.NormalStyle.Style.Fill;
                    pages.Settings = pageSettings;
                    PagesCollection.Add(pages);
                }
                if (pdfSheet.CommentsAndNotes.Range != null)
                {
                    bool savedShowHeadings = pageSettings.ShowHeadings;
                    pageSettings.ShowHeadings = false;
                    var pages = GetNumberOfPages(pageSettings, pdfSheet, ref pdfSheet.CommentsAndNotes);
                    pages = AssignRangeToPages(pageSettings, pdfSheet.CommentsAndNotes, pages);
                    pages = MapPage(pdfSheet.CommentsAndNotes, pages);
                    pageSettings.ShowHeadings = savedShowHeadings;
                    pages.IsCommentsPage = true;
                    pages.Settings = pageSettings;                    
                    PagesCollection.Add(pages);
                }
            }
            return PagesCollection;
        }

        internal static Pages PrecomputeMergedCells(PdfPageSettings pageSettings, PdfRange range, Pages pdfPages)
        {
            for (int i = 0; i < pdfPages.Page.Length; i++)
                pdfPages.Page[i] = PrecomputePageMergedCells(pageSettings, range, pdfPages.Page[i]);
            return pdfPages;
        }

        private static Page PrecomputePageMergedCells(PdfPageSettings pageSettings, PdfRange range, Page page)
        {
            page.MergedCells = new Dictionary<string, MergedCellDrawInfo>();
            // Build a quick lookup: absolute x for each column index on this page.
            // We read ColumnWidth from the first data row; widths are per-column, not per-cell.
            var colX = BuildColumnXPositions(pageSettings, page);
            for (int row = page.FromRow; row <= page.ToRow; row++)
            {
                for (int col = page.FromColumn; col <= page.ToColumn; col++)
                {
                    var cell = page.Map[row, col];
                    if (cell == null || !cell.Merged) continue;
                    string key = cell.MergedAddress.Address;
                    if (page.MergedCells.ContainsKey(key)) continue;
                    var addr = cell.MergedAddress;
                    // --- X ---
                    // Start from the current column and walk left to the merge origin.
                    // Columns within the current page come from colX; columns that lie on
                    // a preceding column-page come from range.ColWidths.
                    double drawX = colX[col - page.FromColumn];
                    for (int c = addr._fromCol; c < col; c++)
                    {
                        int rangeIdx = c - range.Range._fromCol;
                        if (rangeIdx >= 0 && rangeIdx < range.ColWidths.Count)
                            drawX -= range.ColWidths[rangeIdx];
                    }
                    // --- Y ---
                    // Replace the * 15d line with a sum of real row heights
                    double drawY = pageSettings.ContentBounds.Top - page.HeadingHeight - page.PrintTitleHeight;
                    for (int r = page.FromRow; r < row; r++)
                    {
                        drawY -= range.RowHeights[r - range.Range._fromRow].Height;
                    }
                    // Existing loop — just add .Height
                    for (int r = addr._fromRow; r < row; r++)
                    {
                        int rangeIdx = r - range.Range._fromRow;
                        if (rangeIdx >= 0 && rangeIdx < range.RowHeights.Count)
                            drawY += range.RowHeights[rangeIdx].Height;
                    }
                    // --- Width / Height ---
                    // Size the merge from the SAME arrays that produced X/Y above
                    // (range.ColWidths / range.RowHeights) rather than from mainCell.
                    // Those arrays already store 0 for hidden rows/columns and use the
                    // same unit conversion, default-height and auto-fit values as the
                    // rest of the grid, so the merge rectangle can never disagree with
                    // the surrounding cells (which is why columns worked but rows did not).
                    double mergeWidth = 0d;
                    for (int c = addr._fromCol; c <= addr._toCol; c++)
                    {
                        int rangeIdx = c - range.Range._fromCol;
                        if (rangeIdx >= 0 && rangeIdx < range.ColWidths.Count)
                            mergeWidth += range.ColWidths[rangeIdx];
                    }
                    double mergeHeight = 0d;
                    for (int r = addr._fromRow; r <= addr._toRow; r++)
                    {
                        int rangeIdx = r - range.Range._fromRow;
                        if (rangeIdx >= 0 && rangeIdx < range.RowHeights.Count)
                            mergeHeight += range.RowHeights[rangeIdx].Height;
                    }
                    page.MergedCells[key] = new MergedCellDrawInfo
                    {
                        X = drawX,
                        Y = drawY,
                        Width = mergeWidth,
                        Height = mergeHeight
                    };
                }
            }
            return page;
        }

        private static double[] BuildColumnXPositions(PdfPageSettings pageSettings, Page page)
        {
            int colCount = page.ToColumn - page.FromColumn + 1;
            var colX = new double[colCount];
            double x = pageSettings.ContentBounds.Left + page.HeadingWidth + page.PrintTitleWidth;
            for (int col = page.FromColumn; col <= page.ToColumn; col++)
            {
                colX[col - page.FromColumn] = x;
                var cell = page.Map[page.FromRow, col];
                x += cell?.ColumnWidth ?? 0d;
            }
            return colX;
        }

        private static void ComputePrintTitleDimensions(PdfWorksheet pdfSheet, PdfRange range, out double titleHeight, out double titleWidth)
        {
            titleHeight = 0d;
            titleWidth = 0d;
            if (pdfSheet.PrintTitleRowFrom >= 0)
            {
                for (int r = pdfSheet.PrintTitleRowFrom; r <= pdfSheet.PrintTitleRowTo; r++)
                {
                    int idx = r - range.Range._fromRow;
                    if (idx >= 0 && idx < range.RowHeights.Count)
                        titleHeight += range.RowHeights[idx].Height;
                }
            }
            if (pdfSheet.PrintTitleColFrom >= 0)
            {
                for (int c = pdfSheet.PrintTitleColFrom; c <= pdfSheet.PrintTitleColTo; c++)
                {
                    int idx = c - range.Range._fromCol;
                    if (idx >= 0 && idx < range.ColWidths.Count)
                        titleWidth += range.ColWidths[idx];
                }
            }
        }

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
            if (!topBand && !leftBand) return page;

            // Content-column X: same origin/widths the content loop uses (step-2 origin).
            var contentColX = new Dictionary<int, double>();
            double cx = pageSettings.ContentBounds.Left + page.HeadingWidth + page.PrintTitleWidth;
            for (int c = page.FromColumn; c <= page.ToColumn; c++)
            {
                contentColX[c] = cx;
                cx += RangeColWidth(range, page.FromRow, c);
            }
            // Content-row Y.
            var contentRowY = new Dictionary<int, double>();
            double cy = pageSettings.ContentBounds.Top - page.HeadingHeight - page.PrintTitleHeight;
            for (int r = page.FromRow; r <= page.ToRow; r++)
            {
                contentRowY[r] = cy;
                cy -= RangeRowHeight(range, r);
            }
            // Title-column X (left band): just right of the heading gutter.
            var titleColX = new Dictionary<int, double>();
            if (leftBand)
            {
                double tx = pageSettings.ContentBounds.Left + page.HeadingWidth;
                for (int c = pdfSheet.PrintTitleColFrom; c <= pdfSheet.PrintTitleColTo; c++)
                {
                    titleColX[c] = tx;
                    tx += RangeColWidth(range, page.FromRow, c);
                }
            }
            // Title-row Y (top band): just below the heading gutter.
            var titleRowY = new Dictionary<int, double>();
            if (topBand)
            {
                double ty = pageSettings.ContentBounds.Top - page.HeadingHeight;
                for (int r = pdfSheet.PrintTitleRowFrom; r <= pdfSheet.PrintTitleRowTo; r++)
                {
                    titleRowY[r] = ty;
                    ty -= RangeRowHeight(range, r);
                }
            }
            if (topBand)
                ProcessBandRegionCells(page, range, pdfSheet.PrintTitleRowFrom, pdfSheet.PrintTitleRowTo, page.FromColumn, page.ToColumn, contentColX, titleRowY);
            if (leftBand)
                ProcessBandRegionCells(page, range, page.FromRow, page.ToRow, pdfSheet.PrintTitleColFrom, pdfSheet.PrintTitleColTo, titleColX, contentRowY);
            if (topBand && leftBand)
                ProcessBandRegionCells(page, range, pdfSheet.PrintTitleRowFrom, pdfSheet.PrintTitleRowTo, pdfSheet.PrintTitleColFrom, pdfSheet.PrintTitleColTo, titleColX, titleRowY);

            page.PrintTitleGridLines = new List<GridLine>();
            if (pageSettings.ShowGridLines)
            {
                if (topBand)
                    EmitBandGrid(page.PrintTitleGridLines, range, pdfSheet.PrintTitleRowFrom, pdfSheet.PrintTitleRowTo, page.FromColumn, page.ToColumn,
                        BandEdgesX(contentColX, page.FromColumn, page.ToColumn, range, page.FromRow),
                        BandEdgesY(titleRowY, pdfSheet.PrintTitleRowFrom, pdfSheet.PrintTitleRowTo, range));
                if (leftBand)
                    EmitBandGrid(page.PrintTitleGridLines, range, page.FromRow, page.ToRow, pdfSheet.PrintTitleColFrom, pdfSheet.PrintTitleColTo,
                        BandEdgesX(titleColX, pdfSheet.PrintTitleColFrom, pdfSheet.PrintTitleColTo, range, page.FromRow),
                        BandEdgesY(contentRowY, page.FromRow, page.ToRow, range));
                if (topBand && leftBand)
                    EmitBandGrid(page.PrintTitleGridLines, range, pdfSheet.PrintTitleRowFrom, pdfSheet.PrintTitleRowTo, pdfSheet.PrintTitleColFrom, pdfSheet.PrintTitleColTo,
                        BandEdgesX(titleColX, pdfSheet.PrintTitleColFrom, pdfSheet.PrintTitleColTo, range, page.FromRow),
                        BandEdgesY(titleRowY, pdfSheet.PrintTitleRowFrom, pdfSheet.PrintTitleRowTo, range));
            }
            // ---- band headings: original row numbers / column letters, in the gutter gaps ----
            page.PrintTitleHeadings = new List<PrintTitleHeadingDraw>();
            if (pageSettings.ShowHeadings)
            {
                if (topBand)
                    for (int r = pdfSheet.PrintTitleRowFrom; r <= pdfSheet.PrintTitleRowTo; r++)
                    {
                        double h = RangeRowHeight(range, r);
                        if (h <= 0d) continue;
                        page.PrintTitleHeadings.Add(new PrintTitleHeadingDraw
                        {
                            IsRow = true,
                            Index = r,
                            X = pageSettings.ContentBounds.Left,
                            Y = titleRowY[r],
                            Width = page.HeadingWidth,
                            Height = h
                        });
                    }

                if (leftBand)
                    for (int c = pdfSheet.PrintTitleColFrom; c <= pdfSheet.PrintTitleColTo; c++)
                    {
                        double w = RangeColWidth(range, page.FromRow, c);
                        if (w <= 0d) continue;
                        page.PrintTitleHeadings.Add(new PrintTitleHeadingDraw
                        {
                            IsRow = false,
                            Index = c,
                            X = titleColX[c],
                            Y = pageSettings.ContentBounds.Top,
                            Width = w,
                            Height = page.HeadingHeight
                        });
                    }
            }

            // repeated title-row text continues onto the next horizontal page's band
            if (topBand)
                AddIncomingSpill(page, range, pdfSheet.PrintTitleRowFrom, pdfSheet.PrintTitleRowTo, page.FromColumn, page.ToColumn,
                    pageSettings.ContentBounds.Left + page.HeadingWidth + page.PrintTitleWidth,
                    pageSettings.ContentBounds.Top - page.HeadingHeight,
                    isPrintTitle: true);

            // left band: a neighbour whose text spills INTO a title column travels with the repeated column
            if (leftBand)
                AddIncomingSpill(page, range, page.FromRow, page.ToRow,
                    pdfSheet.PrintTitleColFrom, pdfSheet.PrintTitleColTo,
                    pageSettings.ContentBounds.Left + page.HeadingWidth,                                  // band origin X (left edge of the title columns)
                    pageSettings.ContentBounds.Top - page.HeadingHeight - page.PrintTitleHeight,           // content-rows origin Y
                    isPrintTitle: true);

            // corner: same, for the title-rows × title-columns intersection
            if (topBand && leftBand)
                AddIncomingSpill(page, range, pdfSheet.PrintTitleRowFrom, pdfSheet.PrintTitleRowTo,
                    pdfSheet.PrintTitleColFrom, pdfSheet.PrintTitleColTo,
                    pageSettings.ContentBounds.Left + page.HeadingWidth,                                  // band origin X
                    pageSettings.ContentBounds.Top - page.HeadingHeight,                                  // title-rows origin Y
                    isPrintTitle: true);

            return page;
        }

        private static void ProcessBandRegionCells(Page page, PdfRange range, int fromRow, int toRow, int fromCol, int toCol, Dictionary<int, double> colX, Dictionary<int, double> rowY)
        {
            // Band-region rectangle (top-left origin). Non-merged text may spill within this rect,
            // which is bounded by the band edge, so it can never overflow into the content area.
            double regLeft = colX[fromCol];
            double regRight = colX[toCol] + RangeColWidth(range, page.FromRow, toCol);
            double regTop = rowY[fromRow];
            double regBottom = rowY[toRow] - RangeRowHeight(range, toRow);
            double regW = regRight - regLeft;
            double regH = regTop - regBottom;
            var drawnMerges = new HashSet<string>();
            for (int r = fromRow; r <= toRow; r++)
            {
                for (int c = fromCol; c <= toCol; c++)
                {
                    var cell = RangeCell(range, r, c);
                    if (cell == null) continue;

                    if (cell.Merged)
                    {
                        if (!drawnMerges.Add(cell.MergedAddress.Address)) continue;
                        var addr = cell.MergedAddress;
                        var main = cell.Main ?? cell;
                        int vFromCol = Math.Max(addr._fromCol, fromCol);
                        int vToCol = Math.Min(addr._toCol, toCol);
                        int vFromRow = Math.Max(addr._fromRow, fromRow);
                        int vToRow = Math.Min(addr._toRow, toRow);
                        if (vFromCol > vToCol || vFromRow > vToRow) continue;
                        double x = colX[vFromCol];
                        double width = (colX[vToCol] + RangeColWidth(range, page.FromRow, vToCol)) - colX[vFromCol];
                        double y = rowY[vFromRow];
                        double height = (rowY[vFromRow] - rowY[vToRow]) + RangeRowHeight(range, vToRow);
                        if (width <= 0d || height <= 0d) continue;
                        // merged cells don't spill — clip text to the merge itself (matches content)
                        page.PrintTitleCells.Add(new PrintTitleCellDraw
                        {
                            Cell = main,
                            X = x,
                            Y = y,
                            Width = width,
                            Height = height,
                            ClipX = x,
                            ClipY = y,
                            ClipWidth = width,
                            ClipHeight = height
                        });
                        for (int br = vFromRow; br <= vToRow; br++)
                        {
                            double bh = RangeRowHeight(range, br);
                            if (bh <= 0d) continue;
                            for (int bc = vFromCol; bc <= vToCol; bc++)
                            {
                                double bw = RangeColWidth(range, page.FromRow, bc);
                                if (bw <= 0d) continue;
                                var sub = RangeCell(range, br, bc);
                                if (sub == null || !HasBorder(sub.CellStyle)) continue;
                                page.PrintTitleBorders.Add(new PrintTitleCellDraw
                                {
                                    Cell = sub,
                                    X = colX[bc],
                                    Y = rowY[br],
                                    Width = bw,
                                    Height = bh
                                });
                            }
                        }
                    }
                    else
                    {
                        double w = RangeColWidth(range, page.FromRow, c);
                        double h = RangeRowHeight(range, r);
                        if (w <= 0d || h <= 0d) continue;
                        // non-merged: clip to the band region so text can spill within it but not into content
                        page.PrintTitleCells.Add(new PrintTitleCellDraw
                        {
                            Cell = cell,
                            X = colX[c],
                            Y = rowY[r],
                            Width = w,
                            Height = h,
                            ClipX = regLeft,
                            ClipY = regTop,
                            ClipWidth = regW,
                            ClipHeight = regH
                        });
                    }
                }
            }
        }

        private static void EmitBandGrid(List<GridLine> target, PdfRange range, int fromRow, int toRow, int fromCol, int toCol, double[] colX, double[] rowY)
        {
            int nr = toRow - fromRow + 1;
            int nc = toCol - fromCol + 1;
            if (nr <= 0 || nc <= 0) return;
            var spill = BuildBandSpillMask(range, fromRow, toRow, fromCol, toCol, colX);
            double top = rowY[0], bottom = rowY[nr];
            double left = colX[0], right = colX[nc];
            EmitBandFrameH(target, range, top, fromRow, fromCol, nc, colX, CellHasTopBorder);
            EmitBandFrameH(target, range, bottom, toRow, fromCol, nc, colX, CellHasBottomBorder);
            EmitBandFrameV(target, range, left, fromCol, fromRow, nr, rowY, CellHasLeftBorder);
            EmitBandFrameV(target, range, right, toCol, fromRow, nr, rowY, CellHasRightBorder);
            // interior verticals — suppress where a merge spans the gap
            for (int gi = 1; gi < nc; gi++)
            {
                int leftCol = fromCol + gi - 1, rightCol = leftCol + 1;
                double x = colX[gi];
                double? runStart = null; double runEnd = 0d;
                for (int ri = 0; ri < nr; ri++)
                {
                    int r = fromRow + ri;
                    //bool block = SameMerge(range, r, leftCol, r, rightCol) || spill[ri, gi];
                    bool block = SameMerge(range, r, leftCol, r, rightCol) || spill[ri, gi] ||
                                            CellHasRightBorder(RangeCell(range, r, leftCol)) ||
                                            CellHasLeftBorder(RangeCell(range, r, rightCol));
                    if (!block) { if (runStart == null) runStart = rowY[ri]; runEnd = rowY[ri + 1]; }
                    else if (runStart != null) { target.Add(new GridLine(x, runStart.Value, x, runEnd)); runStart = null; }
                }
                if (runStart != null) target.Add(new GridLine(x, runStart.Value, x, runEnd));
            }
            // interior horizontals — suppress where a merge spans the gap
            for (int gj = 1; gj < nr; gj++)
            {
                int topRow = fromRow + gj - 1, bottomRow = topRow + 1;
                double y = rowY[gj];
                double? runStart = null; double runEnd = 0d;
                for (int ci = 0; ci < nc; ci++)
                {
                    int c = fromCol + ci;
                    bool block = SameMerge(range, topRow, c, bottomRow, c) ||
                                            CellHasBottomBorder(RangeCell(range, topRow, c)) ||
                                            CellHasTopBorder(RangeCell(range, bottomRow, c));
                    if (!block) { if (runStart == null) runStart = colX[ci]; runEnd = colX[ci + 1]; }
                    else if (runStart != null) { target.Add(new GridLine(runStart.Value, y, runEnd, y)); runStart = null; }
                }
                if (runStart != null) target.Add(new GridLine(runStart.Value, y, runEnd, y));
            }
        }

        private static bool[,] BuildBandSpillMask(PdfRange range, int fromRow, int toRow, int fromCol, int toCol, double[] colX)
        {
            int nr = toRow - fromRow + 1;
            int nc = toCol - fromCol + 1;
            var blocked = new bool[Math.Max(nr, 1), Math.Max(nc, 1)]; // [ri, g], g in 1..nc-1
            if (nr <= 0 || nc <= 1) return blocked;
            int repRow = fromRow;
            for (int ri = 0; ri < nr; ri++)
            {
                int row = fromRow + ri;
                // (a) cells spilling within the band region
                for (int ci = 0; ci < nc; ci++)
                {
                    var cell = RangeCell(range, row, fromCol + ci);
                    if (cell == null || cell.Merged) continue;
                    if (cell.ContentAligmnet != null && cell.ContentAligmnet.WrapText) continue;
                    if (string.IsNullOrEmpty(cell.Text)) continue;
                    double spill = cell.TotalTextLength - cell.ColumnWidth;
                    if (spill <= 0d) continue;
                    var hal = cell.ContentAligmnet?.HorizontalAlignment ?? (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.General;
                    if (hal == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Left || hal == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.General)
                        BandMarkRight(range, row, ci, nc, fromCol, colX, spill, ri, blocked);
                    else if (hal == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Right)
                        BandMarkLeft(range, row, ci, fromCol, colX, spill, ri, blocked);
                    else if (hal == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Center)
                    {
                        double half = spill / 2d;
                        BandMarkRight(range, row, ci, nc, fromCol, colX, half, ri, blocked);
                        BandMarkLeft(range, row, ci, fromCol, colX, half, ri, blocked);
                    }
                }
                // (b) spill entering from the LEFT of the region (left/general/center → spilling right in)
                double lx = colX[0];
                for (int c = fromCol - 1; c >= range.Range._fromCol; c--)
                {
                    double w = RangeColWidth(range, repRow, c);
                    lx -= w;
                    var cell = RangeCell(range, row, c);
                    if (cell == null) continue;
                    if (cell.Merged) break;
                    if (string.IsNullOrEmpty(cell.Text)) continue;
                    var hal = cell.ContentAligmnet?.HorizontalAlignment ?? (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.General;
                    if (hal == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Right) break; // spills away from the region
                    double rightExtent = (hal == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Center)
                        ? lx + w / 2d + cell.TotalTextLength / 2d
                        : lx + cell.TotalTextLength;
                    for (int g = 1; g <= nc - 1; g++)
                    {
                        var blk = RangeCell(range, row, fromCol + g - 1);
                        if (blk != null && !string.IsNullOrEmpty(blk.Text)) break;
                        if (colX[g] < rightExtent) blocked[ri, g] = true; else break;
                    }
                    break;
                }

                // (c) spill entering from the RIGHT of the region (right/center → spilling left in)
                double rx = colX[nc];
                for (int c = toCol + 1; c <= range.Range._toCol; c++)
                {
                    double w = RangeColWidth(range, repRow, c);
                    var cell = RangeCell(range, row, c);
                    if (cell == null) { rx += w; continue; }
                    if (cell.Merged) break;
                    if (string.IsNullOrEmpty(cell.Text)) { rx += w; continue; }
                    var hal = cell.ContentAligmnet?.HorizontalAlignment ?? (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.General;
                    if (hal == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Left || hal == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.General) break; // spills away
                    double leftExtent = (hal == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Center)
                        ? rx + w / 2d - cell.TotalTextLength / 2d
                        : rx + w - cell.TotalTextLength;
                    for (int g = nc - 1; g >= 1; g--)
                    {
                        var blk = RangeCell(range, row, fromCol + g);
                        if (blk != null && !string.IsNullOrEmpty(blk.Text)) break;
                        if (colX[g] > leftExtent) blocked[ri, g] = true; else break;
                    }
                    break;
                }
            }
            return blocked;
        }

        private static void BandMarkRight(PdfRange range, int row, int ci, int nc, int fromCol, double[] colX, double spill, int ri, bool[,] blocked)
        {
            for (int g = ci + 1; g <= nc - 1; g++)
            {
                var rightCell = RangeCell(range, row, fromCol + g);
                if (rightCell != null && !string.IsNullOrEmpty(rightCell.Text)) break;
                double distToGap = colX[g] - colX[ci + 1];
                if (spill > distToGap) blocked[ri, g] = true; else break;
            }
        }

        private static void BandMarkLeft(PdfRange range, int row, int ci, int fromCol, double[] colX, double spill, int ri, bool[,] blocked)
        {
            for (int g = ci; g >= 1; g--)
            {
                var leftCell = RangeCell(range, row, fromCol + g - 1);
                if (leftCell != null && !string.IsNullOrEmpty(leftCell.Text)) break;
                double distToGap = colX[ci] - colX[g];
                if (spill > distToGap) blocked[ri, g] = true; else break;
            }
        }

        private static bool SameMerge(PdfRange range, int r1, int c1, int r2, int c2)
        {
            var a = RangeCell(range, r1, c1);
            var b = RangeCell(range, r2, c2);
            if (a == null || b == null || !a.Merged || !b.Merged) return false;
            return a.MergedAddress.Address == b.MergedAddress.Address;
        }

        // Column edges for a band region: left edge of every column fromCol..toCol, plus the trailing right edge.
        private static double[] BandEdgesX(Dictionary<int, double> leftX, int fromCol, int toCol, PdfRange range, int repRow)
        {
            int n = toCol - fromCol + 1;
            var arr = new double[n + 1];
            for (int c = fromCol, k = 0; c <= toCol; c++, k++) arr[k] = leftX[c];
            arr[n] = leftX[toCol] + RangeColWidth(range, repRow, toCol);
            return arr;
        }

        // Row edges for a band region: top edge of every row fromRow..toRow, plus the trailing bottom edge.
        private static double[] BandEdgesY(Dictionary<int, double> topY, int fromRow, int toRow, PdfRange range)
        {
            int n = toRow - fromRow + 1;
            var arr = new double[n + 1];
            for (int r = fromRow, k = 0; r <= toRow; r++, k++) arr[k] = topY[r];
            arr[n] = topY[toRow] - RangeRowHeight(range, toRow);
            return arr;
        }

        private static PdfCell RangeCell(PdfRange range, int row, int col)
        {
            if (row < range.Range._fromRow || row > range.Range._toRow) return null;
            if (col < range.Range._fromCol || col > range.Range._toCol) return null;
            return range.Map[row, col];
        }

        private static double RangeColWidth(PdfRange range, int representativeRow, int col)
        {
            var cell = RangeCell(range, representativeRow, col);
            return cell?.ColumnWidth ?? 0d;
        }

        private static double RangeRowHeight(PdfRange range, int row)
        {
            int i = row - range.Range._fromRow;
            return (i >= 0 && i < range.RowHeights.Count) ? range.RowHeights[i].Height : 0d;
        }

        internal static Pages GetNumberOfPages(PdfPageSettings pageSettings, PdfWorksheet pdfSheet, ref PdfRange range)
        {
            var xPages = (int)Math.Max(1, Math.Ceiling(range.TotalWidth / pageSettings.ContentBounds.Width));
            var yPages = (int)Math.Max(1, Math.Ceiling(range.TotalHeight / pageSettings.ContentBounds.Height));

            if (pageSettings.ShowHeadings)
            {
                int prev = 0;
                do
                {
                    prev = xPages;
                    range.AdditionalWidth = xPages * ((rowHeadingWith1CharWidth - pdfSheet.ZeroCharWidth) + (Math.Abs(pdfSheet.ToRow).ToString().Length * pdfSheet.ZeroCharWidth));
                    xPages = (int)Math.Max(1, Math.Ceiling((range.TotalWidth + range.AdditionalWidth) / pageSettings.ContentBounds.Width));
                } while (prev != xPages);
                do
                {
                    prev = yPages;
                    range.AdditionalHeight = yPages * pdfSheet.Worksheet.DefaultRowHeight;
                    yPages = (int)Math.Max(1, Math.Ceiling((range.TotalHeight + range.AdditionalHeight) / pageSettings.ContentBounds.Height));
                } while (prev != yPages);
            }
            for (int i = range.Range._fromCol; i <= range.Range._toCol; i++)
            {
                if (pdfSheet.Worksheet.Column(i).PageBreak)
                    xPages++;
            }
            for (int i = range.Range._fromRow; i <= range.Range._toRow; i++)
            {
                if (pdfSheet.Worksheet.Row(i).PageBreak)
                    yPages++;
            }
            ComputePrintTitleDimensions(pdfSheet, range, out range.PrintTitleHeight, out range.PrintTitleWidth);
            range.PrintTitleRowTo = pdfSheet.PrintTitleRowFrom >= 0 ? pdfSheet.PrintTitleRowTo : -1;
            range.PrintTitleColTo = pdfSheet.PrintTitleColFrom >= 0 ? pdfSheet.PrintTitleColTo : -1;
            Pages p = new Pages();
            p.Width = xPages;
            p.Height = yPages;
            p.Page = null;
            return p;
        }

        internal static Pages AssignRangeToPages(PdfPageSettings pageSettings, PdfRange range, Pages pdfPages)
        {
            var pages = pdfPages;
            var worksheet = range.Range.Worksheet;
            var addedWidth = pages.Width > 0 ? range.AdditionalWidth / pages.Width : 0d;
            var addedHeight = pages.Height > 0 ? range.AdditionalHeight / pages.Height : 0d;

            var colSegments = GetColumnSegments(pageSettings, range, worksheet, addedWidth, range.PrintTitleWidth, range.PrintTitleColTo);
            var rowSegments = GetRowSegments(pageSettings, range, worksheet, addedHeight, range.PrintTitleHeight, range.PrintTitleRowTo);

            pages.Page = new Page[colSegments.Count * rowSegments.Count];
            int i = 0;

            if (pageSettings.PageOrders == PageOrders.DownThenOver)
            {
                for (int ci = 0; ci < colSegments.Count; ci++)
                    for (int ri = 0; ri < rowSegments.Count; ri++)
                        pages.Page[i++] = new Page
                        {
                            FromColumn = colSegments[ci].From,
                            ToColumn = colSegments[ci].To,
                            FromRow = rowSegments[ri].From,
                            ToRow = rowSegments[ri].To,
                            HeadingWidth = addedWidth,
                            HeadingHeight = addedHeight,
                            PrintTitleWidth = (range.PrintTitleColTo >= 0 && colSegments[ci].From > range.PrintTitleColTo) ? range.PrintTitleWidth : 0d,
                            PrintTitleHeight = (range.PrintTitleRowTo >= 0 && rowSegments[ri].From > range.PrintTitleRowTo) ? range.PrintTitleHeight : 0d,
                        };
            }
            else //if (pageSettings.PageOrders == PageOrders.OverThenDown)
            {
                for (int ri = 0; ri < rowSegments.Count; ri++)
                    for (int ci = 0; ci < colSegments.Count; ci++)
                        pages.Page[i++] = new Page
                        {
                            FromColumn = colSegments[ci].From,
                            ToColumn = colSegments[ci].To,
                            FromRow = rowSegments[ri].From,
                            ToRow = rowSegments[ri].To,
                            HeadingWidth = addedWidth,
                            HeadingHeight = addedHeight,
                            PrintTitleWidth = (range.PrintTitleColTo >= 0 && colSegments[ci].From > range.PrintTitleColTo) ? range.PrintTitleWidth : 0d,
                            PrintTitleHeight = (range.PrintTitleRowTo >= 0 && rowSegments[ri].From > range.PrintTitleRowTo) ? range.PrintTitleHeight : 0d,
                        };
            }
            pdfPages = pages;
            return pdfPages;
        }

        private struct PageSegment
        {
            public int From;
            public int To;
            public PageSegment(int from, int to) { From = from; To = to; }
        }

        private static List<PageSegment> GetColumnSegments(PdfPageSettings pageSettings, PdfRange range, ExcelWorksheet worksheet, double addedWidth, double titleWidth, int printTitleColTo)
        {
            var segments = new List<PageSegment>();
            int segStartIdx = 0;
            double width = 0d;
            for (int col = 0; col < range.ColWidths.Count; col++)
            {
                int actualCol = range.Range._fromCol + col;
                bool reserveTitle = titleWidth > 0d && printTitleColTo >= 0 && (range.Map.FromColumn + segStartIdx) > printTitleColTo;
                double effectiveAdded = addedWidth + (reserveTitle ? titleWidth : 0d);
                // Content-bounds overflow: col doesn't fit, end segment before it and reprocess.
                if (width + range.ColWidths[col] + effectiveAdded >= pageSettings.ContentBounds.Width)
                {
                    if(col == segStartIdx)
                    {
                        segments.Add(new PageSegment(range.Map.FromColumn + col, range.Map.FromColumn + col));
                        segStartIdx = col + 1;
                        width = 0d;
                        continue;
                    }
                    segments.Add(new PageSegment(range.Map.FromColumn + segStartIdx, range.Map.FromColumn + col - 1));
                    segStartIdx = col;
                    width = 0d;
                    col--; // reprocess this col as the first col of the next segment
                    continue;
                }
                width += range.ColWidths[col];
                // Explicit page break: col is included on this page, next segment starts after it.
                if (worksheet.Column(actualCol).PageBreak)
                {
                    segments.Add(new PageSegment(range.Map.FromColumn + segStartIdx, range.Map.FromColumn + col));
                    segStartIdx = col + 1;
                    width = 0d;
                }
            }
            // Remaining cols form the last segment.
            if (segStartIdx < range.ColWidths.Count)
                segments.Add(new PageSegment(range.Map.FromColumn + segStartIdx, range.Map.FromColumn + range.ColWidths.Count - 1));
            return segments;
        }

        private static List<PageSegment> GetRowSegments(PdfPageSettings pageSettings, PdfRange range, ExcelWorksheet worksheet, double addedHeight, double titleHeight, int printTitleRowTo)
        {
            var segments = new List<PageSegment>();
            int segStartIdx = 0;
            double height = 0d;
            for (int row = 0; row < range.RowHeights.Count; row++)
            {
                int actualRow = range.Range._fromRow + row;
                bool reserveTitle = titleHeight > 0d && printTitleRowTo >= 0 && (range.Map.FromRow + segStartIdx) > printTitleRowTo;
                double effectiveAdded = addedHeight + (reserveTitle ? titleHeight : 0d);
                // Content-bounds overflow: row doesn't fit, end segment before it and reprocess.
                if (height + range.RowHeights[row].Height + effectiveAdded >= pageSettings.ContentBounds.Height)
                {
                    segments.Add(new PageSegment(range.Map.FromRow + segStartIdx, range.Map.FromRow + row - 1));
                    segStartIdx = row;
                    height = 0d;
                    row--; // reprocess this row as the first row of the next segment
                    continue;
                }
                height += range.RowHeights[row].Height;
                // Explicit page break: row is included on this page, next segment starts after it.
                if (worksheet.Row(actualRow).PageBreak)
                {
                    segments.Add(new PageSegment(range.Map.FromRow + segStartIdx, range.Map.FromRow + row));
                    segStartIdx = row + 1;
                    height = 0d;
                }
            }
            // Remaining rows form the last segment.
            if (segStartIdx < range.RowHeights.Count)
                segments.Add(new PageSegment(range.Map.FromRow + segStartIdx, range.Map.FromRow + range.RowHeights.Count - 1));
            return segments;
        }

        internal static Pages MapPage(PdfRange range, Pages pdfPages)
        {
            var pages = pdfPages;
            for (int i = 0; i < pdfPages.Page.Length; i++)
            {
                var page = pdfPages.Page[i];
                page.Map = new PdfCellCollection(page.FromRow, page.ToRow, page.FromColumn, page.ToColumn);
                page.RowHeights = new double[page.ToRow - page.FromRow + 1];
                for (int row = page.FromRow; row <= page.ToRow; row++)
                {
                    int rangeIdx = row - range.Range._fromRow;
                    page.RowHeights[row - page.FromRow] = range.RowHeights[rangeIdx].Height;
                    for (int col = page.FromColumn; col <= page.ToColumn; col++)
                    {
                        page.Map[row, col] = range.Map[row, col];
                    }
                }
                pdfPages.Page[i] = page;
            }
            pdfPages = pages;
            return pdfPages;
        }

        private static Pages GetHeaderFooter(PdfRange range, Pages pdfPages, PdfWorksheet pdfSheet)
        {
            var pages = pdfPages;
            for (int i = 0; i < pdfPages.Page.Length; i++)
            {
                var page = pdfPages.Page[i];
                page.HeaderFooters = pdfSheet.HeaderFooters;
                pdfPages.Page[i] = page;
            }
            pdfPages = pages;
            return pdfPages;
        }

        private static void SubstitutePageNumbers(PdfPageSettings pageSettings, PdfDictionaries dictionaries, PdfHeaderFooter hf, int pageNumber, int totalPages)
        {
            if (hf == null) return;
            if (hf.PageNumberIndexes.Count > 0)
            {
                foreach (var idx in hf.PageNumberIndexes)
                    hf.Content.TextFragments[idx].Text = pageNumber.ToString();
            }
            if (hf.NumberOfPagesIndexes.Count > 0)
            {
                foreach (var idx in hf.NumberOfPagesIndexes)
                    hf.Content.TextFragments[idx].Text = totalPages.ToString();
            }
            PdfTextShaper.ShapeText(pageSettings, dictionaries, hf.Content);
        }

        private static void AddIncomingSpill(Page page, PdfRange range, int fromRow, int toRow, int windowFromCol, int windowToCol, double windowOriginX, double windowOriginY, bool isPrintTitle)
        {
            if (page.SpillCells == null) return; // initialised by the caller
            double windowRightX = windowOriginX;
            for (int c = windowFromCol; c <= windowToCol; c++) windowRightX += RangeColWidth(range, page.FromRow, c);
            double rowTop = windowOriginY;
            for (int r = fromRow; r <= toRow; r++)
            {
                double rowH = RangeRowHeight(range, r);
                double y = rowTop;
                rowTop -= rowH;
                if (rowH <= 0d) continue;
                // spill entering from the LEFT (left/general/center)
                double lx = windowOriginX;
                for (int c = windowFromCol - 1; c >= range.Range._fromCol; c--)
                {
                    double w = RangeColWidth(range, page.FromRow, c);
                    lx -= w;
                    var cell = RangeCell(range, r, c);
                    if (cell == null) continue;
                    if (cell.Merged) break;                              // merges don't spill, and block
                    if (string.IsNullOrEmpty(cell.Text)) continue;
                    var hal = cell.ContentAligmnet?.HorizontalAlignment ?? (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.General;
                    double rightExtent =
                        (hal == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Center) ? lx + w / 2d + cell.TotalTextLength / 2d :
                        (hal == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Right) ? lx + w : lx + cell.TotalTextLength;
                    if (rightExtent > windowOriginX)
                    {
                        double clipRight = FirstBlockedX(page, range, r, windowFromCol, windowToCol, windowOriginX, windowRightX, fromLeft: true);
                        page.SpillCells.Add(new SpillCellDraw
                        {
                            Cell = cell,
                            X = lx,
                            Y = y,
                            Width = w,
                            Height = rowH,
                            ClipX = windowOriginX,
                            ClipY = y,
                            ClipWidth = clipRight - windowOriginX,
                            ClipHeight = rowH,
                            IsPrintTitle = isPrintTitle
                        });
                    }
                    break;
                }
                // spill entering from the RIGHT (right/center)
                double rx = windowRightX;
                for (int c = windowToCol + 1; c <= range.Range._toCol; c++)
                {
                    double w = RangeColWidth(range, page.FromRow, c);
                    var cell = RangeCell(range, r, c);
                    if (cell == null) { rx += w; continue; }
                    if (cell.Merged) break;
                    if (string.IsNullOrEmpty(cell.Text)) { rx += w; continue; }
                    var hal = cell.ContentAligmnet?.HorizontalAlignment ?? (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.General;
                    double leftExtent =
                        (hal == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Center) ? rx + w / 2d - cell.TotalTextLength / 2d :
                        (hal == (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Right) ? rx + w - cell.TotalTextLength : rx;
                    if (leftExtent < windowRightX)
                    {
                        double clipLeft = FirstBlockedX(page, range, r, windowFromCol, windowToCol, windowOriginX, windowRightX, fromLeft: false);
                        page.SpillCells.Add(new SpillCellDraw
                        {
                            Cell = cell,
                            X = rx,
                            Y = y,
                            Width = w,
                            Height = rowH,
                            ClipX = clipLeft,
                            ClipY = y,
                            ClipWidth = windowRightX - clipLeft,
                            ClipHeight = rowH,
                            IsPrintTitle = isPrintTitle
                        });
                    }
                    break;
                }
            }
        }

        // Where spill into the window is cut off by the first non-empty (or merged) cell inside it.
        private static double FirstBlockedX(Page page, PdfRange range, int row, int windowFromCol, int windowToCol, double windowOriginX, double windowRightX, bool fromLeft)
        {
            if (fromLeft)
            {
                double x = windowOriginX;
                for (int c = windowFromCol; c <= windowToCol; c++)
                {
                    var cell = RangeCell(range, row, c);
                    if (cell != null && (cell.Merged || !string.IsNullOrEmpty(cell.Text))) return x;
                    x += RangeColWidth(range, page.FromRow, c);
                }
                return windowRightX;
            }
            double rx = windowRightX;
            for (int c = windowToCol; c >= windowFromCol; c--)
            {
                double w = RangeColWidth(range, page.FromRow, c);
                rx -= w;
                var cell = RangeCell(range, row, c);
                if (cell != null && (cell.Merged || !string.IsNullOrEmpty(cell.Text))) return rx + w;
            }
            return windowOriginX;
        }

        internal static Pages PrecomputeSpillCells(PdfPageSettings pageSettings, PdfRange range, Pages pdfPages)
        {
            for (int i = 0; i < pdfPages.Page.Length; i++)
            {
                var page = pdfPages.Page[i];
                page.SpillCells = new List<SpillCellDraw>();
                double originX = pageSettings.ContentBounds.Left + page.HeadingWidth + page.PrintTitleWidth;
                double originY = pageSettings.ContentBounds.Top - page.HeadingHeight - page.PrintTitleHeight;
                AddIncomingSpill(page, range, page.FromRow, page.ToRow, page.FromColumn, page.ToColumn, originX, originY, isPrintTitle: false);
                pdfPages.Page[i] = page;
            }
            return pdfPages;
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

        private static void EmitBandFrameH(List<GridLine> target, PdfRange range, double y, int row, int fromCol, int nc, double[] colX, Func<PdfCell, bool> hasBorder)
        {
            double? rs = null; double re = 0d;
            for (int ci = 0; ci < nc; ci++)
            {
                double segL = colX[ci], segR = colX[ci + 1];
                if (segR - segL <= 0d) { if (rs != null) { target.Add(new GridLine(rs.Value, y, re, y)); rs = null; } continue; }
                if (!hasBorder(RangeCell(range, row, fromCol + ci))) { if (rs == null) rs = segL; re = segR; }
                else if (rs != null) { target.Add(new GridLine(rs.Value, y, re, y)); rs = null; }
            }
            if (rs != null) target.Add(new GridLine(rs.Value, y, re, y));
        }

        private static void EmitBandFrameV(List<GridLine> target, PdfRange range, double x, int col, int fromRow, int nr, double[] rowY, Func<PdfCell, bool> hasBorder)
        {
            double? rs = null; double re = 0d;
            for (int ri = 0; ri < nr; ri++)
            {
                double segT = rowY[ri], segB = rowY[ri + 1];
                if (segT - segB <= 0d) { if (rs != null) { target.Add(new GridLine(x, rs.Value, x, re)); rs = null; } continue; }
                if (!hasBorder(RangeCell(range, fromRow + ri, col))) { if (rs == null) rs = segT; re = segB; }
                else if (rs != null) { target.Add(new GridLine(x, rs.Value, x, re)); rs = null; }
            }
            if (rs != null) target.Add(new GridLine(x, rs.Value, x, re));
        }
    }
}

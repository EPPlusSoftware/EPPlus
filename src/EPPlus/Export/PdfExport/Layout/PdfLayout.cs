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
using EPPlus.Export.Pdf.Layout;
using EPPlus.Export.Pdf.Resources;
using EPPlus.Export.Pdf.Settings;
using EPPlus.Graphics;
using OfficeOpenXml.Export.PdfExport.Data;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Export.PdfExport.Layout
{
    internal static class PdfLayout
    {
        internal const double rowHeadingWith1CharWidth = 18d;

        public static Transform GetLayout(PdfPageSettings[] sheetSettings, PdfDictionaries dictionaries, PdfWorksheet[] pdfSheets)
        {
            var PagesCollection = GetPages(sheetSettings, pdfSheets);
            var Catalog = GetCatalog(sheetSettings[0].FirstPageNumber, dictionaries, PagesCollection);
            if (Catalog is PdfCatalogLayout c)
            {
                PdfBookmarkCollector.CollectNamedDestinations(pdfSheets, c.NamedDestinations);
            }
            return Catalog;
        }

        internal static Transform GetCatalog(int firstPageNumber, PdfDictionaries dictionaries, List<Pages> pdfPages)
        {
            Transform Catalog = new PdfCatalogLayout(0d, 0d, 0d, 0d);
            int totalPages = 0;
            var sheetTotalPages = GetTotalPagesPerSheet(pdfPages);
            int displayedPageNumber = 0;
            int physicalPageIndex = 1;
            int currentSheetIndex = -1;
            for (int i = 0; i < pdfPages.Count; i++)
            {
                var pageSettings = pdfPages[i].Settings;
                bool isDraft = pageSettings.Draft;
                bool isBlackWhite = pageSettings.BlackAndWhite;
                if (pdfPages[i].SheetIndex != currentSheetIndex)
                {
                    currentSheetIndex = pdfPages[i].SheetIndex;
                    displayedPageNumber = pageSettings.FirstPageNumber;
                    physicalPageIndex = 1;
                    sheetTotalPages.TryGetValue(currentSheetIndex, out totalPages);
                }
                var pages = pdfPages[i].Page;
                for (int j = 0; j < pages.Length; j++)
                {
                    var page = pages[j];
                    var pageLayout = CreatePageLayout(pdfPages[i], page, pageSettings, dictionaries, out double contentStartX, out double contentStartY);
                    CreateCellLayouts(pageLayout, page, pageSettings, dictionaries, contentStartX, contentStartY, isDraft, isBlackWhite);
                    PdfImagePlacement.EmitPageImages(pageLayout, page, isDraft);
                    PdfHeaderFooterLayout.AddHeaderFooter(pageLayout, page, pageSettings, dictionaries, displayedPageNumber, physicalPageIndex, totalPages);
                    AddAdditionalContent(pageLayout, page, pageSettings, dictionaries);
                    if (!isDraft)
                    {
                        PdfGridlinesLayout.AddGridLines(pageSettings, page, pageLayout, borderOnly: !pageSettings.ShowGridLines || pdfPages[i].IsCommentsPage);
                    }
                    pageLayout.ChildObjects.Sort((a, b) =>
                    {
                        int cmp = a.Z.CompareTo(b.Z);
                        return cmp == 0 ? string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase) : cmp;
                    });
                    displayedPageNumber++;
                    physicalPageIndex++;
                    Catalog.AddChild(pageLayout);
                }
            }
            return Catalog;
        }

        private static PdfPageLayout CreatePageLayout(Pages sheet, Page page, PdfPageSettings pageSettings, PdfDictionaries dictionaries, out double contentStartX, out double contentStartY)
        {
            var pageLayout = new PdfPageLayout(0d, 0d, 0d, 0d);
            pageLayout.Settings = pageSettings;
            pageLayout.isCommentsPage = sheet.IsCommentsPage;
            pageLayout.HeadingWidth = page.HeadingWidth;
            pageLayout.HeadingHeight = page.HeadingHeight;
            pageLayout.PrintTitleWidth = page.PrintTitleWidth;
            pageLayout.PrintTitleHeight = page.PrintTitleHeight;
            contentStartX = GetOriginX(pageSettings, page) + page.HeadingWidth + page.PrintTitleWidth;
            contentStartY = GetOriginY(pageSettings, page) - page.HeadingHeight - page.PrintTitleHeight;
            if (pageSettings.ShowHeadings && !sheet.IsCommentsPage)
            {
                PdfHeadingLayout.AddHeadingCells(pageSettings, dictionaries, page, pageLayout, contentStartX, contentStartY, page.HeadingWidth, page.HeadingHeight, sheet.HeadingFontName, sheet.HeadingFontSize, sheet.HeadingFill);
                PdfHeadingLayout.AddPrintTitleHeadings(pageSettings, dictionaries, page, pageLayout, sheet.HeadingFontName, sheet.HeadingFontSize, sheet.HeadingFill);
            }
            PdfSpillLayout.AddSpillCells(pageSettings, dictionaries, page, pageLayout);
            PdfHeadingLayout.AddPrintTitleCells(pageSettings, dictionaries, page, pageLayout);
            if (!sheet.IsCommentsPage)
            {
                pageLayout.SheetName = sheet.SheetName;
                pageLayout.FromRow = page.FromRow;
                pageLayout.ToRow = page.ToRow;
                pageLayout.FromColumn = page.FromColumn;
                pageLayout.ToColumn = page.ToColumn;
                pageLayout.ColLeftX = BuildColumnXPositions(pageSettings, page);
                var rowTopY = new double[page.ToRow - page.FromRow + 1];
                double ry = contentStartY;
                for (int r = page.FromRow; r <= page.ToRow; r++)
                {
                    rowTopY[r - page.FromRow] = ry;
                    ry -= page.RowHeights[r - page.FromRow];
                }
                pageLayout.RowTopY = rowTopY;
            }
            else
            {
                pageLayout.CommentsSourceSheet = sheet.SheetName;
            }
            return pageLayout;
        }

        private static void CreateCellLayouts(PdfPageLayout pageLayout, Page page, PdfPageSettings pageSettings, PdfDictionaries dictionaries, double contentStartX, double contentStartY, bool isDraft, bool isBlackWhite)
        {
            var drawnMergedCells = new HashSet<string>();
            double y = contentStartY;
            double x = contentStartX;
            for (int row = page.FromRow; row <= page.ToRow; row++)
            {
                double rowHeight = page.RowHeights[row - page.FromRow];
                for (int col = page.FromColumn; col <= page.ToColumn; col++)
                {
                    var map = page.Map[row, col];
                    MergedCellDrawInfo info = new MergedCellDrawInfo();
                    if (map.Hidden && !map.Merged)
                    {
                        x += map.ColumnWidth;
                        continue;
                    }
                    if (map.Merged)
                    {
                        string key = map.MergedAddress.Address;
                        if (!drawnMergedCells.Contains(key) && page.MergedCells.TryGetValue(key, out info))
                        {
                            CreateMergedCellLayout(pageLayout, map, info, pageSettings, dictionaries, isDraft, isBlackWhite);
                            drawnMergedCells.Add(key);
                        }
                    }
                    else
                    {
                        CreateCellLayout(pageLayout, page, map, pageSettings, dictionaries, x, y, rowHeight, row, col, isDraft, isBlackWhite);
                    }
                    if (!isDraft)
                    {
                        CreateCellBorderLayout(pageLayout, page, map, info, x, y, rowHeight, row, col, isBlackWhite);
                    }
                    x += map.ColumnWidth;
                }
                y -= rowHeight;
                x = contentStartX;
            }
        }

        private static bool CellHasText(PdfCell cell) => cell?.TextLines != null && cell.TextLines.Count > 0;

        private static bool CellHasPicture(PdfCell cell) => cell?.CellPictureBytes != null;

        private static void CreateMergedCellLayout(PdfPageLayout pageLayout, PdfCell map, MergedCellDrawInfo info, PdfPageSettings pageSettings, PdfDictionaries dictionaries, bool isDraft, bool isBlackWhite)
        {
            if (!isDraft)
            {
                if (!isBlackWhite)
                {
                    CreateCellFillLayout(pageLayout, map.Main?.CellStyle ?? map.CellStyle, map.Text, info.X, info.Y, info.Width, info.Height, map.Name, pageSettings, dictionaries);
                }
                var picMap = CellHasPicture(map) ? map : CellHasPicture(map.Main) ? map.Main : null;
                if (picMap != null)
                {
                    PdfImagePlacement.AddCellPicture(pageLayout, picMap, info.X, info.Y, info.Width, info.Height);
                }
            }
            var sourceMap = CellHasText(map) ? map : CellHasText(map.Main) ? map.Main : null;
            if (sourceMap != null)
            {
                var text = new PdfCellContentLayout(pageSettings, dictionaries, sourceMap, info.X, info.Y, info.Width, info.Height);
                text.Name = map.Name;
                text.GidsAndCharMap(dictionaries);
                text.SetupClipping(info.X, info.Y, info.Width, info.Height);
                pageLayout.AddChild(text);
                CreateTextLinks(pageLayout, text, map, info.X, info.X + info.Width, info.X, info.Y, info.Width, info.Height);
            }
            else if (map.HyperlinkUri != null || map.HyperlinkTargetSheet != null)
            {
                pageLayout.Links.Add(PdfLinkLayout.CreateLinkDrawInfo(info.X, info.Y, info.Width, info.Height, map));
            }
            if (!isDraft && map.Main != null)
            {
                var mergeMainStyle = map.Main.CellStyle;
                if (PdfBorderResolver.HasDiagonalBorder(mergeMainStyle))
                {
                    var diagBorder = new PdfCellBorderLayout(isMerged: false, corners: MergedCellCorners.All, info: info, x: info.X, y: info.Y, width: info.Width, height: info.Height);
                    PdfBorderResolver.SetBorderStyle(mergeMainStyle, diagBorder, isBlackWhite);
                    diagBorder.Name = map.Name;
                    diagBorder.BorderData.Top.BorderStyle = (EPPlus.Export.Pdf.Enums.ExcelBorderStyle)ExcelBorderStyle.None;
                    diagBorder.BorderData.Bottom.BorderStyle = (EPPlus.Export.Pdf.Enums.ExcelBorderStyle)ExcelBorderStyle.None;
                    diagBorder.BorderData.Left.BorderStyle = (EPPlus.Export.Pdf.Enums.ExcelBorderStyle)ExcelBorderStyle.None;
                    diagBorder.BorderData.Right.BorderStyle = (EPPlus.Export.Pdf.Enums.ExcelBorderStyle)ExcelBorderStyle.None;
                    pageLayout.AddChild(diagBorder);
                }
            }
        }

        private static void CreateCellLayout(PdfPageLayout pageLayout, Page page, PdfCell map, PdfPageSettings pageSettings, PdfDictionaries dictionaries, double x, double y, double rowHeight, int row, int col, bool isDraft, bool isBlackWhite)
        {
            var effectiveWidth = GetClampedCellWidth(pageSettings, x, map.ColumnWidth);
            if (!isDraft)
            {
                if (!isBlackWhite)
                {
                    CreateCellFillLayout(pageLayout, map.CellStyle, map.Text, x, y, effectiveWidth, rowHeight, map.Name, pageSettings, dictionaries);
                }
                PdfImagePlacement.AddCellPicture(pageLayout, map, x, y, effectiveWidth, rowHeight);
            }
            if (map.TextLines != null && map.TextLines.Count > 0)
            {
                var text = new PdfCellContentLayout(pageSettings, dictionaries, map, x, y, effectiveWidth, rowHeight);
                text.Name = map.Name;
                text.GidsAndCharMap(dictionaries);
                var (clipX, clipW, spillLeft, spillRight) = PdfSpillLayout.GetOverflowClip(map, page, row, col, x, effectiveWidth);
                text.SetupClipping(clipX, y, clipW, rowHeight, spillLeft, spillRight);
                pageLayout.AddChild(text);
                CreateTextLinks(pageLayout, text, map, clipX, clipX + clipW, x, y, effectiveWidth, rowHeight);
            }
            else if (map.HyperlinkUri != null || map.HyperlinkTargetSheet != null)
            {
                pageLayout.Links.Add(PdfLinkLayout.CreateLinkDrawInfo(x, y, effectiveWidth, rowHeight, map));
            }
        }

        private static void CreateCellFillLayout(PdfPageLayout pageLayout, PdfCellStyle style, string text, double x, double y, double width, double height, string name, PdfPageSettings pageSettings, PdfDictionaries dictionaries)
        {
            var fill = new PdfCellLayout(x, y, width, height);
            PdfFillResolver.SetFill(dictionaries, style, text, fill);
            fill.UpdateShadingPositionMatrix(pageSettings);
            fill.Name = name;
            pageLayout.AddChild(fill);
        }

        private static void CreateTextLinks(PdfPageLayout pageLayout, PdfCellContentLayout text, PdfCell map, double clipLeft, double clipRight, double boxX, double boxY, double boxWidth, double boxHeight)
        {
            if (map.HyperlinkUri == null && map.HyperlinkTargetSheet == null) return;
            if (PdfLinkLayout.CanTightenLink(text))
            {
                PdfLinkLayout.AddTextLinks(pageLayout, text, clipLeft, clipRight, map.HyperlinkUri, map.HyperlinkTargetSheet, map.HyperlinkTargetRow, map.HyperlinkTargetCol, map.HyperlinkToolTip);
            }
            else
            {
                pageLayout.Links.Add(PdfLinkLayout.CreateLinkDrawInfo(boxX, boxY, boxWidth, boxHeight, map));
            }
        }

        private static void CreateCellBorderLayout(PdfPageLayout pageLayout, Page page, PdfCell map, MergedCellDrawInfo info, double x, double y, double rowHeight, int row, int col, bool isBlackWhite)
        {
            if (!PdfBorderResolver.HasBorder(map.CellStyle)) return;
            var border = new PdfCellBorderLayout(map.Merged, GetCorners(map.MergedAddress, row, col), info, x, y, map.ColumnWidth, rowHeight);
            PdfBorderResolver.SetBorderStyle(map.CellStyle, border, isBlackWhite);
            border.Name = map.Name;
            if (map.Merged && map.MergedAddress != null)
            {
                var addr = map.MergedAddress;
                if (row != addr.Start.Row) border.BorderData.Top.BorderStyle = (EPPlus.Export.Pdf.Enums.ExcelBorderStyle)ExcelBorderStyle.None;
                if (row != addr.End.Row) border.BorderData.Bottom.BorderStyle = (EPPlus.Export.Pdf.Enums.ExcelBorderStyle)ExcelBorderStyle.None;
                if (col != addr.Start.Column) border.BorderData.Left.BorderStyle = (EPPlus.Export.Pdf.Enums.ExcelBorderStyle)ExcelBorderStyle.None;
                if (col != addr.End.Column) border.BorderData.Right.BorderStyle = (EPPlus.Export.Pdf.Enums.ExcelBorderStyle)ExcelBorderStyle.None;
            }
            PdfBorderResolver.SetDoubleBorderMiterFlags(page, row, col, border);
            pageLayout.AddChild(border);
        }

        private static void AddAdditionalContent(PdfPageLayout pageLayout, Page page, PdfPageSettings pageSettings, PdfDictionaries dictionaries)
        {
            if (page.AddedContent == null || page.AddedContent.content?.TextLines == null || page.AddedContent.content.TextLines.Count == 0) return;
            var content = page.AddedContent.content;
            int last = content.TextLines.Count - 1;
            var descent = content.TextLines[last].LargestDescent;
            var ascent = content.TextLines[last].LargestAscent;
            var sx = 10d;
            var sy = descent + 2d;
            var width = pageSettings.PageSize.WidthPu - pageSettings.Margins.LeftPu - pageSettings.Margins.RightPu;
            var cellcont = new PdfCellContentLayout(pageSettings, dictionaries, content, sx, sy, width, 0);
            cellcont.IsAddedContent = true;
            cellcont.GidsAndCharMap(dictionaries);
            pageLayout.AddChild(cellcont);
            sy = pageSettings.PageSize.HeightPu - (descent + ascent + 4d);
            cellcont = new PdfCellContentLayout(pageSettings, dictionaries, content, sx, sy, width, 0);
            cellcont.IsAddedContent = true;
            cellcont.GidsAndCharMap(dictionaries);
            pageLayout.AddChild(cellcont);
        }

        private static Dictionary<int, int> GetTotalPagesPerSheet(List<Pages> pdfPages)
        {
            var totals = new Dictionary<int, int>();
            for (int i = 0; i < pdfPages.Count; i++)
            {
                int si = pdfPages[i].SheetIndex;
                if (!totals.ContainsKey(si)) totals[si] = 0;
                totals[si] += pdfPages[i].Page.Length;
            }
            return totals;
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

        internal static List<Pages> GetPages(PdfPageSettings[] sheetSettings, PdfWorksheet[] pdfSheets)
        {
            List<Pages> PagesCollection = new List<Pages>();
            for (int si = 0; si < pdfSheets.Length; si++)
            {
                var pdfSheet = pdfSheets[si];
                var pageSettings = sheetSettings[si];
                if (pdfSheet.Ranges.Count > 0)
                {
                    PdfPaginator.ResolveFitScale(pageSettings, pdfSheet, pdfSheet.Ranges[0]);
                }
                for (int ri = 0; ri < pdfSheet.Ranges.Count; ri++)
                {
                    var range = pdfSheet.Ranges[ri];
                    var pages = PdfPaginator.GetNumberOfPages(pageSettings, pdfSheet, ref range);
                    pages = PdfPaginator.AssignRangeToPages(pageSettings, range, pages);
                    pages = PdfPaginator.MapPage(range, pages);
                    pages = PdfHeaderFooterLayout.GetHeaderFooter(range, pages, pdfSheet);
                    pages = GetAdditionalContent(pages, pdfSheet);
                    pages = PdfMergedCellLayout.PrecomputeMergedCells(pageSettings, range, pages);
                    pages = PdfSpillLayout.PrecomputeSpillCells(pageSettings, range, pages);
                    pages = PdfHeadingLayout.PrecomputePrintTitleCells(pageSettings, pdfSheet, range, pages);
                    pages = PdfImagePlacement.PrecomputeImages(pageSettings, range, pages, pdfSheet.Drawings, pdfSheet.ZeroCharWidth);
                    pages.HeadingFontName = pdfSheet.NormalStyle.Style.Font.Name;
                    pages.HeadingFontSize = pdfSheet.NormalStyle.Style.Font.Size;
                    pages.HeadingFill = pdfSheet.NormalStyle.Style.Fill;
                    pages.Settings = pageSettings;
                    pages.SheetIndex = si;
                    pages.SheetName = pdfSheet.Worksheet.Name;
                    PagesCollection.Add(pages);
                }
                if (pdfSheet.CommentsAndNotes.Range != null)
                {
                    bool savedShowHeadings = pageSettings.ShowHeadings;
                    pageSettings.ShowHeadings = false;
                    var pages = PdfPaginator.GetNumberOfPages(pageSettings, pdfSheet, ref pdfSheet.CommentsAndNotes);
                    pages = PdfPaginator.AssignRangeToPages(pageSettings, pdfSheet.CommentsAndNotes, pages);
                    pages = PdfPaginator.MapPage(pdfSheet.CommentsAndNotes, pages);
                    pages = PdfHeaderFooterLayout.GetHeaderFooter(pdfSheet.CommentsAndNotes, pages, pdfSheet);
                    pages = GetAdditionalContent(pages, pdfSheet);
                    pageSettings.ShowHeadings = savedShowHeadings;
                    pages.IsCommentsPage = true;
                    pages.Settings = pageSettings;
                    pages.SheetIndex = si;
                    pages.SheetName = pdfSheet.Worksheet.Name;
                    pages.IsCommentsPage = true;
                    for (int k = 0; k < pages.Page.Length; k++)
                    {
                        pages.Page[k].IsCommentsPage = true;
                    }
                    PagesCollection.Add(pages);
                }
            }
            return PagesCollection;
        }

        internal static double[] BuildColumnXPositions(PdfPageSettings pageSettings, Page page)
        {
            int colCount = page.ToColumn - page.FromColumn + 1;
            var colX = new double[colCount];
            double x = GetOriginX(pageSettings, page) + page.HeadingWidth + page.PrintTitleWidth;
            for (int col = page.FromColumn; col <= page.ToColumn; col++)
            {
                colX[col - page.FromColumn] = x;
                var cell = page.Map[page.FromRow, col];
                x += cell?.ColumnWidth ?? 0d;
            }
            return colX;
        }

        private static Pages GetAdditionalContent(Pages pdfPages, PdfWorksheet pdfSheet)
        {
            for (int i = 0; i < pdfPages.Page.Length; i++)
            {
                var page = pdfPages.Page[i];
                page.AddedContent = pdfSheet.AdditionalContent;
                pdfPages.Page[i] = page;
            }
            return pdfPages;
        }

        /// <summary>
        /// The X coordinate where the page's printed block begins.
        /// Currently the left content bound; will include the centering offset.
        /// </summary>
        internal static double GetOriginX(PdfPageSettings pageSettings, Page page)
        {
            if (!pageSettings.CenterOnPageHorizontally) return pageSettings.ContentBounds.Left;
            var contentWidth = page.IsCommentsPage ? pageSettings.ContentBounds.Width : pageSettings.EffectiveContentWidth;
            var offset = (contentWidth - page.UsedWidth) / 2d;
            return pageSettings.ContentBounds.Left + Math.Max(0d, offset);
        }

        /// <summary>
        /// The Y coordinate where the page's printed block begins (top edge).
        /// </summary>
        internal static double GetOriginY(PdfPageSettings pageSettings, Page page)
        {
            if (!pageSettings.CenterOnPageVertically) return pageSettings.ContentBounds.Top;
            var contentHeight = page.IsCommentsPage ? pageSettings.ContentBounds.Height : pageSettings.EffectiveContentHeight;
            var offset = (contentHeight - page.UsedHeight) / 2d;
            return pageSettings.ContentBounds.Top - Math.Max(0d, offset);
        }

        internal static double GetClampedCellWidth(PdfPageSettings pageSettings, double cellX, double cellWidth)
        {
            double s = pageSettings.ContentScale;
            double pageRightNatural = pageSettings.ContentBounds.Left + (pageSettings.PageSize.WidthPu - pageSettings.ContentBounds.Left) / s;
            return System.Math.Min(cellWidth, pageRightNatural - cellX);
        }
    }
}

using EPPlus.Export.Pdf.PdfLayout;
using EPPlus.Export.Pdf.PdfResources;
using EPPlus.Export.Pdf.PdfSettings;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.Integration.DataHolders;
using EPPlus.Graphics;
using OfficeOpenXml;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;

namespace EPPlus.Export.Pdf.PdfCatalog
{

    internal struct MergedCellDrawInfo
    {
        public double X;
        public double Y;
        public double Width;
        public double Height;
    }

    internal struct Page
    {
        public int FromRow;
        public int FromColumn;
        public int ToRow;
        public int ToColumn;

        public bool HasPrintTitle;
        public double PrintTitleWidth;
        public double PrintTitleHeight;

        public PdfCellCollection Map;

        public PdfHeaderFooterCollection HeaderFooters;

        public Dictionary<string, MergedCellDrawInfo> MergedCells;

        public double[] RowHeights;

        public double HeadingWidth;
        public double HeadingHeight;

    }

    internal struct Pages
    {
        public Page[] Page;
        public int Width;
        public int Height;
        public bool IsCommentsPage;
        public string HeadingFontName;
        public float HeadingFontSize;
        public ExcelFill HeadingFill;
        public int Count
        {
            get { return Width * Height; }
        }
    }

    internal class PdfLayout
    {
        private const double rowHeadingWith1CharWidth = 23.25d;

        public static Transform GetLayout(PdfPageSettings pageSettings, PdfDictionaries dictionaries, PdfWorksheet[] pdfSheets)
        {
            var PagesCollection = GetPages(pageSettings, pdfSheets);
            var Catalog = GetCatalog(pageSettings, dictionaries, PagesCollection);
            return Catalog;
        }

        /*Header and footer notes:
         * First header/footer is only used on the first worksheets first page if it exsists
         * Then we use each worksheets odd and even respectively. worksheet 1 has its set of odd and even we use and worksheet 2 has it's own set we will use
         * Page number does not reset and total number of pages is across all worksheets pages
         * starting page number does not affect if first header/footer is used or not.
        */

        internal static Transform GetCatalog(PdfPageSettings pageSettings, PdfDictionaries dictionaries, List<Pages> pdfPages)
        {
            Transform Catalog = new Transform(0d, 0d, 0d, 0d);
            int totalPages = GetTotalPages(pdfPages);
            for (int i = 0; i < pdfPages.Count; i++)
            {

                var pages = pdfPages[i].Page;
                int pageNumber = pageSettings.FirstPageNumber;

                for (int j = 0; j < pages.Length; j++)
                {
                    var page = pages[j];
                    PdfPageLayout pageLayout = new PdfPageLayout(0d, 0d, 0d, 0d);
                    pageLayout.isCommentsPage = pdfPages[i].IsCommentsPage;
                    pageLayout.HeadingWidth = page.HeadingWidth;
                    pageLayout.HeadingHeight = page.HeadingHeight;
                    var drawnMergedCells = new HashSet<string>();
                    //var drawnMergedCellsText = new HashSet<string>();
                    //PdfContentLayout contentLayout = new PdfContentLayout(0d, 0d, pageSettings.ContentBounds);
                    //pageLayout.AddChild(contentLayout);
                    double contentStartX = pageSettings.ContentBounds.Left + page.HeadingWidth;
                    double contentStartY = pageSettings.ContentBounds.Top - page.HeadingHeight;
                    if (pageSettings.ShowHeadings && !pdfPages[i].IsCommentsPage)
                        AddHeadingCells(pageSettings, dictionaries, page, pageLayout, contentStartX, contentStartY, page.HeadingWidth, page.HeadingHeight, pdfPages[i].HeadingFontName, pdfPages[i].HeadingFontSize, pdfPages[i].HeadingFill);

                    double y = contentStartY;//pageSettings.ContentBounds.Top;
                    double x = contentStartX;//pageSettings.ContentBounds.Left;
                    //create cells & headings if exsists
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
                                    var fill = new PdfCellLayout(dictionaries, cellStyle,
                                        info.X, info.Y, info.Width, info.Height);
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
                                                mergeMainStyle,
                                                isMerged: false,            // use X/Y/W/H path in renderer, not info.*
                                                corners: MergedCellCorners.All,
                                                info: info,
                                                x: info.X,           // virtual full-merge top-left X
                                                y: info.Y,           // virtual full-merge top Y
                                                width: info.Width,       // full merge width
                                                height: info.Height);     // full merge height
                                            diagBorder.Name = map.Name;
                                            // Suppress edge borders — this layout exists only for the diagonal
                                            diagBorder.BorderData.Top.BorderStyle = ExcelBorderStyle.None;
                                            diagBorder.BorderData.Bottom.BorderStyle = ExcelBorderStyle.None;
                                            diagBorder.BorderData.Left.BorderStyle = ExcelBorderStyle.None;
                                            diagBorder.BorderData.Right.BorderStyle = ExcelBorderStyle.None;
                                            pageLayout.AddChild(diagBorder);
                                        }
                                    }
                                    drawnMergedCells.Add(key);
                                }
                            }
                            else
                            {
                                //Fill
                                var fill = new PdfCellLayout(dictionaries, map.CellStyle, x, y, map.ColumnWidth, rowHeight);
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
                                var border = new PdfCellBorderLayout(map.CellStyle, map.Merged, GetCorners(map.MergedAddress, row, col), info, x, y, map.ColumnWidth, rowHeight);
                                border.Name = map.Name;
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
                        if(centerH != null)
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
                        if(rightH != null)
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
                        if(leftF != null)
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
                        if(centerF != null)
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
                        if(rightF != null)
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

                    //Print titles
                    pageNumber++;
                    Catalog.AddChild(pageLayout);
                }
            }
            return Catalog;
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
            if (map.ContentAligmnet.HorizontalAlignment == ExcelHorizontalAlignment.Fill || map.ContentAligmnet.WrapText)
                return true;
            if (map.TotalTextLength <= map.ColumnWidth) return false;
            var halign = map.ContentAligmnet.HorizontalAlignment;
            if (halign == ExcelHorizontalAlignment.Left || halign == ExcelHorizontalAlignment.General)
            {
                // Text spills right — clip if the right neighbour has content or we're at the page edge.
                if (col >= page.ToColumn) return true;
                var right = page.Map[row, col + 1];
                return right != null && !string.IsNullOrEmpty(right.Text);
            }
            else if (halign == ExcelHorizontalAlignment.Right)
            {
                // Text spills left — clip if the left neighbour has content or we're at the page edge.
                if (col <= page.FromColumn) return true;
                var left = page.Map[row, col - 1];
                return left != null && !string.IsNullOrEmpty(left.Text);
            }
            else if (halign == ExcelHorizontalAlignment.Center)
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
            if(cellStyle == null) return false;
            bool hasBorders =
                cellStyle.xfTop.Style != ExcelBorderStyle.None ||
                cellStyle.xfBottom.Style != ExcelBorderStyle.None ||
                cellStyle.xfLeft.Style != ExcelBorderStyle.None ||
                cellStyle.xfRight.Style != ExcelBorderStyle.None ||
                (cellStyle.dxfTop?.HasValue ?? false) ||
                (cellStyle.dxfBottom?.HasValue ?? false) ||
                (cellStyle.dxfLeft?.HasValue ?? false) ||
                (cellStyle.dxfRight?.HasValue ?? false) ||
                (cellStyle.Diagonal != null && cellStyle.Diagonal.Style != ExcelBorderStyle.None);
            return hasBorders;
        }

        private static bool HasDiagonalBorder(PdfCellStyle style) => style?.Diagonal != null && style.Diagonal.Style != ExcelBorderStyle.None;

        private static void AddHeadingCells(PdfPageSettings pageSettings, PdfDictionaries dictionaries, Page page, PdfPageLayout pageLayout, double contentStartX, double contentStartY, double headingWidth, double headingHeight, string fontName, float fontSize, ExcelFill fill)
        {
            // An empty style gives a transparent/white fill with no border.
            // TODO: replace with a proper light-grey fill once PdfCellStyle supports
            //       creating standalone fills without a backing ExcelFillXml object.
            var headingStyle = new PdfCellStyle();
            headingStyle.xfFill = fill;

            // ── Corner cell (top-left, where the two strips intersect) ──────────────────────────
            var cornerFill = new PdfCellLayout(dictionaries, headingStyle,
                pageSettings.ContentBounds.Left, pageSettings.ContentBounds.Top,
                headingWidth, headingHeight);
            cornerFill.Name = "Heading_Corner";
            cornerFill.UpdateShadingPositionMatrix(pageSettings);
            pageLayout.AddChild(cornerFill);

            // ── Column headings (one per data column, spanning the top strip) ───────────────────
            double x = contentStartX;
            for (int col = page.FromColumn; col <= page.ToColumn; col++)
            {
                var mapCell = page.Map[page.FromRow, col];
                double colWidth = mapCell?.ColumnWidth ?? 0d;
                if (colWidth == 0d) { x += colWidth; continue; } // hidden column — skip

                string colLetter = ExcelCellBase.GetColumnLetter(col);

                var colFill = new PdfCellLayout(dictionaries, headingStyle,
                    x, pageSettings.ContentBounds.Top, colWidth, headingHeight);
                colFill.Name = "Heading_Col_" + colLetter;
                colFill.IsHeading = true;
                colFill.UpdateShadingPositionMatrix(pageSettings);
                pageLayout.AddChild(colFill);

                var colCell = CreateHeadingPdfCell(pageSettings, dictionaries,
                    colLetter, ExcelHorizontalAlignment.Center,
                    colWidth, headingHeight, fontName, fontSize);

                if (colCell.TextLines != null && colCell.TextLines.Count > 0)
                {
                    var info = new MergedCellDrawInfo { X = x, Y = pageSettings.ContentBounds.Top, Width = colWidth, Height = headingHeight };
                    var colText = new PdfCellContentLayout(pageSettings, dictionaries, colCell, info,
                        x, pageSettings.ContentBounds.Top, colWidth, headingHeight);
                    colText.Name = "Heading_Col_" + colLetter + "_Text";
                    colText.IsHeading = true;
                    colText.GidsAndCharMap(dictionaries);
                    pageLayout.AddChild(colText);
                }
                AddHeadingCellBorder(pageLayout, x, pageSettings.ContentBounds.Top, colWidth, headingHeight);
                x += colWidth;
            }
            // ── Row headings (one per data row, spanning the left strip) ─────────────────────────
            double y = contentStartY;
            for (int row = page.FromRow; row <= page.ToRow; row++)
            {
                double rowHeight = page.RowHeights[row - page.FromRow];
                if (rowHeight == 0d) { y -= rowHeight; continue; } // hidden row — skip

                string rowNum = row.ToString();

                var rowFill = new PdfCellLayout(dictionaries, headingStyle,
                    pageSettings.ContentBounds.Left, y, headingWidth, rowHeight);
                rowFill.Name = "Heading_Row_" + rowNum;
                rowFill.IsHeading = true;
                rowFill.UpdateShadingPositionMatrix(pageSettings);
                pageLayout.AddChild(rowFill);

                var rowCell = CreateHeadingPdfCell(pageSettings, dictionaries,
                    rowNum, ExcelHorizontalAlignment.Center,
                    headingWidth, rowHeight, fontName, fontSize);

                if (rowCell.TextLines != null && rowCell.TextLines.Count > 0)
                {
                    var info = new MergedCellDrawInfo { X = pageSettings.ContentBounds.Left, Y = y, Width = headingWidth, Height = rowHeight };
                    var rowText = new PdfCellContentLayout(pageSettings, dictionaries, rowCell, info,
                        pageSettings.ContentBounds.Left, y, headingWidth, rowHeight);
                    rowText.Name = "Heading_Row_" + rowNum + "_Text";
                    rowText.IsHeading = true;
                    rowText.GidsAndCharMap(dictionaries);
                    pageLayout.AddChild(rowText);
                }
                AddHeadingCellBorder(pageLayout, pageSettings.ContentBounds.Left, y, headingWidth, rowHeight);
                y -= rowHeight;
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
            cell.ContentAligmnet.HorizontalAlignment = hAlign;
            cell.ContentAligmnet.VerticalAlignment = ExcelVerticalAlignment.Bottom;
            cell.ContentAligmnet.WrapText = false;
            if (!string.IsNullOrEmpty(text))
            {
                var tf = new TextFragment();
                tf.Font = new OpenTypeFontInfoBase();
                tf.Font.Family = fontName;
                tf.Font.SubFamily = FontSubFamily.Regular;
                tf.Font.Size = fontSize;
                tf.Text = text;
                tf.RichTextOptions.Bold = false;
                tf.RichTextOptions.Italic = false;
                tf.RichTextOptions.UnderlineType = 12;  // none
                tf.RichTextOptions.StrikeType = 1;   // none
                cell.TextFragments = new List<TextFragment> { tf };
                PdfTextShaper.LayoutAndShapeText(pageSettings, dictionaries, cell);
            }
            return cell;
        }




        // TODO Count total pages when creating them isntead of looping them here.
        private static int GetTotalPages(List<Pages> pdfPages)
        {
            int totalPages = 0;
            for (int i = 0; i < pdfPages.Count; i++)
            {
                totalPages += pdfPages[i].Page.Length;
            }
            return totalPages;
        }

        internal static List<Pages> GetPages(PdfPageSettings pageSettings, PdfWorksheet[] pdfSheets)
        {
            List<Pages> PagesCollection = new List<Pages>();
            foreach (var pdfSheet in pdfSheets)
            {
                for (int ri = 0; ri < pdfSheet.Ranges.Count; ri++)
                {
                    var range = pdfSheet.Ranges[ri];
                    var pages = GetNumberOfPages(pageSettings, pdfSheet, ref range);
                    pages = AssignRangeToPages(pageSettings, range, pages);
                    pages = MapPage(range, pages);
                    pages = GetHeaderFooter(range, pages, pdfSheet);
                    pages = PrecomputeMergedCells(pageSettings, range, pages);
                    pages.HeadingFontName = pdfSheet.NormalStyle.Style.Font.Name;
                    pages.HeadingFontSize = pdfSheet.NormalStyle.Style.Font.Size;
                    pages.HeadingFill = pdfSheet.NormalStyle.Style.Fill;
                    PagesCollection.Add(pages);
                }
                if (pdfSheet.CommentsAndNotes.Range != null)
                {
                    var pages = GetNumberOfPages(pageSettings, pdfSheet, ref pdfSheet.CommentsAndNotes);
                    pages = AssignRangeToPages(pageSettings, pdfSheet.CommentsAndNotes, pages);
                    pages = MapPage(pdfSheet.CommentsAndNotes, pages);
                    pages.IsCommentsPage = true;
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
                    var mainCell = cell.Main ?? cell; // Main == null means this cell IS the top-left
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
                    double drawY = pageSettings.ContentBounds.Top - page.HeadingHeight;
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
                    page.MergedCells[key] = new MergedCellDrawInfo
                    {
                        X = drawX,
                        Y = drawY,
                        Width = mainCell.Width,
                        Height = mainCell.Height
                    };
                }
            }
            return page;
        }

        private static double[] BuildColumnXPositions(PdfPageSettings pageSettings, Page page)
        {
            int colCount = page.ToColumn - page.FromColumn + 1;
            var colX = new double[colCount];
            double x = pageSettings.ContentBounds.Left + page.HeadingWidth;
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

            var colSegments = GetColumnSegments(pageSettings, range, worksheet, addedWidth, range.PrintTitleWidth);
            var rowSegments = GetRowSegments(pageSettings, range, worksheet, addedHeight, range.PrintTitleHeight);

            pages.Page = new Page[colSegments.Count * rowSegments.Count];
            int i = 0;

            if (pageSettings.PageOrders == PageOrders.DownThenOver)
            {
                //foreach (var colSeg in colSegments)
                //    foreach (var rowSeg in rowSegments)
                //        pages.Page[i++] = new Page { FromColumn = colSeg.From, ToColumn = colSeg.To, FromRow = rowSeg.From, ToRow = rowSeg.To };
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
                            PrintTitleWidth = ci > 0 ? range.PrintTitleWidth : 0d,
                            PrintTitleHeight = ri > 0 ? range.PrintTitleHeight : 0d,
                        };
            }
            else //if (pageSettings.PageOrders == PageOrders.OverThenDown)
            {
                //foreach (var rowSeg in rowSegments)
                //    foreach (var colSeg in colSegments)
                //        pages.Page[i++] = new Page { FromColumn = colSeg.From, ToColumn = colSeg.To, FromRow = rowSeg.From, ToRow = rowSeg.To };
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
                            PrintTitleWidth = ci > 0 ? range.PrintTitleWidth : 0d,
                            PrintTitleHeight = ri > 0 ? range.PrintTitleHeight : 0d,
                        };
            }

            //for (int k = 0; k < pages.Page.Length; k++)
            //{
            //    var p = pages.Page[k];
            //    p.HeadingWidth = addedWidth;
            //    p.HeadingHeight = addedHeight;
            //    pages.Page[k] = p;
            //}
            pdfPages = pages;
            return pdfPages;
        }

        private struct PageSegment
        {
            public int From;
            public int To;
            public PageSegment(int from, int to) { From = from; To = to; }
        }

        private static List<PageSegment> GetColumnSegments(PdfPageSettings pageSettings, PdfRange range, ExcelWorksheet worksheet, double addedWidth, double titleWidth)
        {
            var segments = new List<PageSegment>();
            int segStartIdx = 0;
            double width = 0d;

            for (int col = 0; col < range.ColWidths.Count; col++)
            {
                int actualCol = range.Range._fromCol + col;

                double effectiveAdded = addedWidth + (segments.Count == 0 ? 0d : titleWidth);


                // Content-bounds overflow: col doesn't fit, end segment before it and reprocess.
                if (width + range.ColWidths[col] + effectiveAdded >= pageSettings.ContentBounds.Width)
                {
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

        private static List<PageSegment> GetRowSegments(PdfPageSettings pageSettings, PdfRange range, ExcelWorksheet worksheet, double addedHeight, double titleHeight)
        {
            var segments = new List<PageSegment>();
            int segStartIdx = 0;
            double height = 0d;

            for (int row = 0; row < range.RowHeights.Count; row++)
            {
                int actualRow = range.Range._fromRow + row;

                double effectiveAdded = addedHeight + (segments.Count == 0 ? 0d : titleHeight);

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
            PdfTextShaper.LayoutAndShapeText(pageSettings, dictionaries, hf.Content);
        }

    }
}

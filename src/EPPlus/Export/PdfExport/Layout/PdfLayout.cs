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
using EPPlus.Export.Pdf.DocumentObjects;
using EPPlus.Export.Pdf.Helpers;
using EPPlus.Export.Pdf.Layout;
using EPPlus.Export.Pdf.Resources;
using EPPlus.Export.Pdf.Settings;
using EPPlus.Export.Pdf.Settings.PdfPageSizes;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.Integration.DataHolders;
using EPPlus.Graphics;
using EPPlus.Graphics.Geometry;
using EPPlus.Graphics.Units;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Export.PdfExport.Data;
using OfficeOpenXml.Export.PdfExport.TextMapping;
using OfficeOpenXml.Export.PdfExport.TextShaping;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
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
        public double ClipX;
        public double ClipY;
        public double ClipWidth;
        public double ClipHeight;
    }

    internal struct PrintTitleHeadingDraw
    {
        public bool IsRow;
        public int Index;
        public double X;
        public double Y;
        public double Width;
        public double Height;
    }

    internal struct SpillCellDraw
    {
        public PdfCell Cell;
        public double X, Y, Width, Height;
        public double ClipX, ClipY, ClipWidth, ClipHeight;
        public bool IsPrintTitle;
    }

    internal class PdfLayout
    {
        private const double rowHeadingWith1CharWidth = 18d;

        public static Transform GetLayout(PdfPageSettings[] sheetSettings, PdfDictionaries dictionaries, PdfWorksheet[] pdfSheets)
        {
            var PagesCollection = GetPages(sheetSettings, pdfSheets);
            var Catalog = GetCatalog(sheetSettings[0].FirstPageNumber, dictionaries, PagesCollection);
            if (Catalog is PdfCatalogLayout cat)
                PdfBookmarkCollector.CollectNamedDestinations(pdfSheets, cat.NamedDestinations);
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
                    double contentStartX = GetOriginX(pageSettings, page) + page.HeadingWidth + page.PrintTitleWidth;
                    double contentStartY = GetOriginY(pageSettings, page) - page.HeadingHeight - page.PrintTitleHeight;
                    if (pageSettings.ShowHeadings && !pdfPages[i].IsCommentsPage)
                    {
                        PdfHeadingLayout.AddHeadingCells(pageSettings, dictionaries, page, pageLayout, contentStartX, contentStartY, page.HeadingWidth, page.HeadingHeight, pdfPages[i].HeadingFontName, pdfPages[i].HeadingFontSize, pdfPages[i].HeadingFill);
                        PdfHeadingLayout.AddPrintTitleHeadings(pageSettings, dictionaries, page, pageLayout, pdfPages[i].HeadingFontName, pdfPages[i].HeadingFontSize, pdfPages[i].HeadingFill);
                    }
                    PdfSpillLayout.AddSpillCells(pageSettings, dictionaries, page, pageLayout);
                    PdfHeadingLayout.AddPrintTitleCells(pageSettings, dictionaries, page, pageLayout);
                    if (!pdfPages[i].IsCommentsPage)
                    {
                        pageLayout.SheetName = pdfPages[i].SheetName;
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
                        pageLayout.CommentsSourceSheet = pdfPages[i].SheetName;
                    }
                    double y = contentStartY;
                    double x = contentStartX;
                    for (int row = pages[j].FromRow; row <= pages[j].ToRow; row++)
                    {
                        double rowHeight = pages[j].RowHeights[row - pages[j].FromRow];
                        for (int col = pages[j].FromColumn; col <= pages[j].ToColumn; col++)
                        {
                            var map = pages[j].Map[row, col];
                            MergedCellDrawInfo info = new MergedCellDrawInfo();
                            if (map.Hidden && !map.Merged)
                            {
                                x += map.ColumnWidth;
                                continue;
                            }
                            //Merged Cell
                            if (map.Merged)
                            {
                                string key = map.MergedAddress.Address;
                                if (!drawnMergedCells.Contains(key) &&
                                    pages[j].MergedCells.TryGetValue(key, out info))
                                {
                                    if (!isDraft)
                                    {
                                        //Fill
                                        if (!isBlackWhite)
                                        {
                                            var cellStyle = map.Main?.CellStyle ?? map.CellStyle;
                                            var fill = new PdfCellLayout(info.X, info.Y, info.Width, info.Height);
                                            PdfFillResolver.SetFill(dictionaries, cellStyle, map.Text, fill);
                                            fill.Name = map.Name;
                                            fill.UpdateShadingPositionMatrix(pageSettings);
                                            pageLayout.AddChild(fill);
                                        }
                                        //In-cell picture
                                        var picMap = (map.CellPictureBytes != null) ? map : (map.Main != null && map.Main.CellPictureBytes != null) ? map.Main : null;
                                        if (picMap != null)
                                        {
                                            PdfImagePlacement.AddCellPicture(pageLayout, picMap, info.X, info.Y, info.Width, info.Height);
                                        }
                                    }
                                    //Text
                                    var sourceMap = (map.TextLines != null && map.TextLines.Count > 0) ? map : (map.Main != null && map.Main.TextLines != null && map.Main.TextLines.Count > 0) ? map.Main : null;
                                    if (sourceMap != null)
                                    {
                                        var text = new PdfCellContentLayout(pageSettings, dictionaries, sourceMap, info.X, info.Y, info.Width, info.Height);
                                        text.Name = map.Name;
                                        text.GidsAndCharMap(dictionaries);
                                        text.SetupClipping(info.X, info.Y, info.Width, info.Height);
                                        pageLayout.AddChild(text);

                                        if (map.HyperlinkUri != null || map.HyperlinkTargetSheet != null)
                                        {
                                            if (PdfLinkLayout.CanTightenLink(text))
                                                PdfLinkLayout.AddTextLinks(pageLayout, text, info.X, info.X + info.Width, map.HyperlinkUri, map.HyperlinkTargetSheet, map.HyperlinkTargetRow, map.HyperlinkTargetCol, map.HyperlinkToolTip);
                                            else
                                                
                                                pageLayout.Links.Add(PdfLinkLayout.CreateLinkDrawInfo(info.X, info.Y, info.Width, info.Height, map));
                                        }
                                    }
                                    else if (map.HyperlinkUri != null || map.HyperlinkTargetSheet != null)
                                    {
                                        pageLayout.Links.Add(PdfLinkLayout.CreateLinkDrawInfo(info.X, info.Y, info.Width, info.Height, map));
                                    }
                                    if (!isDraft)
                                    {
                                        if (map.Main != null)
                                        {
                                            var mergeMainStyle = map.Main.CellStyle;
                                            if (PdfBorderResolver.HasDiagonalBorder(mergeMainStyle))
                                            {
                                                var diagBorder = new PdfCellBorderLayout(
                                                    isMerged: false,
                                                    corners: MergedCellCorners.All,
                                                    info: info,
                                                    x: info.X,
                                                    y: info.Y,
                                                    width: info.Width,
                                                    height: info.Height);
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
                                    drawnMergedCells.Add(key);
                                }
                            }
                            else
                            {
                                var contentRight = pageSettings.ContentBounds.Left + pageSettings.ContentBounds.Width;
                                var effectiveWidth = GetClampedCellWidth(pageSettings, x, map.ColumnWidth);
                                if (!isDraft)
                                {
                                    if (!isBlackWhite)
                                    {
                                        //Fill
                                        var fill = new PdfCellLayout(x, y, effectiveWidth, rowHeight);
                                        PdfFillResolver.SetFill(dictionaries, map.CellStyle, map.Text, fill);
                                        fill.UpdateShadingPositionMatrix(pageSettings);
                                        fill.Name = map.Name;
                                        pageLayout.AddChild(fill);
                                    }
                                    //In-cell picture
                                    PdfImagePlacement.AddCellPicture(pageLayout, map, x, y, effectiveWidth, rowHeight);
                                }
                                //Text
                                if (map.TextLines != null && map.TextLines.Count > 0)
                                {
                                    var text = new PdfCellContentLayout(pageSettings, dictionaries, map, x, y, effectiveWidth, rowHeight);
                                    text.Name = map.Name;
                                    text.GidsAndCharMap(dictionaries);
                                    var (clipX, clipW, spillLeft, spillRight) = PdfSpillLayout.GetOverflowClip(map, pages[j], row, col, x, effectiveWidth);
                                    text.SetupClipping(clipX, y, clipW, rowHeight, spillLeft, spillRight);
                                    pageLayout.AddChild(text);

                                    if (map.HyperlinkUri != null || map.HyperlinkTargetSheet != null)
                                    {
                                        if (PdfLinkLayout.CanTightenLink(text))
                                            PdfLinkLayout.AddTextLinks(pageLayout, text, clipX, clipX + clipW,
                                                         map.HyperlinkUri, map.HyperlinkTargetSheet, map.HyperlinkTargetRow, map.HyperlinkTargetCol, map.HyperlinkToolTip);
                                        else
                                            pageLayout.Links.Add(PdfLinkLayout.CreateLinkDrawInfo(x, y, effectiveWidth, rowHeight, map));
                                    }
                                }
                                // Empty cell in a range hyperlink
                                else if (map.HyperlinkUri != null || map.HyperlinkTargetSheet != null)
                                {
                                    pageLayout.Links.Add(PdfLinkLayout.CreateLinkDrawInfo(x, y, effectiveWidth, rowHeight, map));
                                }
                            }
                            if (!isDraft)
                            {
                                //Border
                                var borderStyle = (map.Merged && map.Main != null) ? map.Main.CellStyle : map.CellStyle;
                                if (PdfBorderResolver.HasBorder(map.CellStyle))
                                {
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
                                    PdfBorderResolver.SetDoubleBorderMiterFlags(pages[j], row, col, border);
                                    pageLayout.AddChild(border);
                                }
                            }
                            x += map.ColumnWidth;
                        }
                        y -= rowHeight;
                        x = contentStartX;
                    }
                    if (!isDraft)
                    {
                        if (page.Images != null)
                        {
                            foreach (var img in page.Images)
                            {
                                pageLayout.AddChild(new PdfImageLayout(img.X, img.Y, img.Width, img.Height)
                                {
                                    ImageBytes = img.ImageBytes,
                                    Name = img.Name,
                                });
                                if (img.HyperlinkUri != null || img.HyperlinkTargetSheet != null)
                                {
                                    pageLayout.Links.Add(new LinkDrawInfo
                                    {
                                        X = img.X,
                                        Y = img.Y - img.Height,
                                        Width = img.Width,
                                        Height = img.Height,
                                        Uri = img.HyperlinkUri,
                                        TargetSheet = img.HyperlinkTargetSheet,
                                        TargetRow = img.HyperlinkTargetRow,
                                        TargetColumn = img.HyperlinkTargetColumn,
                                        ToolTip = img.HyperlinkToolTip
                                    });
                                }
                            }
                        }
                    }
                    if (page.HeaderFooters != null)
                    {
                        pageLayout.ScaleHeaderFooterWithDocument = page.HeaderFooters.ScaleWithDocument;
                        var hfType = page.HeaderFooters.GetPageType(physicalPageIndex);
                        var leftH = page.HeaderFooters.Get(hfType, HeaderFooterSection.Header, HeaderFooterAlignment.Left);
                        bool alignHf = page.HeaderFooters.AlignWithMargins;
                        var (hfBoxLeft, hfBoxWidth) = PdfHeaderFooterLayout.HeaderFooterBox(pageSettings, alignHf);
                        if (leftH != null && !leftH.HasImage && PdfHeaderFooterLayout.ShapeHeaderFooterText(pageSettings, dictionaries, leftH, displayedPageNumber, totalPages))
                        {
                            var ascent = leftH.Content.TextLines[0].LargestAscent;
                            var hfx = hfBoxLeft;
                            var hfy = pageSettings.PageSize.HeightPu - pageSettings.Margins.HeaderPu - ascent;
                            var text = new PdfCellContentLayout(pageSettings, dictionaries, leftH, hfx, hfy, 0, 0);
                            text.ScaleAnchorX = hfx;                 // left edge
                            text.ScaleAnchorY = hfy + ascent;
                            text.Name = "LeftHeader";
                            PdfHeaderFooterLayout.AddHeaderFooterTextPart(pageLayout, dictionaries, text, HeaderFooterSection.Header, HeaderFooterAlignment.Left, pageSettings, alignHf);
                        }
                        var centerH = page.HeaderFooters.Get(hfType, HeaderFooterSection.Header, HeaderFooterAlignment.Center);
                        if (centerH != null && !centerH.HasImage && PdfHeaderFooterLayout.ShapeHeaderFooterText(pageSettings, dictionaries, centerH, displayedPageNumber, totalPages))
                        {
                            var ascent = centerH.Content.TextLines[0].LargestAscent;
                            var hfx = hfBoxLeft;
                            var hfy = pageSettings.PageSize.HeightPu - pageSettings.Margins.HeaderPu - ascent;
                            var hfWidth = hfBoxWidth;
                            var text = new PdfCellContentLayout(pageSettings, dictionaries, centerH, hfx, hfy, hfWidth, 0);
                            text.ScaleAnchorX = hfx + hfWidth / 2d;
                            text.ScaleAnchorY = hfy + ascent;
                            text.Name = "CenterHeader";
                            PdfHeaderFooterLayout.AddHeaderFooterTextPart(pageLayout, dictionaries, text, HeaderFooterSection.Header, HeaderFooterAlignment.Center, pageSettings, alignHf);
                        }
                        var rightH = page.HeaderFooters.Get(hfType, HeaderFooterSection.Header, HeaderFooterAlignment.Right);
                        if (rightH != null && !rightH.HasImage && PdfHeaderFooterLayout.ShapeHeaderFooterText(pageSettings, dictionaries, rightH, displayedPageNumber, totalPages))
                        {
                            var ascent = rightH.Content.TextLines[0].LargestAscent;
                            var hfx = hfBoxLeft + hfBoxWidth;
                            var hfy = pageSettings.PageSize.HeightPu - pageSettings.Margins.HeaderPu - ascent;
                            var text = new PdfCellContentLayout(pageSettings, dictionaries, rightH, hfx, hfy, 0, 0);
                            text.ScaleAnchorX = hfx;
                            text.ScaleAnchorY = hfy + ascent;
                            text.Name = "RightHeader";
                            PdfHeaderFooterLayout.AddHeaderFooterTextPart(pageLayout, dictionaries, text, HeaderFooterSection.Header, HeaderFooterAlignment.Right, pageSettings, alignHf);
                        }
                        var leftF = page.HeaderFooters.Get(hfType, HeaderFooterSection.Footer, HeaderFooterAlignment.Left);
                        if (leftF != null && !leftF.HasImage && PdfHeaderFooterLayout.ShapeHeaderFooterText(pageSettings, dictionaries, leftF, displayedPageNumber, totalPages))
                        {
                            int last = leftF.Content.TextLines.Count - 1;
                            var descent = leftF.Content.TextLines[last].LargestDescent;
                            var hfx = hfBoxLeft;
                            var hfy = pageSettings.Margins.FooterPu + descent;
                            var text = new PdfCellContentLayout(pageSettings, dictionaries, leftF, hfx, hfy, 0, 0);
                            text.ScaleAnchorX = hfx;
                            text.ScaleAnchorY = hfy - descent;
                            text.Name = "LeftFooter";
                            PdfHeaderFooterLayout.AddHeaderFooterTextPart(pageLayout, dictionaries, text, HeaderFooterSection.Footer, HeaderFooterAlignment.Left, pageSettings, alignHf);
                        }
                        var centerF = page.HeaderFooters.Get(hfType, HeaderFooterSection.Footer, HeaderFooterAlignment.Center);
                        if (centerF != null && !centerF.HasImage && PdfHeaderFooterLayout.ShapeHeaderFooterText(pageSettings, dictionaries, centerF, displayedPageNumber, totalPages))
                        {
                            int last = centerF.Content.TextLines.Count - 1;
                            var descent = centerF.Content.TextLines[last].LargestDescent;
                            var hfx = hfBoxLeft;
                            var hfy = pageSettings.Margins.FooterPu + descent;
                            var hfWidth = hfBoxWidth;
                            var text = new PdfCellContentLayout(pageSettings, dictionaries, centerF, hfx, hfy, hfWidth, 0);
                            text.ScaleAnchorX = hfx + hfWidth / 2d;
                            text.ScaleAnchorY = hfy - descent;
                            text.Name = "CenterFooter";
                            PdfHeaderFooterLayout.AddHeaderFooterTextPart(pageLayout, dictionaries, text, HeaderFooterSection.Footer, HeaderFooterAlignment.Center, pageSettings, alignHf);
                        }
                        var rightF = page.HeaderFooters.Get(hfType, HeaderFooterSection.Footer, HeaderFooterAlignment.Right);
                        if (rightF != null && !rightF.HasImage && PdfHeaderFooterLayout.ShapeHeaderFooterText(pageSettings, dictionaries, rightF, displayedPageNumber, totalPages))
                        {
                            int last = rightF.Content.TextLines.Count - 1;
                            var descent = rightF.Content.TextLines[last].LargestDescent;
                            var hfx = hfBoxLeft + hfBoxWidth;
                            var hfy = pageSettings.Margins.FooterPu + descent;
                            var text = new PdfCellContentLayout(pageSettings, dictionaries, rightF, hfx, hfy, 0, 0);
                            text.ScaleAnchorX = hfx;
                            text.ScaleAnchorY = hfy - descent;
                            text.Name = "RightFooter";
                            PdfHeaderFooterLayout.AddHeaderFooterTextPart(pageLayout, dictionaries, text, HeaderFooterSection.Footer, HeaderFooterAlignment.Right, pageSettings, alignHf);
                        }
                        double hfContentLeft = pageSettings.Margins.LeftPu;
                        double hfContentWidth = pageSettings.PageSize.WidthPu - pageSettings.Margins.LeftPu - pageSettings.Margins.RightPu;
                        foreach (var hfSection in new[] { HeaderFooterSection.Header, HeaderFooterSection.Footer })
                        {
                            foreach (var hfAlign in new[] { HeaderFooterAlignment.Left, HeaderFooterAlignment.Center, HeaderFooterAlignment.Right })
                            {
                                var hf = page.HeaderFooters.Get(hfType, hfSection, hfAlign);
                                if (hf == null || !hf.HasImage) continue;
                                PdfHeaderFooterLayout.ApplyPageNumbers(hf, displayedPageNumber, totalPages);

                                var frags = hf.Content.TextFragments;
                                int splitAt = hf.ImageFragmentIndex;
                                if (splitAt < 0 || splitAt > frags.Count) splitAt = frags.Count;
                                var beforeFrags = frags.GetRange(0, splitAt);
                                var afterFrags = frags.GetRange(splitAt, frags.Count - splitAt);

                                double beforeWidth, beforeAscent, beforeDescent;
                                double afterWidth, afterAscent, afterDescent;
                                var beforeRun = PdfHeaderFooterLayout.ShapeHeaderFooterRun(pageSettings, dictionaries, beforeFrags, hfSection, out beforeWidth, out beforeAscent, out beforeDescent);
                                var afterRun = PdfHeaderFooterLayout.ShapeHeaderFooterRun(pageSettings, dictionaries, afterFrags, hfSection, out afterWidth, out afterAscent, out afterDescent);

                                bool canEmbed = PdfImageXObject.CanEmbed(hf.ImageBytes);
                                double imgW = canEmbed ? hf.ImageWidth : 0d;
                                double imgH = canEmbed ? hf.ImageHeight : 0d;

                                double runWidth = beforeWidth + imgW + afterWidth;
                                double runStart;
                                switch (hfAlign)
                                {
                                    case HeaderFooterAlignment.Left:
                                        runStart = hfContentLeft;
                                        break;
                                    case HeaderFooterAlignment.Right:
                                        runStart = hfContentLeft + hfContentWidth - runWidth;
                                        break;
                                    default: // Center
                                        runStart = hfContentLeft + (hfContentWidth - runWidth) / 2d;
                                        break;
                                }

                                double ascent = System.Math.Max(beforeAscent, afterAscent);
                                double descent = System.Math.Max(beforeDescent, afterDescent);
                                double imgTop = hfSection == HeaderFooterSection.Header
                                                ? pageSettings.PageSize.HeightPu - pageSettings.Margins.HeaderPu
                                                : pageSettings.Margins.FooterPu + imgH;
                                double imgBottom = imgTop - imgH;

                                double naturalBaseline = hfSection == HeaderFooterSection.Header
                                                ? pageSettings.PageSize.HeightPu - pageSettings.Margins.HeaderPu - ascent
                                                : pageSettings.Margins.FooterPu + descent;
                                double textY = System.Math.Min(naturalBaseline, imgBottom + descent);
                                var part = new PdfHeaderFooterPart { ScalePivot = PdfHeaderFooterLayout.HeaderFooterPivot(hfSection, hfAlign, pageSettings, alignHf) };
                                double cursor = runStart;
                                if (beforeRun != null)
                                {
                                    part.Items.Add(PdfHeaderFooterLayout.PlaceHeaderFooterRun(pageSettings, dictionaries, beforeRun, cursor, textY));
                                }
                                cursor += beforeWidth;

                                if (canEmbed)
                                {
                                    part.Items.Add(new PdfImageLayout(cursor, imgTop, imgW, imgH)
                                    {
                                        ImageBytes = hf.ImageBytes,
                                        IsHeaderFooter = true,
                                        Name = $"{hfSection}{hfAlign}Image",
                                    });
                                }
                                cursor += imgW;
                                if (afterRun != null)
                                {
                                    part.Items.Add(PdfHeaderFooterLayout.PlaceHeaderFooterRun(pageSettings, dictionaries, afterRun, cursor, textY));
                                }
                                if (part.Items.Count > 0)
                                    pageLayout.HeaderFooterParts.Add(part);
                            }
                        }
                    }
                    if (page.AddedContent != null && page.AddedContent.content?.TextLines != null && page.AddedContent.content.TextLines.Count > 0)
                    {
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
                    if (!isDraft)
                        PdfGridlinesLayout.AddGridLines(pageSettings, pages[j], pageLayout, borderOnly: !pageSettings.ShowGridLines || pdfPages[i].IsCommentsPage);
                    pageLayout.ChildObjects.Sort((a, b) =>
                    {
                        int cmp = a.Z.CompareTo(b.Z);
                        if (cmp == 0)
                            return string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase);
                        return cmp;
                    });
                    displayedPageNumber++;
                    physicalPageIndex++;
                    Catalog.AddChild(pageLayout);
                }
            }
            return Catalog;
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
                        pages.Page[k].IsCommentsPage = true;
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
            double pageRightNatural = pageSettings.ContentBounds.Left
                                    + (pageSettings.PageSize.WidthPu - pageSettings.ContentBounds.Left) / s;
            return System.Math.Min(cellWidth, pageRightNatural - cellX);
        }


    }
}

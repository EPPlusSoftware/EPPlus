using EPPlus.Export.Pdf.DocumentObjects;
using EPPlus.Export.Pdf.Layout;
using EPPlus.Export.Pdf.Resources;
using EPPlus.Export.Pdf.Settings;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Graphics.Geometry;
using EPPlus.Graphics.Units;
using OfficeOpenXml.Export.PdfExport.Data;
using OfficeOpenXml.Export.PdfExport.TextShaping;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.PdfExport.Layout
{
    internal static class PdfHeaderFooterLayout
    {
        internal static Pages GetHeaderFooter(PdfRange range, Pages pdfPages, PdfWorksheet pdfSheet)
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

        internal static void AddHeaderFooter(PdfPageLayout pageLayout, Page page, PdfPageSettings pageSettings, PdfDictionaries dictionaries, int displayedPageNumber, int physicalPageIndex, int totalPages)
        {
            if (page.HeaderFooters == null) return;
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

        private static void SubstitutePageNumbers(PdfPageSettings pageSettings, PdfDictionaries dictionaries, PdfHeaderFooter hf, int pageNumber, int totalPages)
        {
            if (hf == null) return;
            ApplyPageNumbers(hf, pageNumber, totalPages);
            PdfTextShaper.ShapeText(pageSettings, dictionaries, hf.Content);
        }

        internal static bool ShapeHeaderFooterText(PdfPageSettings pageSettings, PdfDictionaries dictionaries, PdfHeaderFooter hf, int pageNumber, int totalPages)
        {
            SubstitutePageNumbers(pageSettings, dictionaries, hf, pageNumber, totalPages);
            return hf.Content != null && hf.Content.TextLines != null && hf.Content.TextLines.Count > 0;
        }

        internal static Vector2 HeaderFooterPivot(HeaderFooterSection section, HeaderFooterAlignment align, PdfPageSettings ps, bool alignWithMargins)
        {
            var (left, width) = HeaderFooterBox(ps, alignWithMargins);
            double x = align == HeaderFooterAlignment.Left ? left
                     : align == HeaderFooterAlignment.Right ? left + width
                     : left + width / 2d;
            double y = section == HeaderFooterSection.Header
                     ? ps.PageSize.HeightPu - ps.Margins.HeaderPu
                     : ps.Margins.FooterPu;
            return new Vector2(x, y);
        }
        internal static (double Left, double Width) HeaderFooterBox(PdfPageSettings ps, bool alignWithMargins)
        {
            if (alignWithMargins)
                return (ps.Margins.LeftPu, ps.PageSize.WidthPu - ps.Margins.LeftPu - ps.Margins.RightPu);

            double d = UnitConversion.MmToPoints(19.05);
            return (d, ps.PageSize.WidthPu - 2d * d);
        }
        internal static void AddHeaderFooterTextPart(PdfPageLayout pageLayout, PdfDictionaries dictionaries, PdfCellContentLayout text, HeaderFooterSection section, HeaderFooterAlignment align, PdfPageSettings ps, bool alignWithMargins)
        {
            text.IsHeaderFooter = true;
            text.GidsAndCharMap(dictionaries);
            var part = new PdfHeaderFooterPart { ScalePivot = HeaderFooterPivot(section, align, ps, alignWithMargins) };
            part.Items.Add(text);
            pageLayout.HeaderFooterParts.Add(part);
        }

        internal static void ApplyPageNumbers(PdfHeaderFooter hf, int pageNumber, int totalPages)
        {
            if (hf == null) return;
            foreach (var idx in hf.PageNumberIndexes)
                hf.Content.TextFragments[idx].Text = pageNumber.ToString();
            foreach (var idx in hf.NumberOfPagesIndexes)
                hf.Content.TextFragments[idx].Text = totalPages.ToString();
        }

        internal static PdfHeaderFooter ShapeHeaderFooterRun(PdfPageSettings pageSettings, PdfDictionaries dictionaries, List<TextFragment> frags, HeaderFooterSection section,
    out double width, out double ascent, out double descent)
        {
            width = 0d; ascent = 0d; descent = 0d;
            if (frags == null || frags.Count == 0) return null;
            var sub = new PdfHeaderFooter(frags, new List<int>(), new List<int>(), HeaderFooterType.Odd, HeaderFooterAlignment.Left, section);
            sub.Content.ContentAligmnet = new PdfCellAlignmentData
            {
                HorizontalAlignment = EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment.Left
            };
            PdfTextShaper.ShapeText(pageSettings, dictionaries, sub.Content);
            if (sub.Content.TextLines == null || sub.Content.TextLines.Count == 0 || sub.Content.TextLines.LineFragments == null || sub.Content.TextLines.LineFragments.Count == 0 || sub.Content.TotalTextLength <= 0d)
                return null;
            width = sub.Content.TextLines[0].Width;
            ascent = sub.Content.TextLines[0].LargestAscent;
            descent = sub.Content.TextLines[0].LargestDescent;
            return sub;
        }

        internal static PdfCellContentLayout PlaceHeaderFooterRun(PdfPageSettings pageSettings, PdfDictionaries dictionaries, PdfHeaderFooter run, double x, double baselineY)
        {
            var layout = new PdfCellContentLayout(pageSettings, dictionaries, run, x, baselineY, 0, 0);
            layout.IsHeaderFooter = true;
            layout.GidsAndCharMap(dictionaries);
            return layout;
        }
    }
}

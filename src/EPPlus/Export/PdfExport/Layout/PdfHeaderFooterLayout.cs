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
    internal class PdfHeaderFooterLayout
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

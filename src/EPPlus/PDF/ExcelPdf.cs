using OfficeOpenXml.PDF.PdfFontData;
using OfficeOpenXml.PDF.PdfGraphics;
using OfficeOpenXml.PDF.Pdfhelpers;
using OfficeOpenXml.PDF.PdfObjects;
using OfficeOpenXml.PDF.PdfSettings;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml.PDF.PdfSettings.PdfPageData;
using OfficeOpenXml.PDF.PdfLayout;

namespace OfficeOpenXml.PDF
{
    public class ExcelPdf
    {
        internal ExcelWorksheet _ws;
        internal string header = "%PDF-1.7\n";
        internal List<PdfObject> body = new List<PdfObject>();
        internal PdfCrossRefTable crossRefTable;
        internal readonly string defaultFontName;
        internal List<PdfExcelPageData> pagesData = new List<PdfExcelPageData>();
        internal readonly Dictionary<string, PdfFontResource> fontResources = new Dictionary<string, PdfFontResource>();

        private static Dictionary<uint, PdfFontProperties> _fonts;
        private PdfPageSettings PageSettings;
        private double cellMargin = 0.2d;
        private PdfContentBounds bounds;

        public ExcelPdf(ExcelWorksheet worksheet)
        {
            _ws = worksheet;
            defaultFontName = worksheet.Workbook.ThemeManager.CurrentTheme.FontScheme.MinorFont[0].Typeface;
            if (!PdfExcelPageDataLookup.PdfExcelA4PageData.ContainsKey(defaultFontName))
            {
                pagesData.Add(new PdfExcelPageData(-1, -1));
            }
            else
            {
                pagesData.Add(new PdfExcelPageData(PdfExcelPageDataLookup.PdfExcelA4PageData[defaultFontName][0], PdfExcelPageDataLookup.PdfExcelA4PageData[defaultFontName][1]));
            }
            PageSettings = new PdfPageSettings();
            bounds = new PdfContentBounds(PageSettings.Margins, PageSettings.PageSize);
        }

        public ExcelPdf(ExcelWorksheet worksheet, PdfPageSettings pageSettings)
        {
            _ws = worksheet;
            defaultFontName = worksheet.Workbook.ThemeManager.CurrentTheme.FontScheme.MinorFont[0].Typeface;
            if (!PdfExcelPageDataLookup.PdfExcelA4PageData.ContainsKey(defaultFontName))
            {
                PdfExcelPageDataLookup.PdfExcelA4PageData.Add(defaultFontName, [-1, -1]);

            }
            PageSettings = pageSettings;
            bounds = new PdfContentBounds(PageSettings.Margins, PageSettings.PageSize);
        }

        internal string GetFontLabel(string fontName, string subFamily, double fontSize)
        {
            if (!fontResources.ContainsKey(fontName))
            {
                int label = 1;
                if (fontResources.Count > 0)
                {
                    label = fontResources.Last().Value.labelNumber + 1;
                }
                PdfFontResource fr = new PdfFontResource(fontName, subFamily, label, PageSettings);
                if (fontName != "Courier New")
                {
                    body.Add(fr.GetFontDescriptorObject(body.Count + 1));
                    body.Add(fr.GetWidthsObject(body.Count + 1));
                }
                body.Add(fr.GetFontObject(body.Count + 1));
                fontResources.Add(fontName, fr);
            }
            return fontResources[fontName].Label;
        }

        internal void AddFontData()
        {
            foreach (var font in fontResources)
            {
                if (font.Key != "Courier New")
                {
                    body.Add(font.Value.GetFontDescriptorObject(body.Count + 1));
                    body.Add(font.Value.GetWidthsObject(body.Count + 1));
                }
                body.Add(font.Value.GetFontObject(body.Count + 1));
            }
        }

        //create page
        private PdfPage AddPage(int pagesObjectNumber, List<int> contentObjectNumbers, PdfPageSettings settings)
        {
            var page = new PdfPage(body.Count + 1, pagesObjectNumber, contentObjectNumbers, settings.PageSize, fontResources);
            body.Add(page);
            return page;
        }
        //create pages
        private PdfPages AddPages()
        {
            var pages = new PdfPages(body.Count + 1, new List<int>{});
            body.Add(pages);
            return pages;
        }
        //create Catalog
        private PdfCatalog AddCatalog(int pagesObjectNumber)
        {
            var catalog = new PdfCatalog(body.Count + 1, pagesObjectNumber);
            body.Add(catalog);
            return catalog;
        }

        private void CreateStreamContentFromCell(PdfTransform pageLayout, PdfPage page)
        {
            var cells = pageLayout.ChildObjects.Where(t => t is PdfCellLayout || t is PdfCellContentLayout || t is PdfCellBorderLayout).GroupBy(t => t.Name);
            var contentStream = new PdfContentStream(body.Count + 1);
            if (PageSettings.ShowGridLines)
            {
                DrawGridLines(contentStream, pageLayout, page);
            }
            foreach (var cell in cells)
            {
                foreach (var cellPart in cell)
                {
                    contentStream.AddCommand("q");
                    switch (cellPart)
                    {
                        case PdfCellLayout layout:
                            contentStream.AddCellLayout(layout);
                            break;
                        case PdfCellContentLayout contentLayout:
                            contentStream.AddCellContentLayout(contentLayout, GetFontLabel(contentLayout.FontData.FontName, contentLayout.FontData.SubFamily, contentLayout.FontData.FontSize));
                            break;
                        case PdfCellBorderLayout borderLayout:
                            contentStream.AddBorderLayout(borderLayout);
                            break;
                    }
                    contentStream.AddCommand("Q");
                }
            }
            if (PageSettings.ShowGridLines)
            {
                DrawBorderLines(contentStream, pageLayout, page);
            }
            body.Add(contentStream);
            page.contentObjectNumbers.Add(contentStream.objectNumber);
        }
        private void DrawGridLines(PdfContentStream contentStream, PdfTransform pageLayout, PdfPage page)
        {
            if (pageLayout is not PdfPageLayout pl)
                return;

            contentStream.AddCommand("q");
            contentStream.AddCommand($"{GridLine.Width.ToPdfString()} w");
            contentStream.AddCommand(PdfColor.Black.ToFillCommand());
            foreach (var line in pl.GridLines)
            {
                string w, h;
                if (line.X1 == line.X2)
                {
                    w = GridLine.Width.ToPdfString();
                    h = System.Math.Abs(line.Y2 - line.Y1).ToPdfString();
                }
                else
                {
                    w = System.Math.Abs(line.X2 - line.X1).ToPdfString();
                    h = GridLine.Width.ToPdfString();
                }
                contentStream.AddCommand($"{(line.X1).ToPdfString()} {(line.Y1).ToPdfString()} {w} {h} re");
            }
            contentStream.AddCommand("f");
            contentStream.AddCommand("Q");
        }
        private void DrawBorderLines(PdfContentStream contentStream, PdfTransform pageLayout, PdfPage page)
        {
            if (pageLayout is not PdfPageLayout pl)
                return;

            contentStream.AddCommand("q");
            contentStream.AddCommand("1.0 w");
            contentStream.AddCommand("2 J");
            contentStream.AddCommand("[] 0 d");
            contentStream.AddCommand(PdfColor.Black.ToStrokeCommand());
            foreach (var line in pl.BorderLines)
            {
                contentStream.AddCommand($"{line.X1.ToPdfString()} {line.Y1.ToPdfString()} m");
                contentStream.AddCommand($"{line.X2.ToPdfString()} {line.Y2.ToPdfString()} l");
            }
            contentStream.AddCommand("S");
            contentStream.AddCommand("Q");
        }

        public void CreatePdf(string Filename)
        {
            var catalogLayout = new PdfCatalogLayout(_ws, PageSettings, bounds, fontResources);
            var catalog = AddCatalog(2);
            var pagesLayout = catalogLayout.ChildObjects[0];
            var pages = AddPages();
            AddFontData();
            for (int i = 0; i < pagesLayout.ChildObjects.Count; i++)
            {
                var pageLayout = pagesLayout.ChildObjects[i];
                var page = AddPage(2, new List<int>(), PageSettings);
                CreateStreamContentFromCell(pageLayout, page);
                pages.pageObjectNumbers.Add(page.objectNumber);
            }

            //write to pdf
            crossRefTable = new PdfCrossRefTable();
            string debugString = "";
            //start wring pdf binary
            using (var fs = new FileStream(Filename, FileMode.Create, FileAccess.Write))
            {
                using (var bw = new BinaryWriter(fs, Encoding.ASCII))
                {
                    //Write header
                    bw.Write(Encoding.ASCII.GetBytes(header));
                    debugString += header;
                    //Write body
                    foreach (var pdfobj in body)
                    {
                        crossRefTable.AddPosition(fs.Position);
                        bw.Write(pdfobj.ToPdfBytes());
                        debugString += pdfobj.ToPdfString();
                    }
                    //Write CrossReference
                    crossRefTable.Write(bw, fs.Position, body.Count);
                    debugString += crossRefTable.WriteString(body.Count);
                    // Write trailer
                    PdfTrailer.Write(bw, body.Count, catalog.objectNumber, crossRefTable.StartPosition);
                    debugString += PdfTrailer.WriteString(body.Count, catalog.objectNumber, crossRefTable.StartPosition);
                }
            }
        }

        #region DEBUG


        internal void DrawMarginAndHeaderLines(PdfContentBounds bounds)
        {
            //Bottom line
            DrawLine(PdfColor.Black, 0, bounds.Bottom, PageSettings.PageSize.WidthPu, bounds.Bottom);
            DrawLine(new PdfColor(1, 0, 1), bounds.X, bounds.Bottom, bounds.X + bounds.Width, bounds.Bottom);
            //Top line
            DrawLine(PdfColor.Black, 0, bounds.Top, PageSettings.PageSize.WidthPu, bounds.Top);
            DrawLine(new PdfColor(1, 0, 1), bounds.X, bounds.Top, bounds.X + bounds.Width, bounds.Top);
            //Left line
            DrawLine(PdfColor.Black, bounds.Left, 0, bounds.Left, PageSettings.PageSize.HeightPu);
            DrawLine(new PdfColor(1, 0, 1), bounds.Left, bounds.Y, bounds.Left, bounds.Y + bounds.Height);
            //Right line
            DrawLine(PdfColor.Black, bounds.Right, 0, bounds.Right, PageSettings.PageSize.HeightPu);
            DrawLine(new PdfColor(1, 0, 1), bounds.Right, bounds.Y, bounds.Right, bounds.Y + bounds.Height);
            //Header line
            DrawLine(new PdfColor(1, 0, 1), bounds.Right, bounds.HeaderY, bounds.Left, bounds.HeaderY);
            DrawLine(new PdfColor(1, 0, 1), bounds.CenterHeaderX, bounds.HeaderY, bounds.CenterHeaderX, bounds.Top);
            DrawLine(new PdfColor(1, 0, 1), bounds.RightHeaderX, bounds.HeaderY, bounds.RightHeaderX, bounds.Top);
            //Footer line
            DrawLine(new PdfColor(1, 0, 1), bounds.Right, bounds.FooterY, bounds.Left, bounds.FooterY);
            DrawLine(new PdfColor(1, 0, 1), bounds.CenterFooterX, bounds.FooterY, bounds.CenterFooterX, bounds.Bottom);
            DrawLine(new PdfColor(1, 0, 1), bounds.RightFooterX, bounds.FooterY, bounds.RightFooterX, bounds.Bottom);
        }

        //Might use this for drawing grid later. so might move this.
        internal void DrawLine(PdfColor color, double x1, double y1, double x2, double y2)
        {
            var content = new PdfContentStream(body.Count + 1);
            content.AddCommand(color.ToStrokeCommand());
            content.AddCommand($"{x1.ToPdfString()} {y1.ToPdfString()} m");
            content.AddCommand($"{x2.ToPdfString()} {y2.ToPdfString()} l");
            content.AddCommand("S");
            body.Add(content);
        }

        internal void DrawCrossHair(PdfColor color, double x, double y, double size = 2)
        {
            var half = size / 2d;
            var content = new PdfContentStream(body.Count + 1);
            content.AddCommand(color.ToStrokeCommand());
            content.AddCommand($"{x.ToPdfString()} {(y - half).ToPdfString()} m");
            content.AddCommand($"{x.ToPdfString()} {(y + half).ToPdfString()} l");
            content.AddCommand($"{(x - half).ToPdfString()}   {y.ToPdfString()} m");
            content.AddCommand($"{(x + half).ToPdfString()}   {y.ToPdfString()} l");
            content.AddCommand("S");
            body.Add(content);
        }


        #endregion

    }
}


/* TODO:
 * Print workbook
 * Print worksheets
 * print selected range
 */
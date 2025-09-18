using OfficeOpenXml.PDF.PdfGraphics;
using OfficeOpenXml.PDF.Pdfhelpers;
using OfficeOpenXml.PDF.PdfObjects;
using OfficeOpenXml.PDF.PdfSettings;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml.PDF.PdfLayout;
using OfficeOpenXml.PDF.PdfResources;

namespace OfficeOpenXml.PDF
{
    /// <summary>
    /// Class for exporting to PDF format.
    /// </summary>
    public class ExcelPdf
    {
        internal List<ExcelWorksheet> _workheets = new List<ExcelWorksheet>();
        internal ExcelRangeBase _range;
        private PdfPageSettings PageSettings;
        internal List<PdfObject> Document = new List<PdfObject>();
        internal string header = "%PDF-1.7\n";
        internal readonly Dictionary<string, PdfFontResource> fontResources = new Dictionary<string, PdfFontResource>();
        internal readonly Dictionary<string, PdfPatternResource> patternResources = new Dictionary<string, PdfPatternResource>();

        /// <summary>
        /// Create a PDF Document from the worksheet and settings.
        /// </summary>
        /// <param name="worksheet">The worksheet to convert to PDF Document</param>
        /// <param name="pageSettings">The settings object</param>
        public ExcelPdf(ExcelWorksheet worksheet, PdfPageSettings pageSettings = null)
        {
            _workheets.Add(worksheet);
            PageSettings = pageSettings == null ? new PdfPageSettings() : pageSettings;
            PageSettings.defaultFontName = worksheet.Workbook.ThemeManager.CurrentTheme.FontScheme.MinorFont[0].Typeface;
        }

        /// <summary>
        /// Create a PDF Document from the selected worksheets and settings. NOT IMPLEMENTED
        /// </summary>
        /// <param name="worksheet">The worksheets to convert to PDF Document</param>
        /// <param name="pageSettings">The Settings object</param>
        public ExcelPdf(ExcelWorksheet[] worksheet, PdfPageSettings pageSettings = null)
        {
            //_ws = worksheet[0];
            //defaultFontName = worksheet[0].Workbook.ThemeManager.CurrentTheme.FontScheme.MinorFont[0].Typeface;
            //PageSettings = pageSettings == null ? new PdfPageSettings() : pageSettings;
        }

        /// <summary>
        /// Create a PDF Document from the entire worksbook and settings.NOT IMPLEMENTED
        /// </summary>
        /// <param name="workbook">Workbook to convert to PDF Document</param>
        /// <param name="pageSettings">The settings object</param>
        public ExcelPdf(ExcelWorkbook workbook, PdfPageSettings pageSettings = null)
        {
            //_ws = worksheet;
            //defaultFontName = worksheet.Workbook.ThemeManager.CurrentTheme.FontScheme.MinorFont[0].Typeface;
            //PageSettings = pageSettings == null ? new PdfPageSettings() : pageSettings;
        }

        /// <summary>
        /// Create a PDF Document from the selected range and settings. NOT IMPLEMENTED
        /// </summary>
        /// <param name="Range">Range to convert to PDF Document</param>
        /// <param name="pageSettings">The settings object</param>
        public ExcelPdf(ExcelRangeBase Range, PdfPageSettings pageSettings = null)
        {
            //_range = Range;
            //defaultFontName = Range.Worksheet.Workbook.ThemeManager.CurrentTheme.FontScheme.MinorFont[0].Typeface;
            //PageSettings = pageSettings == null ? new PdfPageSettings() : pageSettings;
        }

        //Get font label //need to update this one too for same reasons as AddFontData
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
                    Document.Add(fr.GetFontDescriptorObject(Document.Count + 1));
                    Document.Add(fr.GetWidthsObject(Document.Count + 1));
                }
                Document.Add(fr.GetFontObject(Document.Count + 1));
                fontResources.Add(fontName, fr);
            }
            return fontResources[fontName].Label;
        }

        //Add Fonts //Need to update this method a bit. We should check for all default fonts and not onlt courier new?
        internal void AddFontData()
        {
            foreach (var font in fontResources)
            {
                if (font.Key != "Courier New")
                {
                    Document.Add(font.Value.GetFontDescriptorObject(Document.Count + 1));
                    Document.Add(font.Value.GetWidthsObject(Document.Count + 1));
                }
                Document.Add(font.Value.GetFontObject(Document.Count + 1));
            }
        }

        //Create Page
        private PdfPage AddPage(int pagesObjectNumber, List<int> contentObjectNumbers, PdfPageSettings settings)
        {
            var page = new PdfPage(Document.Count + 1, pagesObjectNumber, contentObjectNumbers, settings.PageSize, fontResources);
            Document.Add(page);
            return page;
        }
        //Create Pages
        private PdfPages AddPages()
        {
            var pages = new PdfPages(Document.Count + 1, new List<int>{});
            Document.Add(pages);
            return pages;
        }
        //Create Catalog
        private PdfCatalog AddCatalog(int pagesObjectNumber)
        {
            var catalog = new PdfCatalog(Document.Count + 1, pagesObjectNumber);
            Document.Add(catalog);
            return catalog;
        }

        //Create Content
        private void AddContent(PdfTransform pageLayout, PdfPage page)
        {
            var cells = pageLayout.ChildObjects.Where(t => t is PdfCellLayout || t is PdfCellContentLayout || t is PdfCellBorderLayout).GroupBy(t => t.Name);
            var contentStream = new PdfContentStream(Document.Count + 1);
            if (PageSettings.ShowGridLines)
            {
                contentStream.AddInnerGridLines(pageLayout);
            }
            //Add clipping rectangle around page content.
            contentStream.AddCommand("q");
            contentStream.AddMarginClipping(pageLayout, PageSettings.ContentBounds);
            foreach (var cell in cells)
            {
                foreach (var cellPart in cell)
                {
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
                }
            }
            //Close the clipping rectangle
            contentStream.AddCommand("Q");
            if (PageSettings.ShowGridLines)
            {
                contentStream.AddOuterGridBorder(pageLayout);
            }
            Document.Add(contentStream);
            page.contentObjectNumbers.Add(contentStream.objectNumber);
        }

        /// <summary>
        /// Create the pdf from the supplied worksheet.
        /// </summary>
        /// <param name="Filename">The file name</param>
        public void CreatePdf(string Filename)
        {
            //Create Catalog
            var catalogLayout = new PdfCatalogLayout(_workheets[0], PageSettings, fontResources);
            var catalog = AddCatalog(2);
            //Create Pages
            var pagesLayout = catalogLayout.ChildObjects[0];
            var pages = AddPages();
            //Create Fonts
            AddFontData();
            //Create Page and Content
            for (int i = 0; i < pagesLayout.ChildObjects.Count; i++)
            {
                var pageLayout = pagesLayout.ChildObjects[i];
                var page = AddPage(2, new List<int>(), PageSettings);
                AddContent(pageLayout, page);
                pages.pageObjectNumbers.Add(page.objectNumber);
            }
            string debugString = "";
            //write to pdf
            PdfCrossRefTable crossRefTable = new PdfCrossRefTable();
            //start wring pdf binary
            using (var fs = new FileStream(Filename, FileMode.Create, FileAccess.Write))
            {
                using (var bw = new BinaryWriter(fs, Encoding.ASCII))
                {
                    //Write header
                    bw.Write(Encoding.ASCII.GetBytes(header));
                    debugString += header;
                    //Write body
                    foreach (var pdfobj in Document)
                    {
                        crossRefTable.AddPosition(fs.Position);
                        bw.Write(pdfobj.ToPdfBytes());
                        debugString += pdfobj.ToPdfString();
                    }
                    //Write CrossReference
                    crossRefTable.Write(bw, fs.Position, Document.Count);
                    debugString += crossRefTable.WriteString(Document.Count);
                    // Write trailer
                    PdfTrailer.Write(bw, Document.Count, catalog.objectNumber, crossRefTable.StartPosition);
                    debugString += PdfTrailer.WriteString(Document.Count, catalog.objectNumber, crossRefTable.StartPosition);
                }
            }
            //Write pdf as txt for debug.
            if (PageSettings.Debug && PageSettings.PrintAsText)
            {
                using (var fs = new FileStream(Filename + ".txt", FileMode.Create, FileAccess.Write))
                {
                    using ( var wr = new StreamWriter(fs))
                    {
                        wr.Write(debugString);
                    }
                }
            }
        }

        #region DEBUG
        //These methods need to be rewritten if they should be used.

        internal void DrawMarginAndHeaderLines(PdfContentBounds bounds, PdfPage page)
        {
            var content = new PdfContentStream(Document.Count + 1);
            //Bottom line
            DrawLine(content, PdfColor.Black, 0, bounds.Bottom, PageSettings.PageSize.WidthPu, bounds.Bottom);
            DrawLine(content, new PdfColor(1, 0, 1), bounds.X, bounds.Bottom, bounds.X + bounds.Width, bounds.Bottom);
            //Top line
            DrawLine(content, PdfColor.Black, 0, bounds.Top, PageSettings.PageSize.WidthPu, bounds.Top);
            DrawLine(content, new PdfColor(1, 0, 1), bounds.X, bounds.Top, bounds.X + bounds.Width, bounds.Top);
            //Left line
            DrawLine(content, PdfColor.Black, bounds.Left, 0, bounds.Left, 0);
            DrawLine(content, new PdfColor(1, 0, 1), bounds.Left, bounds.Y + bounds.Height, bounds.Left, bounds.Y + bounds.Height);
            //Right line
            DrawLine(content, PdfColor.Black, bounds.Right, 0, bounds.Right, 0);
            DrawLine(content, new PdfColor(1, 0, 1), bounds.Right, bounds.Y + bounds.Height, bounds.Right, bounds.Y + bounds.Height);
            //Header line
            DrawLine(content, new PdfColor(1, 0, 1), bounds.Right, bounds.HeaderY, bounds.Left, bounds.HeaderY);
            DrawLine(content, new PdfColor(1, 0, 1), bounds.CenterHeaderX, bounds.Top, bounds.CenterHeaderX, bounds.Top);
            DrawLine(content, new PdfColor(1, 0, 1), bounds.RightHeaderX, bounds.Top, bounds.RightHeaderX, bounds.Top);
            //Footer line
            DrawLine(content, new PdfColor(1, 0, 1), bounds.Right, bounds.FooterY, bounds.Left, bounds.FooterY);
            DrawLine(content, new PdfColor(1, 0, 1), bounds.CenterFooterX, bounds.Bottom, bounds.CenterFooterX, bounds.Bottom);
            DrawLine(content, new PdfColor(1, 0, 1), bounds.RightFooterX, bounds.Bottom, bounds.RightFooterX, bounds.Bottom);
            Document.Add(content);
            page.contentObjectNumbers.Add(content.objectNumber);
        }

        internal void DrawLine(PdfContentStream content, PdfColor color, double x1, double y1, double x2, double y2)
        {
            content.AddCommand(color.ToStrokeCommand());
            content.AddCommand($"{x1.ToPdfString()} {y1.ToPdfString()} m");
            content.AddCommand($"{x2.ToPdfString()} {y2.ToPdfString()} l");
            content.AddCommand("S");
        }

        internal void DrawCrossHair(PdfColor color, double x, double y, double size = 2)
        {
            var half = size / 2d;
            var content = new PdfContentStream(Document.Count + 1);
            content.AddCommand(color.ToStrokeCommand());
            content.AddCommand($"{x.ToPdfString()} {(y - half).ToPdfString()} m");
            content.AddCommand($"{x.ToPdfString()} {(y + half).ToPdfString()} l");
            content.AddCommand($"{(x - half).ToPdfString()}   {y.ToPdfString()} m");
            content.AddCommand($"{(x + half).ToPdfString()}   {y.ToPdfString()} l");
            content.AddCommand("S");
            Document.Add(content);
        }


        #endregion
    }
}

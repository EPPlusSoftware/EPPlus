using FontLab1.GenericMeasurements;
using FontLab1;
using OfficeOpenXml.PDF.PdfFontData;
using OfficeOpenXml.PDF.PdfGraphics;
using OfficeOpenXml.PDF.Pdfhelpers;
using OfficeOpenXml.PDF.PdfObjects;
using OfficeOpenXml.PDF.PdfSettings;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml.PDF.PdfSettings.PdfPageData;
using FontLab1.Tables.Os2;
using OfficeOpenXml.PDF.PdfLayout;
using System.Runtime;
using OfficeOpenXml.Packaging.Ionic.Zip;
using System.Runtime.InteropServices.ComTypes;

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

        internal void AddText(string text, string cellFontname, double size, double x, double y, PdfPage page)
        {
            var label = GetFontLabel(cellFontname, "Regular", size);
            var content = new PdfContentStream(body.Count + 1);
            content.AddText(label , size, x, y, text);
            body.Add(content);
            page.contentObjectNumbers.Add(content.objectNumber);
        }

        internal void AddRectangle(double x, double y, double width, double height, PdfColor stroke = null, PdfColor fill = null)
        {
            var content = new PdfContentStream(body.Count + 1);
            content.AddRectangle(x, y, width, height, stroke != null ? true : false, fill != null ? true : false, stroke, fill);
            body.Add(content);
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

        private void CreateContentFromCell(PdfTransform child, PdfPage page, PdfContentStream contentStream, string previousName)
        {
            //check child is cell, merged cell or drawing
            if (child is PdfCellLayout cell)
            {
                //add the operations for cell border and fill data
            }
            else if (child is PdfCellContentLayout content)
            {
                //add the operations for text and font data
                AddText(content.FontData.Text, content.FontData.FontName, content.FontData.FontSize, content.LocalPosition.X, content.LocalPosition.Y, page);
            }
            else if (child is PdfMergedCellLayout merged)
            {
                AddRectangle(merged.Position.X, merged.Position.Y, merged.Size.X, merged.Size.Y, null, merged.CellFillData.BackgroundColor);
                //AddText(merged.FontData.Text, merged.FontData.FontName, merged.FontData.FontSize, merged.Position.X, merged.Position.Y, page);
            }
            else if (child is PdfDrawingLayout drawing)
            {
            }
        }

        private void CreateStreamContentFromCell(PdfTransform pageLayout, PdfPage page)
        {
            var cells = pageLayout.ChildObjects.Where(t => t is PdfCellLayout || t is PdfCellContentLayout).GroupBy(t => t.Name);
            foreach (var cell in cells)
            {
                var contentStream = new PdfContentStream(body.Count + 1, "q");
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
                    }
                }
                contentStream.AddCommand("Q");
                body.Add(contentStream);
                page.contentObjectNumbers.Add(contentStream.objectNumber);
            }
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
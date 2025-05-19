using OfficeOpenXml.PDF.PdfGraphics;
using OfficeOpenXml.PDF.PdfObjects;
using OfficeOpenXml.PDF.PdfPageSettings;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF
{
    public class ExcelPdf
    {
        string header = "%PDF-1.7\n";
        List<PdfObject> body = new List<PdfObject>();
        PdfCrossRefTable crossRefTable;

        public readonly Dictionary<int, string> fontResources = new Dictionary<int, string>();


        PdfPageSettings.PdfPageSettings PageSettings;

        public ExcelPdf()
        {
            PageSettings = new PdfPageSettings.PdfPageSettings();
        }


        public ExcelPdf(PdfPageSettings.PdfPageSettings pageSettings)
        {
            PageSettings = pageSettings;
        }

        public void AddFont(string fontName = "Helvetica")
        {
            var font = new PdfFont(body.Count + 1, fontName);
            body.Add(font);
            fontResources.Add(body.IndexOf(font) + 1, "F" + (fontResources.Count + 1));
        }

        public void AddText(string text, string fontResourceName, int size, float x, float y)
        {
            var content = new PdfContentStream(body.Count + 1);
            content.AddText(fontResourceName, size, x, y, text);
            body.Add(content);
        }

        public void AddRectangle(float x, float y, float width, float height, PdfColor stroke = null, PdfColor fill = null)
        {
            var content = new PdfContentStream(body.Count + 1);
            content.AddRectangle(x, y, width, height, stroke != null ? true : false, fill != null ? true : false, stroke, fill);
            body.Add(content);
        }

        //create page
        private PdfPage AddPage(int pagesObjectNumber, List<int> contentObjectNumbers, PdfPageSettings.PdfPageSettings settings)
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


        private void AddWorksheetCells(ExcelWorksheet ws, PdfContentBounds bounds)
        {
            AddFont();
            float prevWidth = 0;
            float prevHeight = 0;
            var x = 0f;
            var y = bounds.Y + bounds.Height;

            for (int i = ws.Dimension._fromRow; i <= ws.Dimension._toRow; i++)
            {
                for (int j = ws.Dimension._fromCol; j <= ws.Dimension._toCol; j++)
                {
                    var cell = ws.Cells[i,j];

                    x = bounds.X + prevWidth;
                    y = bounds.Y + bounds.Height - prevHeight;
                    if (x >= bounds.Width)
                    {
                        prevHeight += (float)cell.Worksheet.Row(1).Height;
                        prevWidth = 0;
                        x = bounds.X + prevWidth;
                        y = bounds.Y + bounds.Height - prevHeight;
                        if (y < bounds.Height)
                        {
                            //new page..
                            break;
                        }
                    }
                    if(cell.Value != null)
                        AddText(cell.Value.ToString(), "F1", (int)cell.Style.Font.Size, x, y);

                    prevWidth += (float)PdfUnits.ExcelColumnWidthToPoints(cell.EntireColumn.Width);
                }
            }
        }

        public void CreatePdf(string Filename, ExcelWorksheet worksheet, PdfPageSettings.PdfPageSettings pageSettings = null)
        {
            if(pageSettings != null)
                PageSettings = pageSettings;

            PdfContentBounds bounds = new PdfContentBounds(PageSettings.Margins, PageSettings.PageSize);

            AddWorksheetCells(worksheet, bounds);

            var pages = AddPages();
            List<int> contentObjectNumbers = new List<int>();
            contentObjectNumbers = body.OfType<PdfContentStream>().Select(con => con.objectNumber).ToList();
            var page = AddPage(pages.objectNumber, contentObjectNumbers, PageSettings);
            pages.pageObjectNumbers.Add(page.objectNumber);
            var catalog = AddCatalog(pages.objectNumber);

            crossRefTable = new PdfCrossRefTable();

            //start wring pdf binary
            using (var fs = new FileStream(Filename, FileMode.Create, FileAccess.Write))
            {
                using (var bw = new BinaryWriter(fs, Encoding.ASCII))
                {
                    //Write header
                    bw.Write(Encoding.ASCII.GetBytes(header));
                    //Write body
                    foreach (var pdfobj in body)
                    {
                        crossRefTable.AddPosition(fs.Position);
                        bw.Write(pdfobj.ToPdfBytes());
                    }
                    //Write CrossReference
                    crossRefTable.Write(bw, fs.Position, body.Count);
                    // Write trailer
                    PdfTrailer.Write(bw, body.Count, catalog.objectNumber, crossRefTable.StartPosition);
                }
            }
        }
    }
}


/* TODO:
 * Print workbook
 * Print worksheets
 * print selected range
 
 */
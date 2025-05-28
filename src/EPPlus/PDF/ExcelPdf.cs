using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements;
using OfficeOpenXml.Core.Worksheet.Fonts.GenericFontMetrics;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.PDF.PdfFontData;
using OfficeOpenXml.PDF.PdfGraphics;
using OfficeOpenXml.PDF.Pdfhelpers;
using OfficeOpenXml.PDF.PdfObjects;
using OfficeOpenXml.PDF.PdfPageSettings;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using static System.Net.Mime.MediaTypeNames;

namespace OfficeOpenXml.PDF
{
    public class ExcelPdf
    {
        string header = "%PDF-1.7\n";
        List<PdfObject> body = new List<PdfObject>();
        PdfCrossRefTable crossRefTable;

        /// <summary>
        /// Key is the font name. Value is a dict where key is font object number and value is the name to use in a text.
        /// </summary>
        public readonly Dictionary<string, Dictionary<int, string>> fontResources = new Dictionary<string, Dictionary<int, string>>();
        private static Dictionary<uint, PdfFontProperties> _fonts;

        PdfPageSettings.PdfPageSettings PageSettings;

        public ExcelPdf()
        {
            PageSettings = new PdfPageSettings.PdfPageSettings();
        }


        public ExcelPdf(PdfPageSettings.PdfPageSettings pageSettings)
        {
            PageSettings = pageSettings;
        }

        public void AddFont(string fontName = "Helvetica", PdfFontSubType fontSubType = PdfFontSubType.Type1, PdfFontEncoding encoding = PdfFontEncoding.WinAnsiEncoding)
        {
            if (!fontResources.ContainsKey(fontName))
            {
                if(_fonts == null)
                    _fonts =  PdfFontMetricsLoader.LoadFontMetrics();

                //_fonts[26]

                if (Enum.IsDefined(typeof(FontMetricsFamilies), fontName.Replace(" ", "")))
                {
                    //var fi = (FontMetricsFamilies)Enum.Parse(typeof(FontMetricsFamilies), fontName.Replace(" ", ""));
                    uint fi = 1703936;
                    var defaults = _fonts[fi].DefaultWidthClass;
                    var fontDescriptor = new PdfFontDescriptor(body.Count + 1, fontName,
                        _fonts[fi].flags,
                        _fonts[fi].fontBBox,
                        _fonts[fi].italicAngle,
                        _fonts[fi].ascent, _fonts[fi].descent,
                        _fonts[fi].stemV,
                        _fonts[fi].capheight);
                    body.Add(fontDescriptor);

                    var Width = new PdfFontWidths(body.Count + 1, _fonts[fi].ClassWidths, _fonts[fi].CharMetrics);
                    body.Add(Width);

                    var font = new PdfFont(body.Count + 1, fontName, fontSubType, _fonts[fi].firstChar, _fonts[fi].lastChar,Width.objectNumber, fontDescriptor.objectNumber, encoding);
                    body.Add(font);
                    fontResources.Add(fontName, new Dictionary<int, string> { { body.IndexOf(font) + 1, "F" + (fontResources.Count + 1) } });
                }
                else
                {
                    //should do fallback font here. Maybe a user setting to throw or use fallback.
                    throw new Exception("This is a temporary exception");
                }
            }
        }

        public void AddText(string text, string cellFontname, float size, double x, double y)
        {
            //check fontname ang get resource name
            AddFont(cellFontname);

            var content = new PdfContentStream(body.Count + 1);
            content.AddText(fontResources[cellFontname].Values.First(), size, x, y, text);
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
            double prevWidth = 0;
            double prevHeight = 0;
            var x = 0d;
            var y = bounds.Y + bounds.Height;

            for (int i = ws.Dimension._fromRow; i <= ws.Dimension._toRow; i++)
            {
                for (int j = ws.Dimension._fromCol; j <= ws.Dimension._toCol; j++)
                {
                    var cell = ws.Cells[i,j];

                    x = bounds.X + prevWidth;
                    y = 775 - prevHeight; //bounds.Y + bounds.Height - prevHeight;
                    if (x >= bounds.Width)
                    {
                        prevHeight += cell.Worksheet.Row(1).Height + 0.25d;
                        prevWidth = 0;
                        x = bounds.X + prevWidth;
                        y = 775 - prevHeight;//bounds.Y + bounds.Height - prevHeight;
                        if (y < bounds.Y)
                        {
                            //new page..
                        }
                    }
                    if (cell.Value != null)
                    {
                        if (cell.Style.HorizontalAlignment == Style.ExcelHorizontalAlignment.General)
                        {
                            if (double.TryParse(cell.Value.ToString(), out double value))
                            {
                                //calculate new x
                                GenericFontMetricsTextMeasurer tm = new GenericFontMetricsTextMeasurer();
                                MeasurementFont font = new MeasurementFont();
                                font.FontFamily = cell.Style.Font.Name;
                                font.Size = cell.Style.Font.Size;
                                font.Style = MeasurementFontStyles.Regular;

                                var values = tm.MeasureIndividualCharacters(cell.Value.ToString(), font, 72);

                                uint sum = 0;

                                foreach (uint val in values)
                                {
                                    sum += val;
                                }
                                var result = tm.MeasureText(cell.Value.ToString(), font);
                                //convert result to points
                                var strWidth = (sum / 1000.0/*units per em*/) * font.Size;

                                x = x + PdfUnits.ExcelColumnWidthToPoints(cell.EntireColumn.Width) - ((double)result.Width - sum);
                            }
                        }
                        AddText(cell.Value.ToString(), cell.Style.Font.Name, cell.Style.Font.Size, x, y);
                    }

                    prevWidth += PdfUnits.ExcelColumnWidthToPoints(cell.EntireColumn.Width);
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
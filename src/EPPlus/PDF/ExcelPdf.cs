using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements;
using OfficeOpenXml.Core.Worksheet.Fonts.GenericFontMetrics;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Database;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.PDF.PdfFontData;
using OfficeOpenXml.PDF.PdfGraphics;
using OfficeOpenXml.PDF.Pdfhelpers;
using OfficeOpenXml.PDF.PdfObjects;
using OfficeOpenXml.PDF.PdfSettings;
using OfficeOpenXml.PDF.PdfSettings.PdfPageSizes;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using static System.Net.Mime.MediaTypeNames;

namespace OfficeOpenXml.PDF
{
    public class ExcelPdf
    {
        internal string header = "%PDF-1.7\n";
        internal List<PdfObject> body = new List<PdfObject>();
        internal PdfCrossRefTable crossRefTable;

        /// <summary>
        ///
        /// </summary>
        internal readonly Dictionary<string, PdfFontResource> fontResources = new Dictionary<string, PdfFontResource>();
        private static Dictionary<uint, PdfFontProperties> _fonts;

        private PdfPageSettings PageSettings;

        public ExcelPdf()
        {
            PageSettings = new PdfPageSettings();
        }


        public ExcelPdf(PdfPageSettings pageSettings)
        {
            PageSettings = pageSettings;
        }

        //public void AddFont(string fontName = "Helvetica", PdfFontSubType fontSubType = PdfFontSubType.Type1, PdfFontEncoding encoding = PdfFontEncoding.WinAnsiEncoding)
        //{
        //    if (!fontResources.ContainsKey(fontName))
        //    {
        //        if(_fonts == null)
        //            _fonts =  PdfFontMetricsLoader.LoadFontMetrics();

        //        //_fonts[26]

        //        if (Enum.IsDefined(typeof(FontMetricsFamilies), fontName.Replace(" ", "")))
        //        {
        //            //var fi = (FontMetricsFamilies)Enum.Parse(typeof(FontMetricsFamilies), fontName.Replace(" ", ""));
        //            uint fi = 1703936;
        //            var defaults = _fonts[fi].DefaultWidthClass;
        //            var fontDescriptor = new PdfFontDescriptor(body.Count + 1, fontName,
        //                _fonts[fi].flags,
        //                _fonts[fi].fontBBox,
        //                _fonts[fi].italicAngle,
        //                _fonts[fi].ascent, _fonts[fi].descent,
        //                _fonts[fi].stemV,
        //                _fonts[fi].capheight);
        //            body.Add(fontDescriptor);

        //            var Width = new PdfFontWidths(body.Count + 1, _fonts[fi].ClassWidths, _fonts[fi].CharMetrics);
        //            body.Add(Width);

        //            var font = new PdfFont(body.Count + 1, fontName, fontSubType, _fonts[fi].firstChar, _fonts[fi].lastChar,Width.objectNumber, fontDescriptor.objectNumber, encoding);
        //            body.Add(font);
        //            fontResources.Add(fontName, new Dictionary<int, string> { { body.IndexOf(font) + 1, "F" + (fontResources.Count + 1) } });
        //        }
        //        else
        //        {
        //            //should do fallback font here. Maybe a user setting to throw or use fallback.
        //            throw new Exception("This is a temporary exception");
        //        }
        //    }
        //}

        public void AddText(string text, string cellFontname, float size, double x, double y)
        {
            var content = new PdfContentStream(body.Count + 1);
            content.AddText(fontResources[cellFontname].labelPrefix + fontResources[cellFontname].label , size, x, y, text);
            body.Add(content);
        }

        public void AddRectangle(double x, double y, double width, double height, PdfColor stroke = null, PdfColor fill = null)
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

        private double MeasureString(string text, string fontName, string subFamily, double fontSize)
        {
            if(!fontResources.ContainsKey(fontName))
            {
                PdfFontResource fr = new PdfFontResource(fontName, subFamily, fontResources.Last().Value.label, PageSettings);
                body.Add(fr.GetFontDescriptorObject(body.Count+1));
                body.Add(fr.GetWidthsObject(body.Count + 1));
                body.Add(fr.GetFontObject(body.Count + 1));
                return fontResources[fontName].MeasureText(text, fontSize);
            }
            else
            {
                return fontResources[fontName].MeasureText(text, fontSize);
            }
        }

        private void AddWorksheetCells(ExcelWorksheet ws, PdfContentBounds bounds)
        {
            double prevWidth = 0;
            double prevHeight = 0;
            var x = 0d;
            var y = bounds.Y + bounds.Height;
            PdfRect contentRect = new PdfRect();
            contentRect.X = bounds.X;
            contentRect.Y = bounds.Y;
            for (int i = ws.Dimension._fromRow; i <= ws.Dimension._toRow; i++)
            {
                for (int j = ws.Dimension._fromCol; j <= ws.Dimension._toCol; j++)
                {
                    bool textWasAdded = false;
                    var cell = ws.Cells[i,j];
                    x = bounds.X + prevWidth;
                    y = bounds.Top - prevHeight; //bounds.Y + bounds.Height - prevHeight;
                    if (x >= bounds.Width)
                    {
                        prevHeight += cell.Worksheet.Row(i).Height + 0.25d;
                        contentRect.Height = prevHeight;
                        prevWidth = 0;
                        x = bounds.X + prevWidth;
                        y = bounds.Top - prevHeight;//bounds.Y + bounds.Height - prevHeight;
                        if (y < bounds.Y)
                        {
                            //new page..
                        }
                    }
                    if (cell.Value != null)
                    {
                        var textX = x;
                        var textY = y;
                        //check and measure content:
                        string subFamily = cell.Style.Font.Bold ? (cell.Style.Font.Italic ? "Bold Italic" : "Bold") : (cell.Style.Font.Italic ? "Italic" : "Regular");
                        var textLength = MeasureString(cell.Value.ToString(), cell.Style.Font.Name, subFamily, cell.Style.Font.Size);

                        if (cell.Style.HorizontalAlignment == Style.ExcelHorizontalAlignment.General)
                        {
                            if (double.TryParse(cell.Value.ToString(), out double value))
                            {
                                ////calculate new x
                                //GenericFontMetricsTextMeasurer tm = new GenericFontMetricsTextMeasurer();
                                //MeasurementFont font = new MeasurementFont();
                                //font.FontFamily = cell.Style.Font.Name;
                                //font.Size = cell.Style.Font.Size;
                                //font.Style = MeasurementFontStyles.Regular;

                                //var values = tm.MeasureIndividualCharacters(cell.Value.ToString(), font, 72);

                                //uint sum = 0;

                                //foreach (uint val in values)
                                //{
                                //    sum += val;
                                //}
                                //var result = tm.MeasureText(cell.Value.ToString(), font);
                                ////convert result to points
                                //var strWidth = (sum / 1000.0/*units per em*/) * font.Size;

                                //textX = x + PdfUnits.ExcelColumnWidthToPoints(cell.EntireColumn.Width) - ((double)result.Width - sum);
                            }
                        }
                        AddText(cell.Value.ToString(), cell.Style.Font.Name, cell.Style.Font.Size, textX, textY);
                        textWasAdded = true;
                    }
                    //if (PageSettings.ShowGridLines)
                    //{
                    //    //get cell width
                    //    var width = PdfUnits.ExcelColumnWidthToPoints(cell.EntireColumn.Width);
                    //    var height = cell.Worksheet.Row(i).Height + 0.25d;
                    //    var rectY = y - (height / 4d); //move rectnagle in y one fourth up to center text insice grid rectangle
                    //    var rectX = x; //hardcoded 2. Should probably calculate padding based on cell width.
                    //    if (textWasAdded)
                    //    {
                    //        var textObj = (PdfContentStream)body.Last();
                    //        textObj.AddCommand(PdfColor.LightGray.ToStrokeCommand());
                    //        textObj.AddCommand($"{rectX.ToString("F", CultureInfo.InvariantCulture)} {rectY.ToString("F", CultureInfo.InvariantCulture)} {width.ToString("F", CultureInfo.InvariantCulture)} {height.ToString("F", CultureInfo.InvariantCulture)} re");
                    //        textObj.AddCommand("S");
                    //    }
                    //    else
                    //    {
                    //        AddRectangle(rectX, rectY, width, height, PdfColor.LightGray);
                    //    }
                    //    if (i == ws.Dimension._fromRow && j == ws.Dimension._fromCol)
                    //    {
                    //        contentRect.X = rectX;
                    //        contentRect.Y = rectY + ws.Row(ws.Dimension._fromRow).Height + 0.25d; ;
                    //    }
                    //}
                    prevWidth += PdfUnits.ExcelColumnWidthToPoints(cell.EntireColumn.Width);
                    contentRect.Width = prevWidth;
                }
            }
            //if (PageSettings.ShowGridLines)
            //{
            //    contentRect.Height += ws.Row(ws.Dimension._fromRow).Height + 0.25d; //First row is skipped in algorithm when calculating so we add it here.
            //    AddRectangle(contentRect.X, contentRect.Y - contentRect.Height, contentRect.Width, contentRect.Height, PdfColor.Black);
            //}
        }

        public void CreatePdf(string Filename, ExcelWorksheet worksheet, PdfPageSettings pageSettings = null)
        {
            if(pageSettings != null)
                PageSettings = pageSettings;

            PdfContentBounds bounds = new PdfContentBounds(PageSettings.Margins, PageSettings.PageSize);

            AddWorksheetCells(worksheet, bounds);

            //draw cell contents and headings
            //draw grid
            //draw drawings
            //Draw header and footer
            //

            //Debug
            if (PageSettings.Debug)
            {
                DrawMarginAndHeaderLines(bounds);
            }

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
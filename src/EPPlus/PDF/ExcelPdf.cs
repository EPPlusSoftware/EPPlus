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

        /// <summary>
        ///
        /// </summary>
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
            content.AddText(fontResources[cellFontname].labelPrefix + fontResources[cellFontname].labelNumber , size, x, y, text);
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
                int label = 1;
                if(fontResources.Count > 0)
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
                return fontResources[fontName].MeasureText(text, fontSize);
            }
            else
            {
                return fontResources[fontName].MeasureText(text, fontSize);
            }
        }

        internal double CalculateDefaultRowHeight(ExcelWorksheet ws)
        {
            TtfFont font = GenericFonts.GetFontData(PageSettings, defaultFontName);
            double ascender = 0, descender = 0, lineGap = 0, size = 11, em = font.HeadTable.UnitsPerEm;
            //if (PdfExcelFontDataLookup.ExcelFontData.ContainsKey(font.FullName))
            //{
                
            //}
            //else
            //{
                if ((font.Os2Table.SelectionFlags & Os2Table.FsSelectionFlags.UseTypoMetrics) != 0)
                {
                    ascender = font.Os2Table.sTypoAscender;
                    descender = font.Os2Table.sTypoDescender;
                    lineGap = font.Os2Table.sTypoLineGap;
                }
                else
                {
                    ascender = font.Os2Table.usWinAscent;
                    descender = font.Os2Table.usWinDescent;
                    lineGap = 0;
                }
            //}
            var lineHeight = ascender + System.Math.Abs(descender) + lineGap;
            var lineHeightPt = lineHeight * (size / em);
            var lineHeightPad = lineHeightPt + 1d;
            return lineHeightPad;
        }

        private double CenterContentsOnPage(double pageheight, double pageStartY, double rangeHeight )
        {
            return pageStartY + (pageheight - rangeHeight) / 2d;
        }

        internal List<PdfExcelPageData> CreatePageData()
        {
            var ranges = GetRangeForPages();
            pagesData[0].PageRange = ranges[0];
            for (int i = 1; i<ranges.Count;i++)
            {
                pagesData.Add(new PdfExcelPageData(pagesData[0].RowMaxCount, pagesData[0].ColumnMaxCount));
                pagesData[i].PageRange = ranges[i];
            }
            foreach(var pageData in pagesData)
            {
                GetContentSizeAndCoords(pageData);
            }
            return pagesData;
        }

        internal void GetContentSizeAndCoords(PdfExcelPageData pageData)
        {
            double height = 0d;
            double width = 0d;
            for (int i = pageData.PageRange._fromRow; i <= pageData.PageRange._toRow; i++)
            {
                height += _ws.Row(i).Hidden ? 0d : _ws.Row(i).Height;
            }
            for (int j = pageData.PageRange._fromCol; j <= pageData.PageRange._toCol; j++)
            {
                width += PdfUnits.ExcelColumnWidthToPoints(_ws.Column(j).Width);
            }
            pageData.contentWidth = width;
            pageData.contentHeight = height;

            double x = PageSettings.CenterOnPageHorizontally ? (bounds.X + (bounds.Width - width) / 2d) : bounds.Left;
            double y = PageSettings.CenterOnPageVertically ? (bounds.Y + (bounds.Height - height) / 2d) + height : bounds.Top;
            pageData.rowLineCoords.Add([x, y]);
            pageData.colLineCoords.Add([x, y]);
            double currentX = x;
            double currentY = y;
            for (int i = pageData.PageRange._fromRow; i <= pageData.PageRange._toRow; i++)
            {
                currentY += _ws.Row(i).Hidden ? 0d : _ws.Row(i).Height;
                pageData.rowLineCoords.Add([x, currentY]);
            }
            for (int j = pageData.PageRange._fromCol; j <= pageData.PageRange._toCol; j++)
            {
                currentX += PdfUnits.ExcelColumnWidthToPoints(_ws.Column(j).Width);
                pageData.colLineCoords.Add([currentX, y]);
            }
        }

        /// <summary>
        /// Calculates the excel range for each page and the start position on page.
        /// </summary>
        /// <returns></returns>
        internal List<ExcelRangeBase> GetRangeForPages()
        {
            List<string> RowPages = new List<string>();
            List<string> ColPages = new List<string>();
            List<ExcelRangeBase> RangePages = new List<ExcelRangeBase>();
            var y = 0d;
            var x = 0d;
            var rowStart = 1;
            var rowCount = 0;
            var colStart = 1;
            var colCount = 0;
            for (int i = rowStart; i <= _ws.Dimension._toRow; i++)
            {
                if (i == _ws.Dimension._toRow)
                {
                    RowPages.Add(rowStart + ":" + i);
                    break;
                }
                rowCount++;
                var nextRow = i + 1;
                var rowHeight = _ws.Row(i).Hidden ? 0d : _ws.Row(i).Height;
                var nextRowHeight = _ws.Row(nextRow).Hidden ? 0d : _ws.Row(nextRow).Height;
                var currentHeight = y + rowHeight;
                var nextHeight = currentHeight + nextRowHeight;
                if (nextHeight >= bounds.Height || rowCount >= pagesData[0].RowMaxCount)
                {
                    RowPages.Add(rowStart + ":" + i);
                    rowStart = nextRow;
                    y = 0;
                    rowCount = 0;
                }
                else
                {
                    y = currentHeight;
                }
            }
            for (int j = colStart; j <= _ws.Dimension._toCol; j++)
            {
                if (j == _ws.Dimension._toCol)
                {
                    ColPages.Add(colStart + ":" + j);
                    break;
                }
                colCount++;
                var nextCol = j + 1;
                var colWidth = _ws.Column(j).Hidden ? 0d : PdfUnits.ExcelColumnWidthToPoints(_ws.Column(j).Width);
                var nextColWidth = _ws.Column(nextCol).Hidden ? 0d : PdfUnits.ExcelColumnWidthToPoints(_ws.Column(nextCol).Width);
                var currentWidth = x + colWidth;
                var nextWidth = currentWidth + nextColWidth;
                if (nextWidth >= bounds.Width || colCount >= pagesData[0].ColumnMaxCount)
                {
                    ColPages.Add(colStart + ":" + j);
                    colStart = nextCol;
                    x = 0;
                    colCount = 0;
                }
                else
                {
                    x = currentWidth;
                }
            }
            if (PageSettings.PageOrders == PageOrders.DownThenOver)
            {
                for (int j = 0; j < ColPages.Count; j++)
                {
                    var colPage = ColPages[j].Split(':');
                    for (int i = 0; i < RowPages.Count; i++)
                    {
                        var rowPage = RowPages[i].Split(':');
                        string cell1 = ExcelCellAddress.GetColumnLetter(int.Parse(colPage[0])) + rowPage[0];
                        string cell2 = ExcelCellAddress.GetColumnLetter(int.Parse(colPage[1])) + rowPage[1];
                        var range = new ExcelRangeBase(_ws, cell1 + ":" + cell2);
                        if (!range.IsEmpty())
                        {
                            RangePages.Add(range);
                        }
                    }
                }
            }
            else if (PageSettings.PageOrders == PageOrders.OverThenDown)
            {
                for (int i = 0; i < RowPages.Count; i++)
                {
                    var rowPage = RowPages[i].Split(':');
                    for (int j = 0; j < ColPages.Count; j++)
                    {
                        var colPage = ColPages[j].Split(':');
                        string cell1 = ExcelCellAddress.GetColumnLetter(int.Parse(colPage[0])) + rowPage[0];
                        string cell2 = ExcelCellAddress.GetColumnLetter(int.Parse(colPage[1])) + rowPage[1];
                        var range = new ExcelRangeBase(_ws, cell1 + ":" + cell2);
                        if(!range.IsEmpty())
                        {
                            RangePages.Add(range);
                        }
                    }
                }
            }
            return RangePages;
        }

        private void CreatePages()
        {
            double prevWidth = 0;
            double prevHeight = _ws.Row(1).Height + cellMargin;

            foreach (var page in pagesData)
            {
                var x = page.colLineCoords[0];
                var y = page.rowLineCoords[0];
                for (int i = page.PageRange._fromRow; i <= page.PageRange._toRow; i++)
                {
                    for (int j = page.PageRange._fromCol; j <= page.PageRange._toCol; j++)
                    {
                        var cell = _ws.Cells[i, j];
                        var wdith = PdfUnits.ExcelColumnWidthToPoints(_ws.Column(j).Width);
                        var height = _ws.Row(i).Height;
                        if(cell.Value != null)
                        {
                            //get text in cell and measure it
                            //check wrapped text
                            //if wrapped, divide length wid cell width and make sure words fit in row
                            //if not wrapped then check cells to the right until a cell with value is found. cut of text. also need to remember if text is overlapping to another page

                            //2572
                            //Rethink placement and pages. Create a global coordinate system that contains pages in a grid. each pages has its own local space where we write the actual text, but placing stuff happens globally.
                            //a string that overlaps 2 pages checks where the string breaks and creates a text object for each page. We then convert from global space to each pages local page
                            
                            //centering means always centering
                            //but we need to take empty cells into consideration that comes before cells with content. We only stop at the final row or column with content. 
                            //we also include empty pages
                            //Option to diregard empty pages

                            //AddText(cell.Value.ToString(), cell.Style.Font.Name, cell.Style.Font.Size, textX, textY);
                        }
                    }
                }
                //Next page setup
            }
        }

        //old
        private void AddWorksheetCells()
        {
            double prevWidth = 0;
            double prevHeight = _ws.Row(1).Height + cellMargin;
            var x = 0d;
            var y = bounds.Y + bounds.Height;
            PdfRect contentRect = new PdfRect();
            contentRect.X = bounds.X;
            contentRect.Y = bounds.Y;

            /*
             tänk om denna del. antingen tar vi ut en stor range som vi går igenom likadant som denna.
            eller addedar vi ihop kolumners längder tills vi når mållängden, gör samma för rader. och sedan går vi cell för cell som nedan. sparar den rad och kolumn vi är på och repeterar sedan därifrån vid ny sida.
            eller så kikar vi bara de celler vi har värden i. om cellen är t ex d4 så räknar vi från den cellen antalet celler innan kolumnvis och adderar deras bredder och sedan samma för rader.
             */

            for (int i = _ws.Dimension._fromRow; i <= _ws.Dimension._toRow; i++)
            {
                for (int j = _ws.Dimension._fromCol; j <= _ws.Dimension._toCol; j++)
                {
                    bool textWasAdded = false;
                    var cell = _ws.Cells[i,j];
                    x = bounds.X + prevWidth;
                    y = bounds.Top - prevHeight; //bounds.Y + bounds.Height - prevHeight;
                    if (x >= bounds.Width)
                    {
                        prevHeight += cell.Worksheet.Row(i).Height + cellMargin;
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
                        var textLength = MeasureString(cell.Value.ToString(), cell.Style.Font.Name, "Regular", cell.Style.Font.Size);

                        if (cell.Style.HorizontalAlignment == Style.ExcelHorizontalAlignment.General)
                        {
                            if (double.TryParse(cell.Value.ToString(), out double value))
                            {
                                //calculate new x
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
                                //var len = ((double)result.Width - sum)

                                textX = x + PdfUnits.ExcelColumnWidthToPoints(cell.EntireColumn.Width) - textLength;
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

        public void CreatePdf(string Filename, PdfPageSettings pageSettings = null)
        {
            if(pageSettings != null)
                PageSettings = pageSettings;

            CreatePageData();
            AddWorksheetCells();


            //Need
            //[X]range of cells for page                       //ExcelRange for all cells in a page
            //[X]dimensions of cells for page                  //ExcelRange that starts at first cell in page and goes to the last with data.
            //[]grid length and height                         //calculated from dimensions
            //[]grid length and height per page                //A collection per page of coordinates
            //[x]grid Y position                               //calulated based on height and check if start from top or place in middle from PageSettings
            //[]position for each grid divider                 //dictionary for containgn column and row as key and value for position
            //[]position of text relative to cell coordinates  //A padding value for all cells to place text.
            //
            //How to handle merged cell?
            //Drawings are drawn on top
            //Borders
            //in cell pictures
            //

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
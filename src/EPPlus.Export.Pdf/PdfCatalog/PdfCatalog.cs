using EPPlus.Export.Pdf.PdfResources;
using EPPlus.Export.Pdf.PdfSettings;
using EPPlus.Graphics;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace EPPlus.Export.Pdf.PdfCatalog
{
    public class PdfCatalog
    {
        internal PdfDictionaries Dictionaries = new PdfDictionaries();
        private bool AddTextForHeadings = true;

        //Constructors
        public PdfCatalog() { }

        public PdfCatalog(PdfPageSettings pageSettings, ExcelWorkbook workbook)
        {
            HandleWorksheetCollection(pageSettings, workbook.Worksheets.ToArray());
        }

        public PdfCatalog(PdfPageSettings pageSettings, ExcelWorksheet[] worksheets)
        {
            HandleWorksheetCollection(pageSettings, worksheets);
        }

        public PdfCatalog(PdfPageSettings pageSettings, List<ExcelWorksheet> worksheets)
        {
            HandleWorksheetCollection(pageSettings, worksheets.ToArray());
        }

        private void HandleWorksheetCollection(PdfPageSettings pageSettings, ExcelWorksheet[] worksheets)
        {
            var pdfSheets = GetPdfWorksheets(pageSettings, worksheets);
            foreach (var pdfSheet in pdfSheets)
            {
                ShapeTextInPdfWorksheet(pageSettings, pdfSheet);
            }
        }

        public PdfCatalog(PdfPageSettings pageSettings, ExcelWorksheet worksheet, string fileName)
        {
            pageSettings.defaultFontName = worksheet.Workbook.ThemeManager.CurrentTheme.FontScheme.MinorFont[0].Typeface;

            Stopwatch sw = Stopwatch.StartNew();

            //Collect Text
            PdfWorksheet pdfSheet = GetPdfWorksheet(pageSettings, worksheet);
            sw.Stop();
            var CollectTextTime = sw.ElapsedMilliseconds;
            sw.Reset();
            sw.Start();

            //Shape Text
            ShapeTextInPdfWorksheet(pageSettings, pdfSheet);
            sw.Stop();
            var ShapeTextTime = sw.ElapsedMilliseconds;
            sw.Reset();
            sw.Start();

            //Auto-Fit Rows
            PdfCalculateRowHeight.ResizeRowHeights(pdfSheet);
            sw.Stop();
            var AutoFitRowTime = sw.ElapsedMilliseconds;
            sw.Reset();
            sw.Start();

            //Create Layout
            var Layout = GetLayout(pageSettings, pdfSheet);
            sw.Stop();
            var CreateLayoutTime = sw.ElapsedMilliseconds;
            sw.Reset();
            sw.Start();

            //Create Pdf
            ExcelPdf excelPdf = new ExcelPdf();
            excelPdf.CreatePdf(pageSettings, Dictionaries, Layout, fileName);
            sw.Stop();
            var CreatePdfTime = sw.ElapsedMilliseconds;
            sw.Reset();
        }

        public PdfCatalog(PdfPageSettings pageSettings, ExcelRangeBase range)
        {
            PdfWorksheet pdfSheet = GetPdfWorksheet(pageSettings, range);
            ShapeTextInPdfWorksheet(pageSettings, pdfSheet);
        }

        internal PdfCellCollection GetCellCollectionFromRange(PdfPageSettings pageSettings, ExcelRangeBase range)
        {
            PdfWorksheet pdfSheet = GetPdfWorksheet(pageSettings, range);
            ShapeTextInPdfWorksheet(pageSettings, pdfSheet);
            return pdfSheet.Ranges[0].Map;
        }

        //Private Methods
        //Create Layout Methods
        private Transform GetLayout(PdfPageSettings pageSettings, PdfWorksheet pdfSheet)
        {
            PdfWorksheet[] pdfSheets = new PdfWorksheet[1] { pdfSheet };
            var Layout = PdfLayout.GetLayout(pageSettings, Dictionaries, pdfSheets);
            return Layout;
        }

        //Shape Text Methods
        internal void ShapeTextInPdfWorksheet(PdfPageSettings pageSettings, PdfWorksheet pdfSheet)
        {
            // Pass 1: collect text per font
            IterateCells(pdfSheet, cell => PdfTextShaper.CollectText(Dictionaries, cell));

            // Pass 2: build one provider per font
            foreach (var kvp in Dictionaries.Fonts)
            {
                Dictionaries.ShapedProviders[kvp.Key] = kvp.Value.fontSubsetManager.CreateSubsettedProvider();
            }

            // Pass 3: shape text using the pre-built providers
            IterateCells(pdfSheet, cell => PdfTextShaper.ShapeText(pageSettings, Dictionaries, cell));
        }

        private void IterateCells(PdfWorksheet pdfSheet, System.Action<PdfCell> action)
        {
            foreach (var range in pdfSheet.Ranges)
            {
                for (int i = range.Map.FromRow; i <= range.Map.ToRow; i++)
                {
                    for (int j = range.Map.FromColumn; j <= range.Map.ToColumn; j++)
                    {
                        action(range.Map[i, j]);
                    }
                }
            }

            if (pdfSheet.CommentsAndNotes.Map != null)
            {
                for (int i = pdfSheet.CommentsAndNotes.Map.FromRow; i <= pdfSheet.CommentsAndNotes.Map.ToRow; i++)
                {
                    for (int j = pdfSheet.CommentsAndNotes.Map.FromColumn; j <= pdfSheet.CommentsAndNotes.Map.ToColumn; j++)
                    {
                        action(pdfSheet.CommentsAndNotes.Map[i, j]);
                    }
                }
            }
        }

        //Collect Text Methods
        private PdfWorksheet[] GetPdfWorksheets(PdfPageSettings pageSettings, ExcelWorksheet[] worksheets)
        {
            PdfWorksheet[] pdfSheets = new PdfWorksheet[worksheets.Length];
            for (int i = 0; i < pdfSheets.Length; i++)
            {
                pdfSheets[i] = GetPdfWorksheet(pageSettings, worksheets[i]);
            }
            return pdfSheets;
        }

        internal PdfWorksheet GetPdfWorksheet(PdfPageSettings pageSettings, ExcelWorksheet worksheet)
        {
            PdfWorksheet pdfSheet = new PdfWorksheet();
            pdfSheet.Ranges = new List<PdfRange>();
            pdfSheet.Worksheet = worksheet;
            pdfSheet.Ranges = GetRanges(pdfSheet.Worksheet);
            if (pageSettings.ShowHeadings && AddTextForHeadings) Dictionaries.AddFont(pageSettings, pdfSheet.NormalStyle.Style.Font.Name, pdfSheet.GetSubFamilyFromNormalStyle, "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890");
            AddTextForHeadings = false;
            GetMaps(pageSettings, pdfSheet, pdfSheet.Ranges);
            GetHeaderFooter(pageSettings, pdfSheet);
            GetCommentsAndNotes(pageSettings, pdfSheet);
            return pdfSheet;
        }

        private PdfWorksheet GetPdfWorksheet(PdfPageSettings pageSettings, ExcelRangeBase excelRange)
        {
            PdfWorksheet pdfSheet = new PdfWorksheet();
            pdfSheet.Ranges = new List<PdfRange>();
            pdfSheet.Worksheet = excelRange.Worksheet;
            pdfSheet.Ranges.Add(new PdfRange(excelRange, false));
            if (pageSettings.ShowHeadings && AddTextForHeadings) Dictionaries.AddFont(pageSettings, pdfSheet.NormalStyle.Style.Font.Name, pdfSheet.GetSubFamilyFromNormalStyle, "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890");
            AddTextForHeadings = false;
            pdfSheet.Ranges[0] = GetMaps(pageSettings, pdfSheet, pdfSheet.Ranges[0]);
            GetHeaderFooter(pageSettings, pdfSheet);
            GetCommentsAndNotes(pageSettings, pdfSheet);
            return pdfSheet;
        }

        private List<PdfRange> GetRanges(ExcelWorksheet worksheet)
        {
            List<PdfRange> ranges = new List<PdfRange>();
            if (worksheet.Names.ContainsKey("_xlnm.Print_Area"))
            {
                for (int i = 0; i < worksheet.Names["_xlnm.Print_Area"].Addresses.Count; i++)
                {
                    var range = worksheet.Cells[worksheet.Names["_xlnm.Print_Area"].Addresses[i].Address];
                    ranges.Add(new PdfRange(range, false));
                }
            }
            else
            {
                var range = worksheet.Dimension;
                var pdfRange = new PdfRange(range, true);
                pdfRange.ExtendColumns = true;
                ranges.Add(pdfRange);
            }
            return ranges;
        }

        private void GetMaps(PdfPageSettings pageSettings, PdfWorksheet pdfSheet, List<PdfRange> ranges)
        {
            for (int i = 0; i < ranges.Count; i++)
            {
                ranges[i] = GetMaps(pageSettings, pdfSheet, ranges[i]);
            }
        }

        private PdfRange GetMaps(PdfPageSettings pageSettings, PdfWorksheet pdfSheet, PdfRange range)
        {
            var temp = range;
            temp.Map = PdfTextMap.SetTextMap(pageSettings, Dictionaries, pdfSheet, ref temp);
            range = temp;
            return range;
        }

        private void GetHeaderFooter(PdfPageSettings pageSettings, PdfWorksheet pdfSheet)
        {
            pdfSheet.HeaderFooters = new PdfHeaderFooterCollection(pageSettings, Dictionaries, pdfSheet, pdfSheet.Worksheet.HeaderFooter);
        }

        private void GetCommentsAndNotes(PdfPageSettings pageSettings, PdfWorksheet pdfSheet)
        {
            if (pageSettings.CommentsAndNotes == CommentsAndNotes.AtEndOfSheet && pdfSheet.CommentsAndNotesCollections.Count > 0)
            {
                var cnPageSettings = new PdfPageSettings();
                cnPageSettings.CommentsAndNotes = CommentsAndNotes.None;
                pdfSheet.CommentsAndNotesSheet = PdfCommentsAndNotes.CreateCommentAndNotesPages(pdfSheet.CommentsAndNotesCollections, pdfSheet.Worksheet);
                pdfSheet.CommentsAndNotes = new PdfRange(pdfSheet.CommentsAndNotesSheet.Dimension, false);
                pdfSheet.CommentsAndNotes = GetMaps(cnPageSettings, pdfSheet, pdfSheet.CommentsAndNotes);
            }
        }
    }
}
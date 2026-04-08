using EPPlus.Export.Pdf.PdfResources;
using EPPlus.Export.Pdf.PdfSettings;
using EPPlus.Graphics;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace EPPlus.Export.Pdf.PdfCatalog
{
    internal class PdfCatalog
    {
        internal PdfDictionaries Dictionaries = new PdfDictionaries();
        private bool AddTextForHeadings = true;

        //Constructors
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

        public PdfCatalog(PdfPageSettings pageSettings, ExcelWorksheet worksheet)
        {
            Stopwatch sw = Stopwatch.StartNew();

            //Collecto Text
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

            //Create Layout
            GetLayout(pageSettings, pdfSheet);

            sw.Stop();
            var CreateLayoutTime = sw.ElapsedMilliseconds;
            sw.Reset();
        }

        public PdfCatalog(PdfPageSettings pageSettings, ExcelRangeBase range)
        {
            PdfWorksheet pdfSheet = GetPdfWorksheet(pageSettings, range);
            ShapeTextInPdfWorksheet(pageSettings, pdfSheet);
        }

        //Private Methods
        //Create Layout Methods
        private Transform GetLayout(PdfPageSettings pageSettings, PdfWorksheet pdfSheet)
        {
            PdfWorksheet[] pdfSheets = new PdfWorksheet[1]{ pdfSheet };
            PdfLayout.GetLayout(pageSettings, pdfSheets);
            return null;
        }

        //Shape Text Methods
        private void ShapeTextInPdfWorksheet(PdfPageSettings pageSettings, PdfWorksheet pdfSheet)
        {
            foreach (var range in pdfSheet.Ranges)
            {
                for (int i = range.Map.FromRow; i < range.Map.ToRow; i++)
                {
                    for (int j = range.Map.FromColumn; j < range.Map.ToColumn; j++)
                    {
                        var cell = range.Map[i, j];
                        PdfTextShaper.LayoutAndShapeText(pageSettings, Dictionaries, cell);
                    }
                }
            }
            for (int i = pdfSheet.CommentsAndNotes.Map.FromRow; i < pdfSheet.CommentsAndNotes.Map.ToRow; i++)
            {
                for (int j = pdfSheet.CommentsAndNotes.Map.FromColumn; j < pdfSheet.CommentsAndNotes.Map.ToColumn; j++)
                {
                    var cell = pdfSheet.CommentsAndNotes.Map[i, j];
                    PdfTextShaper.LayoutAndShapeText(pageSettings, Dictionaries, cell);
                }
            }
            foreach (var hf in pdfSheet.HeaderFooters.PdfHeaderFooterEntries)
            {
                PdfTextShaper.LayoutAndShapeText(pageSettings, Dictionaries, hf.Content);
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

        private PdfWorksheet GetPdfWorksheet(PdfPageSettings pageSettings, ExcelWorksheet worksheet)
        {
            PdfWorksheet pdfSheet = new PdfWorksheet();
            pdfSheet.Ranges = new List<PdfRange>();
            pdfSheet.Worksheet = worksheet;
            pdfSheet.Ranges = GetRanges(pdfSheet.Worksheet);
            pdfSheet.HeaderFooters = new PdfHeaderFooterCollection(pageSettings, Dictionaries, pdfSheet, pdfSheet.Worksheet.HeaderFooter);
            if(pageSettings.ShowHeadings && AddTextForHeadings) Dictionaries.AddFont(pageSettings, pdfSheet.NormalStyle.Style.Font.Name, pdfSheet.GetSubFamilyFromNormalStyle, "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890");
            AddTextForHeadings = false;
            GetMaps(pageSettings, pdfSheet, pdfSheet.Ranges);
            GetCommentsAndNotes(pageSettings, pdfSheet);
            return pdfSheet;
        }

        private PdfWorksheet GetPdfWorksheet(PdfPageSettings pageSettings, ExcelRangeBase excelRange)
        {
            PdfWorksheet pdfSheet = new PdfWorksheet();
            pdfSheet.Ranges = new List<PdfRange>();
            pdfSheet.Worksheet = excelRange.Worksheet;
            pdfSheet.Ranges.Add(new PdfRange(excelRange, true));
            if (pageSettings.ShowHeadings && AddTextForHeadings) Dictionaries.AddFont(pageSettings, pdfSheet.NormalStyle.Style.Font.Name, pdfSheet.GetSubFamilyFromNormalStyle, "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890");
            AddTextForHeadings = false;
            GetMaps(pageSettings, pdfSheet, pdfSheet.Ranges[0]);
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
                ranges.Add(new PdfRange(range, true));
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

using EPPlus.Export.Pdf.PdfLayout;
using EPPlus.Export.Pdf.PdfResources;
using EPPlus.Export.Pdf.PdfSettings;
using EPPlus.Export.Pdf.PdfSettings.PdfPageSizes;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace EPPlus.Export.Pdf.PdfCatalog
{
    internal class PdfCatalog
    {
        internal PdfDictionaries Dictionaries = new PdfDictionaries();

        public PdfCatalog(PdfPageSettings pageSettings, ExcelRangeBase range) 
        {
            PdfWorksheet pdfSheet = new PdfWorksheet();
            pdfSheet.Ranges = new List<PdfRange>[1];
            if (pageSettings.CommentsAndNotes == CommentsAndNotes.AtEndOfSheet)
                pdfSheet.CommentsAndNotesCollections = new Dictionary<string, PdfCommentsAndNotes>();
            pdfSheet.Worksheet = range.Worksheet;
            pdfSheet.Ranges[0].Add(new PdfRange(range, true));
            GetMaps(pageSettings, pdfSheet, pdfSheet.Ranges[0]);
        }

        public PdfCatalog(PdfPageSettings pageSettings, ExcelWorksheet worksheet)
        {
            PdfWorksheet pdfSheet = new PdfWorksheet();
            pdfSheet.Ranges = new List<PdfRange>[1];
            if (pageSettings.CommentsAndNotes == CommentsAndNotes.AtEndOfSheet)
                pdfSheet.CommentsAndNotesCollections = new Dictionary<string, PdfCommentsAndNotes>();
            pdfSheet.Worksheet = worksheet;
            pdfSheet.Ranges[0] = GetRanges(pdfSheet.Worksheet);
            GetMaps(pageSettings, pdfSheet, pdfSheet.Ranges[0]);
        }

        public PdfCatalog(PdfPageSettings pageSettings, ExcelWorksheet[] worksheets)
        {
            HandleWorksheetCollection(pageSettings, worksheets);
        }

        public PdfCatalog(PdfPageSettings pageSettings, List<ExcelWorksheet> worksheets)
        {
            HandleWorksheetCollection(pageSettings, worksheets.ToArray());
        }

        public PdfCatalog(PdfPageSettings pageSettings, ExcelWorkbook workbook)
        {
            HandleWorksheetCollection(pageSettings, workbook.Worksheets.ToArray());
        }

        private void HandleWorksheetCollection(PdfPageSettings pageSettings, ExcelWorksheet[] worksheets)
        {
            PdfWorksheet pdfSheet = new PdfWorksheet();
            pdfSheet.Ranges = new List<PdfRange>[worksheets.Length];
            for (int i = 0; i < worksheets.Length; i++)
            {
                if (pageSettings.CommentsAndNotes == CommentsAndNotes.AtEndOfSheet)
                    pdfSheet.CommentsAndNotesCollections =  new Dictionary<string, PdfCommentsAndNotes>();
                pdfSheet.Worksheet = worksheets[i];
                pdfSheet.HeaderFooters = new PdfHeaderFooterCollection(pdfSheet.Worksheet.HeaderFooter);
                pdfSheet.Ranges[i] = GetRanges(pdfSheet.Worksheet);
            }
            foreach (var range in pdfSheet.Ranges)
            {
                GetMaps(pageSettings, pdfSheet, range);
            }
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

        private void GetHeaderFooter(PdfPageSettings pageSettings, PdfWorksheet pdfSheet)
        {

        }

        private void GetMaps(PdfPageSettings pageSettings, PdfWorksheet pdfSheet, List<PdfRange> ranges)
        {
            for (int i = 0; i < ranges.Count; i++)
            {
                var temp = ranges[i];
                temp.Map = PdfTextMap.SetTextMap(pageSettings, Dictionaries, pdfSheet, ranges[i]);
                ranges[i] = temp;
            }
        }

        private void HandleMapCollection()
        {
            //foreach var map in maps
                //shape
                //Layout
                //Pages
                //return transform to ExcelPdf!
        }
    }
}

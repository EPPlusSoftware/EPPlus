using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.PdfCatalog
{
    internal class PdfCatalog
    {
        public PdfCatalog(ExcelRangeBase range) 
        {
            GetMaps(range);
        }

        public PdfCatalog(ExcelWorksheet worksheet)
        {
            var ranges = GetRanges(worksheet);
            GetMaps(ranges);
        }

        public PdfCatalog(ExcelWorksheet[] worksheets)
        {
            HandleWorksheetCollection(worksheets);
        }

        public PdfCatalog(List<ExcelWorksheet> worksheets)
        {
            HandleWorksheetCollection(worksheets.ToArray());
        }

        public PdfCatalog(ExcelWorkbook workbook)
        {
            HandleWorksheetCollection(workbook.Worksheets.ToArray());
        }

        private void HandleWorksheetCollection(ExcelWorksheet[] worksheets)
        {
            List<PdfRange>[] ranges = new List<PdfRange>[worksheets.Length];
            for (int i = 0; i < worksheets.Length; i++)
            {
                var worksheet = worksheets[i];
                ranges[i] = GetRanges(worksheet);
            }
            foreach (var range in ranges)
            {
                GetMaps(range);
            }
        }

        private List<PdfRange> GetRanges(ExcelWorksheet worksheet)
        {
            List<PdfRange> ranges = new List<PdfRange>();
            if (worksheet.Names.ContainsKey("_xlnm.Print_Area"))
            {
                for (int i = 0; i < worksheet.Names["_xlnm.Print_Area"].Addresses.Count; i++)
                {
                    PdfRange range = new PdfRange();
                    range.ExtendColumns = true;
                    range.Range = worksheet.Cells[worksheet.Names["_xlnm.Print_Area"].Addresses[i].Address];
                    ranges.Add(range);
                }
            }
            else
            {
                PdfRange range = new PdfRange();
                range.ExtendColumns = true;
                range.Range = worksheet.Dimension;
                ranges.Add(range);
            }
            return ranges;
        }



        private void GetMaps(List<ExcelRangeBase> ranges)
        {
            foreach (var range in ranges)
            {
                var map = new  PdfTextMap(range);
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

using EPPlus.Export.Pdf.Settings;
using EPPlusTest;
using OfficeOpenXml;
using OfficeOpenXml.Export.PdfExport;

namespace EPPlus.Export.Pdf.Tests
{
    public abstract class PdfTestBase : TestBase
    {
        protected static string _pdfPath = _worksheetPath + "\\PDF\\";

        protected void SaveAsPdf(ExcelWorksheet sheet, string pdfFileName)
        {
            if (!pdfFileName.ToLower().EndsWith(".pdf"))
            {
                pdfFileName += ".pdf";
            }
            var path = Path.Combine(_pdfPath, pdfFileName);
            sheet.SaveAsPdf(path);
        }

        protected void SaveAsPdf(ExcelWorkbook wb, string pdfFileName)
        {
            if (!pdfFileName.ToLower().EndsWith(".pdf"))
            {
                pdfFileName += ".pdf";
            }
            var path = Path.Combine(_pdfPath, pdfFileName);
            wb.SaveAsPdf(path);
        }

        protected void SaveAsPdf(ExcelWorkbook wb, string pdfFileName, params ExcelRangeBase[] ranges)
        {
            if (!pdfFileName.ToLower().EndsWith(".pdf"))
            {
                pdfFileName += ".pdf";
            }
            var path = Path.Combine(_pdfPath, pdfFileName);
            if (ranges.Count() > 1)
                wb.SaveAsPdf(path, ranges);
            else
                ranges[0].SaveAsPdf(path);
        }

        /// <summary>
        /// Exports with a caller-supplied PdfPageSettings instead of one built from the
        /// worksheet's printer settings. Needed for anything ExcelWorksheet.SaveAsPdf has no way
        /// to express - e.g. GsubFeatures/GposFeatures, or a custom font engine/directory -
        /// since that extension method always builds its own PdfPageSettings internally
        /// (GetPdfSettings.GetPdfSettingsFromPrinterSettings) with no way to override it.
        /// Drives PdfCatalog directly: the same internal entry point SaveAsPdf itself calls into.
        /// </summary>
        protected void SaveAsPdf(ExcelWorksheet sheet, string pdfFileName, PdfPageSettings settings)
        {
            if (!pdfFileName.ToLower().EndsWith(".pdf"))
            {
                pdfFileName += ".pdf";
            }
            var path = Path.Combine(_pdfPath, pdfFileName);
            new PdfCatalog(settings, sheet).Save(path);
        }

        /// <summary>
        /// Workbook-level counterpart to <see cref="SaveAsPdf(ExcelWorksheet, string, PdfPageSettings)"/>.
        /// See that overload's remarks for why this bypasses ExcelWorkbook.SaveAsPdf.
        /// </summary>
        protected void SaveAsPdf(ExcelWorkbook wb, string pdfFileName, PdfPageSettings settings)
        {
            if (!pdfFileName.ToLower().EndsWith(".pdf"))
            {
                pdfFileName += ".pdf";
            }
            var path = Path.Combine(_pdfPath, pdfFileName);
            new PdfCatalog(settings, wb).Save(path);
        }

        /// <summary>
        /// Range-collection counterpart to <see cref="SaveAsPdf(ExcelWorksheet, string, PdfPageSettings)"/>.
        /// See that overload's remarks for why this bypasses ExcelRangeBase/ExcelWorkbook.SaveAsPdf.
        /// </summary>
        protected void SaveAsPdf(ExcelWorkbook wb, string pdfFileName, PdfPageSettings settings, params ExcelRangeBase[] ranges)
        {
            if (!pdfFileName.ToLower().EndsWith(".pdf"))
            {
                pdfFileName += ".pdf";
            }
            var path = Path.Combine(_pdfPath, pdfFileName);
            if (ranges.Count() > 1)
                new PdfCatalog(settings, ranges).Save(path);
            else
                new PdfCatalog(settings, ranges[0]).Save(path);
        }
    }
}
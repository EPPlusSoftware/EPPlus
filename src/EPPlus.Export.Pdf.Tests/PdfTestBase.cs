using EPPlusTest;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Export.Pdf.Tests
{
    public abstract class PdfTestBase : TestBase
    {
        protected static string _pdfPath = _worksheetPath + "\\PDF\\";

        protected void SaveAsPdf(ExcelWorksheet sheet, string pdfFileName)
        {
            if(!pdfFileName.ToLower().EndsWith(".pdf"))
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
    }
}

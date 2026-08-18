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
    }
}

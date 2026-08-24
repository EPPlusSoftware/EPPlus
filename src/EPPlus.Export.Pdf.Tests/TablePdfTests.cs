using EPPlusTest;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Export.Pdf.Tests
{
    [TestClass]
    public class TablePdfTests : PdfTestBase
    {
        [TestMethod]
        public void ExportSimpleTableToPdf()
        {
            using var package = OpenTemplatePackage("SimpleTableToPdf.xlsx");
            var ws = package.Workbook.Worksheets["Sheet1"];
            SaveAsPdf(ws, "SimpleTableToPdf1");
        }

        [TestMethod]
        public void ExportSimpleTableWithTotalsRowToPdf()
        {
            using var package = OpenTemplatePackage("SimpleTableToPdf.xlsx");
            var ws = package.Workbook.Worksheets["Sheet2"];
            SaveAsPdf(ws, "SimpleTableToPdf2");
        }

        [TestMethod]
        public void WorkbookToPdfTest1()
        {
            using var package = OpenTemplatePackage("WorkbookToPdfTest1.xlsx");
            var ws = package.Workbook.Worksheets["S2 (Deleted)"];
            SaveAsPdf(ws, "SimpleTableToPdf3");
        }
    }
}

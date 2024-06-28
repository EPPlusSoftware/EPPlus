using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.OleObject;
using OfficeOpenXml.Style;
using System.IO;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;

namespace EPPlusTest.Drawing
{
    [TestClass]
    public class OLETests : TestBase
    {
        [TestMethod]
        public void TestReadEmbeddedObjectBin()
        {
            using var p = OpenTemplatePackage("OLE3.xlsx");
            var ws = p.Workbook.Worksheets[0];

            var ole = ws.Drawings[0] as ExcelOleObject;
            ws.Drawings.AddOleObject("myFile", false);
        }
    }
}
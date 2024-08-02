using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Drawing.OleObject;

namespace EPPlusTest.Drawing
{
    [TestClass]
    public class OLETests : TestBase
    {
        [TestMethod]
        public void TestReadEmbeddedObjectBin()
        {
            using var p = OpenTemplatePackage("OleEmbeddedFilesTest.xlsx");
            var ws = p.Workbook.Worksheets[0];

            var ole = ws.Drawings[0] as ExcelOleObject;
            ws.Drawings.AddOleObject("C:\\Users\\AdrianParnéus\\Downloads\\drukhari.pdf", false);

            SaveAndCleanup(p);
        }

        [TestMethod]
        public void TestReadEmbeddedObjects()
        {
            {
                using var p = OpenTemplatePackage("OleEmbedded_PDF1.xlsx");
                var ole = p.Workbook.Worksheets[0].Drawings[0];
            }
            {
                using var p = OpenTemplatePackage("OleEmbedded_PDF2.xlsx");
                var ole = p.Workbook.Worksheets[0].Drawings[0];
            }
            {
                using var p = OpenTemplatePackage("OleEmbedded_ZIP.xlsx");
                var ole = p.Workbook.Worksheets[0].Drawings[0];
            }
            {
                using var p = OpenTemplatePackage("OleEmbedded_EXE.xlsx");
                var ole = p.Workbook.Worksheets[0].Drawings[0];
            }
            {
                using var p = OpenTemplatePackage("OleEmbedded_MP4.xlsx");
                var ole = p.Workbook.Worksheets[0].Drawings[0];
            }
            {
                using var p = OpenTemplatePackage("OleEmbedded_MP3.xlsx");
                var ole = p.Workbook.Worksheets[0].Drawings[0];
            }
            {
                using var p = OpenTemplatePackage("OleEmbedded_WAV.xlsx");
                var ole = p.Workbook.Worksheets[0].Drawings[0];
            }
            {
                using var p = OpenTemplatePackage("OleEmbedded_TXT.xlsx");
                var ole = p.Workbook.Worksheets[0].Drawings[0];
            }
        }

    }
}
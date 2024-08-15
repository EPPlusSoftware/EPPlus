using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Drawing.OleObject;
using OfficeOpenXml;

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
        }

        [TestMethod]
        public void TestRead2EmbeddedObjectBin()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\MSSample.xlsx");
            var ws = p.Workbook.Worksheets[0];
            var ole = ws.Drawings[0] as ExcelOleObject;
        }

        [TestMethod]
        public void TestWriteEmbeddedObjectBin()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("MyFirstSheet");
            var ole = ws.Drawings.AddOleObject(@"C:\epplusTest\sample.mp3", false, @"C:\epplusTest\epplusobject.png");
            p.SaveAs(@"C:\epplusTest\sample.xlsx");
        }

        [TestMethod]
        public void TestReadEmbeddedObjectBin2()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\sample.xlsx");
            var ws = p.Workbook.Worksheets[0];
            var ole = ws.Drawings[0] as ExcelOleObject;
        }

        [TestMethod]
        public void TestReadEmbeddedObjects()
        {
            //{
            //    using var p = OpenTemplatePackage("OleEmbedded_PDF4.xlsx");
            //    var ole = p.Workbook.Worksheets[0].Drawings[0];
            //}
            //{
            //    using var p = OpenTemplatePackage("OleEmbedded_PDF3.xlsx");
            //    var ole = p.Workbook.Worksheets[0].Drawings[0];
            //}
            //{
            //    using var p = OpenTemplatePackage("OleEmbedded_PDF2.xlsx");
            //    var ole = p.Workbook.Worksheets[0].Drawings[0];
            //}
            //{
            //    using var p = OpenTemplatePackage("OleEmbedded_GraphChart.xlsx");
            //    var ole = p.Workbook.Worksheets[0].Drawings[0];
            //}
            //{
            //    using var p = OpenTemplatePackage("OleEmbedded_OpenDocumentPresent1.xlsx");
            //    var ole = p.Workbook.Worksheets[0].Drawings[0];
            //}
            //{
            //    using var p = OpenTemplatePackage("OleEmbedded_OpenDocumentText1.xlsx");
            //    var ole = p.Workbook.Worksheets[0].Drawings[0];
            //}
            //{
            //    using var p = OpenTemplatePackage("OleEmbedded_OrgChart.xlsx");
            //    var ole = p.Workbook.Worksheets[0].Drawings[0];
            //}
            //{
            //    using var p = OpenTemplatePackage("OleEmbedded_Package1.xlsx");
            //    var ole = p.Workbook.Worksheets[0].Drawings[0];
            //}
            //{
            //    using var p = OpenTemplatePackage("OleEmbedded_Package2.xlsx");
            //    var ole = p.Workbook.Worksheets[0].Drawings[0];
            //}
            //{
            //    using var p = OpenTemplatePackage("OleEmbedded_PaintbrushPic.xlsx");
            //    var ole = p.Workbook.Worksheets[0].Drawings[0];
            //}
            //{
            //    using var p = OpenTemplatePackage("OleEmbedded_PDF.xlsx");
            //    var ole = p.Workbook.Worksheets[0].Drawings[0];
            //}
            //{
            //    using var p = OpenTemplatePackage("OleEmbedded_PDFSSD.xlsx");
            //    var ole = p.Workbook.Worksheets[0].Drawings[0];
            //}
            //{
            //    using var p = OpenTemplatePackage("OleEmbedded_PDFXML.xlsx");
            //    var ole = p.Workbook.Worksheets[0].Drawings[0];
            //}
            //{
            //    using var p = OpenTemplatePackage("OleEmbedded_PowerPoint97-Present.xlsx");
            //    var ole = p.Workbook.Worksheets[0].Drawings[0];
            //}
            //{
            //    using var p = OpenTemplatePackage("OleEmbedded_PowerPoint97-Slide.xlsx");
            //    var ole = p.Workbook.Worksheets[0].Drawings[0];
            //}
            //{
            //    using var p = OpenTemplatePackage("OleEmbedded_PowerPointMacro-Present.xlsx");
            //    var ole = p.Workbook.Worksheets[0].Drawings[0];
            //}
            //{
            //    using var p = OpenTemplatePackage("OleEmbedded_PowerPointMacro-Slide.xlsx");
            //    var ole = p.Workbook.Worksheets[0].Drawings[0];
            //}
            //{
            //    using var p = OpenTemplatePackage("OleEmbedded_PowerPointPresentation.xlsx");
            //    var ole = p.Workbook.Worksheets[0].Drawings[0];
            //}
            //{
            //    using var p = OpenTemplatePackage("OleEmbedded_PowerPointSlide.xlsx");
            //    var ole = p.Workbook.Worksheets[0].Drawings[0];
            //}
            //{
            //    using var p = OpenTemplatePackage("OleEmbedded_Word.xlsx");
            //    var ole = p.Workbook.Worksheets[0].Drawings[0];
            //}
            //{
            //    using var p = OpenTemplatePackage("OleEmbedded_Word97.xlsx");
            //    var ole = p.Workbook.Worksheets[0].Drawings[0];
            //}
            //{
            //    using var p = OpenTemplatePackage("OleEmbedded_WordMacro.xlsx");
            //    var ole = p.Workbook.Worksheets[0].Drawings[0];
            //}
            //{
            //    using var p = OpenTemplatePackage("OleEmbedded_WordPad.xlsx");
            //    var ole = p.Workbook.Worksheets[0].Drawings[0];
            //}
            //{
            //    using var p = OpenTemplatePackage("OleEmbeddedFilesTest.xlsx");
            //    var ole = p.Workbook.Worksheets[0].Drawings[0];
            //}
        }
    }
}
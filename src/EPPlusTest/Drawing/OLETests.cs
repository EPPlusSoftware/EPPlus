using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.Drawing
{
    [TestClass]
    public class OLETests : TestBase
    {
        //WRITE FILES

        [TestMethod]
        public void WriteEmbedded_MP3()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            var ole = ws.Drawings.AddOleObject(@"C:\epplusTest\OleTest\Files\sample.mp3");
            p.SaveAs(@"C:\epplusTest\OleTest\EPPlusEmbedded_MP3.xlsx");
        }
        [TestMethod]
        public void WriteEmbedded_MP4()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            var ole = ws.Drawings.AddOleObject(@"C:\epplusTest\OleTest\Files\sample.mp3");
            p.SaveAs(@"C:\epplusTest\OleTest\EPPlusEmbedded_MP4.xlsx");
        }
        [TestMethod]
        public void WriteEmbedded_ODS()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            var ole = ws.Drawings.AddOleObject(@"C:\epplusTest\OleTest\Files\sample.mp3");
            p.SaveAs(@"C:\epplusTest\OleTest\EPPlusEmbedded_ODS.xlsx");
        }
        [TestMethod]
        public void WriteEmbedded_ODT()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            var ole = ws.Drawings.AddOleObject(@"C:\epplusTest\OleTest\Files\sample.mp3");
            p.SaveAs(@"C:\epplusTest\OleTest\EPPlusEmbedded_ODT.xlsx");
        }
        [TestMethod]
        public void WriteEmbedded_PDF()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            var ole = ws.Drawings.AddOleObject(@"C:\epplusTest\OleTest\Files\aeldari.pdf");
            p.SaveAs(@"C:\epplusTest\OleTest\EPPlusEmbedded_PDF.xlsx");
        }
        [TestMethod]
        public void WriteEmbedded_TXT()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            var ole = ws.Drawings.AddOleObject(@"C:\epplusTest\OleTest\Files\MyTextDocument.txt");
            p.SaveAs(@"C:\epplusTest\OleTest\EPPlusEmbedded_TXT.xlsx");
        }
        [TestMethod]
        public void WriteEmbedded_WAV()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            var ole = ws.Drawings.AddOleObject(@"C:\epplusTest\OleTest\Files\sample.mp3");
            p.SaveAs(@"C:\epplusTest\OleTest\EPPlusEmbedded_WAV.xlsx");
        }
        [TestMethod]
        public void WriteEmbedded_ZIP()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            var ole = ws.Drawings.AddOleObject(@"C:\epplusTest\OleTest\Files\sample.mp3");
            p.SaveAs(@"C:\epplusTest\OleTest\EPPlusEmbedded_ZIP.xlsx");
        }


        //READ EXCEL FILES

        [TestMethod]
        public void ReadExcelEmbedded_MP3()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\Excels\MP3.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0];
        }
        [TestMethod]
        public void ReadExcelEmbedded_MP4()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\Excels\MP4.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0];
        }
        [TestMethod]
        public void ReadExcelEmbedded_ODS()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\Excels\ODS.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0];
        }
        [TestMethod]
        public void ReadExcelEmbedded_ODT()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\Excels\ODT.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0];
        }
        [TestMethod]
        public void ReadExcelEmbedded_PDF()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\Excels\PDF.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0];
        }
        [TestMethod]
        public void ReadExcelEmbedded_TXT()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\Excels\TXT.xlsx");
            //using var p = new ExcelPackage(@"C:\epplusTest\OleTest\EPPlusEmbedded_TXT.xlsx"); 
            var ole = p.Workbook.Worksheets[0].Drawings[0];
           // p.SaveAs(@"c:\temp\ole.xlsx");
        }
        [TestMethod]
        public void ReadExcelEmbedded_WAV()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\Excels\WAV.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0];
        }
        [TestMethod]
        public void ReadExcelEmbedded_ZIP()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\Excels\ZIP.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0];
        }

        //READ EPPLUS FILES

        [TestMethod]
        public void ReadEPPlusEmbedded_MP3()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\EPPlusEmbedded_MP3.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0];
        }
        [TestMethod]
        public void ReadEPPlusEmbedded_MP4()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\EPPlusEmbedded_MP4.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0];
        }
        [TestMethod]
        public void ReadEPPlusEmbedded_ODS()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\EPPlusEmbedded_ODS.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0];
        }
        [TestMethod]
        public void ReadEPPlusEmbedded_ODT()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\EPPlusEmbedded_ODT.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0];
        }
        [TestMethod]
        public void ReadEPPlusEmbedded_PDF()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\EPPlusEmbedded_PDF.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0];
        }
        [TestMethod]
        public void ReadEPPlusEmbedded_TXT()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\EPPlusEmbedded_TXT.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0];
        }
        [TestMethod]
        public void ReadEPPlusEmbedded_WAV()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\EPPlusEmbedded_WAV.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0];
        }
        [TestMethod]
        public void ReadEPPlusEmbedded_ZIP()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\EPPlusEmbedded_\ZIP.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0];
        }
    }
}
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.OleObject;

namespace EPPlusTest.Drawing
{
    [TestClass]
    public class OLETests : TestBase
    {
        // <summary>
        // LINKED FILES
        // </summary>

        [TestMethod]
        public void ReadExcelExternal_MP3()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\Excels\MP3_LINK.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0] as ExcelOleObject;
        }

        [TestMethod]
        public void WriteExternal_MP3()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            var ole = ws.Drawings.AddOleObject(@"C:\epplusTest\OleTest\Files\sample.mp3", true);
            p.SaveAs(@"C:\epplusTest\OleTest\EPPlusExternal_MP3.xlsx");
        }

        [TestMethod]
        public void WriteExternal_ZIP()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            var ole = ws.Drawings.AddOleObject(@"C:\epplusTest\OleTest\Files\Audio-Sample-files-master.zip", true, OleObjectType.Default, true, @"C:\epplusTest\OleTest\EMF\BigMaskTest.bmp");
            p.SaveAs(@"C:\epplusTest\OleTest\EPPlusExternal_ZIP.xlsx");
        }




        // <summary>
        // EMBEDDED FILES
        // </summary>

        [TestMethod]
        public void ReadXlsx()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\EPPlusEmbedded_XLSX.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0] as ExcelOleObject;
        }
        [TestMethod]
        public void ReadDocx()
        {
            using var p = new ExcelPackage();
            var ole = p.Workbook.Worksheets[0].Drawings[0] as ExcelOleObject;
        }
        [TestMethod]
        public void ReadPptx()
        {
            using var p = new ExcelPackage();
            var ole = p.Workbook.Worksheets[0].Drawings[0] as ExcelOleObject;
        }
        [TestMethod]
        public void WriteXlsx()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            var ole = ws.Drawings.AddOleObject(@"C:\epplusTest\OleTest\Files\MySheet.xlsx", false, OleObjectType.DOC);
            p.SaveAs(@"C:\epplusTest\OleTest\EPPlusEmbedded_XLSX.xlsx");
        }
        [TestMethod]
        public void WriteDocx()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            var ole = ws.Drawings.AddOleObject(@"C:\epplusTest\OleTest\Files\MyWord.docx", false, OleObjectType.DOC);
            p.SaveAs(@"C:\epplusTest\OleTest\EPPlusEmbedded_DOCX.xlsx");
        }
        [TestMethod]
        public void WritePptx()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            var ole = ws.Drawings.AddOleObject(@"C:\epplusTest\OleTest\Files\MyPresent.pptx", false, OleObjectType.DOC);
            p.SaveAs(@"C:\epplusTest\OleTest\EPPlusEmbedded_PPTX.xlsx");
        }



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
            var ole = ws.Drawings.AddOleObject(@"C:\epplusTest\OleTest\Files\Bathory - One Rode To Asa Bay -Official Music Video-.mp4");
            p.SaveAs(@"C:\epplusTest\OleTest\EPPlusEmbedded_MP4.xlsx");
        }
        [TestMethod]
        public void WriteEmbedded_ODS()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            var ole = ws.Drawings.AddOleObject(@"C:\epplusTest\OleTest\Files\MySheets.ods");
            p.SaveAs(@"C:\epplusTest\OleTest\EPPlusEmbedded_ODS.xlsx");
        }
        [TestMethod]
        public void WriteEmbedded_ODT()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            var ole = ws.Drawings.AddOleObject(@"C:\epplusTest\OleTest\Files\MyTextDoc.odt");
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
            var ole = ws.Drawings.AddOleObject(@"C:\epplusTest\OleTest\Files\sample.wav");
            p.SaveAs(@"C:\epplusTest\OleTest\EPPlusEmbedded_WAV.xlsx");
        }
        [TestMethod]
        public void WriteEmbedded_ZIP()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            var ole = ws.Drawings.AddOleObject(@"C:\epplusTest\OleTest\Files\Audio-Sample-files-master.zip");
            p.SaveAs(@"C:\epplusTest\OleTest\EPPlusEmbedded_ZIP.xlsx");
        }


        //READ EXCEL FILES

        [TestMethod]
        public void ReadExcelEmbedded_MP3()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\Excels\MP3.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0] as ExcelOleObject;
            ole.ExportOleObjectData(@"C:\epplusTest\OleTest\Results.xlsx");
        }
        [TestMethod]
        public void ReadExcelEmbedded_MP4()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\Excels\MP4.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0] as ExcelOleObject;
            ole.ExportOleObjectData(@"C:\epplusTest\OleTest\Results.xlsx");
        }
        [TestMethod]
        public void ReadExcelEmbedded_ODS()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\Excels\ODS.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0] as ExcelOleObject;
            ole.ExportOleObjectData(@"C:\epplusTest\OleTest\Results.xlsx");
        }
        [TestMethod]
        public void ReadExcelEmbedded_ODT()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\Excels\ODT.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0] as ExcelOleObject;
            ole.ExportOleObjectData(@"C:\epplusTest\OleTest\Results.xlsx");
        }
        [TestMethod]
        public void ReadExcelEmbedded_PDF()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\Excels\PDF.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0] as ExcelOleObject;
            ole.ExportOleObjectData(@"C:\epplusTest\OleTest\Results.xlsx");
        }
        [TestMethod]
        public void ReadExcelEmbedded_TXT()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\Excels\TXT.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0] as ExcelOleObject;
            ole.ExportOleObjectData(@"C:\epplusTest\OleTest\Results.xlsx");
        }
        [TestMethod]
        public void ReadExcelEmbedded_WAV()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\Excels\WAV.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0] as ExcelOleObject;
            ole.ExportOleObjectData(@"C:\epplusTest\OleTest\Results.xlsx");
        }
        [TestMethod]
        public void ReadExcelEmbedded_ZIP()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\Excels\ZIP.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0] as ExcelOleObject;
            ole.ExportOleObjectData(@"C:\epplusTest\OleTest\Results.xlsx");
        }

        //READ EPPLUS FILES

        [TestMethod]
        public void ReadEPPlusEmbedded_MP3()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\EPPlusEmbedded_MP3.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0] as ExcelOleObject;
            ole.ExportOleObjectData(@"C:\epplusTest\OleTest\Results.xlsx");
        }
        [TestMethod]
        public void ReadEPPlusEmbedded_MP4()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\EPPlusEmbedded_MP4.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0] as ExcelOleObject;
            ole.ExportOleObjectData(@"C:\epplusTest\OleTest\Results.xlsx");
        }
        [TestMethod]
        public void ReadEPPlusEmbedded_ODS()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\EPPlusEmbedded_ODS.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0] as ExcelOleObject;
            ole.ExportOleObjectData(@"C:\epplusTest\OleTest\Results.xlsx");
        }
        [TestMethod]
        public void ReadEPPlusEmbedded_ODT()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\EPPlusEmbedded_ODT.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0] as ExcelOleObject;
            ole.ExportOleObjectData(@"C:\epplusTest\OleTest\Results.xlsx");
        }
        [TestMethod]
        public void ReadEPPlusEmbedded_PDF()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\EPPlusEmbedded_PDF.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0] as ExcelOleObject;
            ole.ExportOleObjectData(@"C:\epplusTest\OleTest\Results.xlsx");
        }
        [TestMethod]
        public void ReadEPPlusEmbedded_TXT()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\EPPlusEmbedded_TXT.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0] as ExcelOleObject;
            ole.ExportOleObjectData(@"C:\epplusTest\OleTest\Results.xlsx");
        }
        [TestMethod]
        public void ReadEPPlusEmbedded_WAV()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\EPPlusEmbedded_WAV.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0] as ExcelOleObject;
            ole.ExportOleObjectData(@"C:\epplusTest\OleTest\Results.xlsx");
        }
        [TestMethod]
        public void ReadEPPlusEmbedded_ZIP()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\OleTest\EPPlusEmbedded_\ZIP.xlsx");
            var ole = p.Workbook.Worksheets[0].Drawings[0] as ExcelOleObject;
            ole.ExportOleObjectData(@"C:\epplusTest\OleTest\Results.xlsx");
        }
    }
}
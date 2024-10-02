using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.OleObject;
using OfficeOpenXml.Drawing.OleObject.Structures;

namespace EPPlusTest.Drawing
{
    [TestClass]
    public class OLETests : TestBase
    {
        //Generic OLE Object
        [TestMethod]
        public void ReadEmbeddedOleObject()
        {
            //Read generic ole object.
            var genericOlePackage = OpenTemplatePackage("OleObjectTest_Embed_GENERIC.xlsx");
            var genericOleWs = genericOlePackage.Workbook.Worksheets[0];
            var genericOle = genericOleWs.Drawings[0];
            bool isExcelOleObject = genericOle is ExcelOleObject;
            Assert.IsTrue(isExcelOleObject);
            var embededOle = genericOle as ExcelOleObject;
            Assert.IsFalse(embededOle.IsExternalLink);

            //Read PDF Object
            var pdfOlePackage = OpenTemplatePackage("OleObjectTest_Embed_PDF.xlsx");
            var pdfOleWs = pdfOlePackage.Workbook.Worksheets[0];
            var pdfOle = pdfOleWs.Drawings[0];
            isExcelOleObject = pdfOle is ExcelOleObject;
            Assert.IsTrue(isExcelOleObject);
            embededOle = pdfOle as ExcelOleObject;
            Assert.IsFalse(embededOle.IsExternalLink);

            //Read DOCX Object
            var docxOlePackage = OpenTemplatePackage("OleObjectTest_Embed_DOCX.xlsx");
            var docxOleWs = docxOlePackage.Workbook.Worksheets[0];
            var docxOle = docxOleWs.Drawings[0];
            isExcelOleObject = docxOle is ExcelOleObject;
            Assert.IsTrue(isExcelOleObject);
            embededOle = docxOle as ExcelOleObject;
            Assert.IsFalse(embededOle.IsExternalLink);

            //Read PPTX Object
            var pptxOlePackage = OpenTemplatePackage("OleObjectTest_Embed_PPTX.xlsx");
            var pptxOleWs = pptxOlePackage.Workbook.Worksheets[0];
            var pptxOle = pptxOleWs.Drawings[0];
            isExcelOleObject = pptxOle is ExcelOleObject;
            Assert.IsTrue(isExcelOleObject);
            embededOle = pptxOle as ExcelOleObject;
            Assert.IsFalse(embededOle.IsExternalLink);

            //Read XLSX Object
            var xlsxOlePackage = OpenTemplatePackage("OleObjectTest_Embed_XLSX.xlsx");
            var xlsxOleWs = xlsxOlePackage.Workbook.Worksheets[0];
            var xlsxOle = xlsxOleWs.Drawings[0];
            isExcelOleObject = xlsxOle is ExcelOleObject;
            Assert.IsTrue(isExcelOleObject);
            embededOle = xlsxOle as ExcelOleObject;
            Assert.IsFalse(embededOle.IsExternalLink);

            //Read ODS Object
            var odsOlePackage = OpenTemplatePackage("OleObjectTest_Embed_ODS.xlsx");
            var odsOleWs = odsOlePackage.Workbook.Worksheets[0];
            var odsOle = odsOleWs.Drawings[0];
            isExcelOleObject = odsOle is ExcelOleObject;
            Assert.IsTrue(isExcelOleObject);
            embededOle = odsOle as ExcelOleObject;
            Assert.IsFalse(embededOle.IsExternalLink);
        }

        [TestMethod]
        public void ReadLinkedOleObject()
        {
            //Read generic ole object.
            var genericOlePackage = OpenTemplatePackage("OleObjectTest_Link_GENERIC.xlsx");
            var genericOleWs = genericOlePackage.Workbook.Worksheets[0];
            var genericOle = genericOleWs.Drawings[0];
            bool isExcelOleObject = genericOle is ExcelOleObject;
            Assert.IsTrue(isExcelOleObject);
            var linkedOle = genericOle as ExcelOleObject;
            Assert.IsTrue(linkedOle.IsExternalLink);

            //Read PDF Object
            var pdfOlePackage = OpenTemplatePackage("OleObjectTest_Link_PDF.xlsx");
            var pdfOleWs = pdfOlePackage.Workbook.Worksheets[0];
            var pdfOle = pdfOleWs.Drawings[0];
            isExcelOleObject = pdfOle is ExcelOleObject;
            Assert.IsTrue(isExcelOleObject);
            linkedOle = pdfOle as ExcelOleObject;
            Assert.IsTrue(linkedOle.IsExternalLink);

            //Read DOCX Object
            var docxOlePackage = OpenTemplatePackage("OleObjectTest_Link_DOCX.xlsx");
            var docxOleWs = docxOlePackage.Workbook.Worksheets[0];
            var docxOle = docxOleWs.Drawings[0];
            isExcelOleObject = docxOle is ExcelOleObject;
            Assert.IsTrue(isExcelOleObject);
            linkedOle = docxOle as ExcelOleObject;
            Assert.IsTrue(linkedOle.IsExternalLink);
            
            //Read PPTX Object
            var pptxOlePackage = OpenTemplatePackage("OleObjectTest_Link_PPTX.xlsx");
            var pptxOleWs = pptxOlePackage.Workbook.Worksheets[0];
            var pptxOle = pptxOleWs.Drawings[0];
            isExcelOleObject = pptxOle is ExcelOleObject;
            Assert.IsTrue(isExcelOleObject);
            linkedOle = pptxOle as ExcelOleObject;
            Assert.IsTrue(linkedOle.IsExternalLink);

            //Read XLSX Object
            var xlsxOlePackage = OpenTemplatePackage("OleObjectTest_Link_XLSX.xlsx");
            var xlsxOleWs = xlsxOlePackage.Workbook.Worksheets[0];
            var xlsxOle = xlsxOleWs.Drawings[0];
            isExcelOleObject = xlsxOle is ExcelOleObject;
            Assert.IsTrue(isExcelOleObject);
            linkedOle = xlsxOle as ExcelOleObject;
            Assert.IsTrue(linkedOle.IsExternalLink);

            //Read ODS Object
            var odsOlePackage = OpenTemplatePackage("OleObjectTest_Link_ODS.xlsx");
            var odsOleWs = odsOlePackage.Workbook.Worksheets[0];
            var odsOle = odsOleWs.Drawings[0];
            isExcelOleObject = odsOle is ExcelOleObject;
            Assert.IsTrue(isExcelOleObject);
            linkedOle = odsOle as ExcelOleObject;
            Assert.IsTrue(linkedOle.IsExternalLink);
        }

        [TestMethod]
        public void WriteEmbeddedOleObject()
        {
            //Write Generic Object
            using var genericOlePackage = OpenPackage("EpplusOleObject_Generic.xlsx");
            var generiWs = genericOlePackage.Workbook.Worksheets.Add("Sheet 1");
            var genericOle = generiWs.Drawings.AddOleObject(@"C:\epplusTest\Workbooks\OleObjectFiles\MyTextDocument.txt");
            Assert.IsTrue(genericOle._document.Storage.DataStreams.ContainsKey(Ole10Native.OLE10NATIVE_STREAM_NAME));
            Assert.IsTrue(genericOle._document.Storage.DataStreams.ContainsKey(CompObj.COMPOBJ_STREAM_NAME));
            SaveAndCleanup(genericOlePackage);

            //Write PDF Object
            using var pdfOlePackage = OpenPackage("EpplusOleObject_PDF.xlsx");
            var pdfWs = pdfOlePackage.Workbook.Worksheets.Add("Sheet 1");
            var pdfOle = pdfWs.Drawings.AddOleObject(@"C:\epplusTest\Workbooks\OleObjectFiles\MyPDF.pdf");
            Assert.IsTrue(pdfOle._document.Storage.DataStreams.ContainsKey(Ole.OLE_STREAM_NAME));
            Assert.IsTrue(pdfOle._document.Storage.DataStreams.ContainsKey(CompObj.COMPOBJ_STREAM_NAME));
            Assert.IsTrue(pdfOle._document.Storage.DataStreams.ContainsKey(OleDataFile.CONTENTS_STREAM_NAME));
            SaveAndCleanup(pdfOlePackage);
        }
        [TestMethod]
        public void WriteLinkedOleObject()
        {

        }

        [TestMethod]
        public void CheckCompoundDocument_Generic()
        {
            var p = OpenTemplatePackage("OleObjectTest_Embed_GENERIC.xlsx");
            var ws = p.Workbook.Worksheets[0];
            var ole = ws.Drawings[0] as ExcelOleObject;
            Assert.IsTrue(ole._document.Storage.DataStreams.ContainsKey(Ole10Native.OLE10NATIVE_STREAM_NAME));
            Assert.IsTrue(ole._document.Storage.DataStreams.ContainsKey(CompObj.COMPOBJ_STREAM_NAME));
        }
        [TestMethod]
        public void CheckCompoundDocument_PDF()
        {
            var p = OpenTemplatePackage("OleObjectTest_Embed_PDF.xlsx");
            var ws = p.Workbook.Worksheets[0];
            var ole = ws.Drawings[0] as ExcelOleObject;
            Assert.IsTrue(ole._document.Storage.DataStreams.ContainsKey(OleDataFile.CONTENTS_STREAM_NAME));
            Assert.IsTrue(ole._document.Storage.DataStreams.ContainsKey(CompObj.COMPOBJ_STREAM_NAME));
            Assert.IsTrue(ole._document.Storage.DataStreams.ContainsKey(Ole.OLE_STREAM_NAME));
        }
        [TestMethod]
        public void CheckCompoundDocument_ODS()
        {
            var p = OpenTemplatePackage("OleObjectTest_Embed_ODS.xlsx");
            var ws = p.Workbook.Worksheets[0];
            var ole = ws.Drawings[0] as ExcelOleObject;
            Assert.IsTrue(ole._document.Storage.DataStreams.ContainsKey(OleDataFile.EMBEDDEDODF_STREAM_NAME));
            Assert.IsTrue(ole._document.Storage.DataStreams.ContainsKey(CompObj.COMPOBJ_STREAM_NAME));
        }
        [TestMethod]
        public void CheckMsOff_DOCX()
        {
            var p = OpenTemplatePackage("OleObjectTest_Embed_DOCX.xlsx");
            var ws = p.Workbook.Worksheets[0];
            var ole = ws.Drawings[0] as ExcelOleObject;
            Assert.IsTrue(ole.oleObjectPart.Uri.ToString().Contains(".docx"));
        }
        [TestMethod]
        public void CheckMsOff_PPTX()
        {
            var p = OpenTemplatePackage("OleObjectTest_Embed_PPTX.xlsx");
            var ws = p.Workbook.Worksheets[0];
            var ole = ws.Drawings[0] as ExcelOleObject;
            Assert.IsTrue(ole.oleObjectPart.Uri.ToString().Contains(".pptx"));
        }
        [TestMethod]
        public void CheckMsOff_XLSX()
        {
            var p = OpenTemplatePackage("OleObjectTest_Embed_XLSX.xlsx");
            var ws = p.Workbook.Worksheets[0];
            var ole = ws.Drawings[0] as ExcelOleObject;
            Assert.IsTrue(ole.oleObjectPart.Uri.ToString().Contains(".xlsx"));
        }





        //OLD TESTS FOR INSPIRATION





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
            var ole = ws.Drawings.AddOleObject(@"C:\epplusTest\OleTest\Files\Audio-Sample-files-master.zip", true, true, @"C:\epplusTest\OleTest\EMF\BigMaskTest.bmp");
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
            var ole = ws.Drawings.AddOleObject(@"C:\epplusTest\OleTest\Files\MySheet.xlsx", false, false, ""/*, OleObjectType.DOC*/);
            p.SaveAs(@"C:\epplusTest\OleTest\EPPlusEmbedded_XLSX.xlsx");
        }
        [TestMethod]
        public void WriteDocx()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            var ole = ws.Drawings.AddOleObject(@"C:\epplusTest\OleTest\Files\MyWord.docx", false, false, ""/*, OleObjectType.DOC*/);
            p.SaveAs(@"C:\epplusTest\OleTest\EPPlusEmbedded_DOCX.xlsx");
        }
        [TestMethod]
        public void WritePptx()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            var ole = ws.Drawings.AddOleObject(@"C:\epplusTest\OleTest\Files\MyPresent.pptx", false, false, ""/*, OleObjectType.DOC*/);
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
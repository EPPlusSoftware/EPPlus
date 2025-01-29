using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.OleObject;
using OfficeOpenXml.Drawing.OleObject.Structures;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace EPPlusTest.Drawing
{
    [TestClass]
    public class OLETests : TestBase
    {
        [TestMethod]
        public void ReadEmbeddedOleObject()
        {
            //Read generic ole object.
            var genericOlePackage = OpenTemplatePackage("OleObjectTest_Embed_GENERIC.xlsx");
            var genericOleWs = genericOlePackage.Workbook.Worksheets[0];
            var genericOle = genericOleWs.Drawings["MyTextFile"];
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

            //Read ODP Object
            var odpOlePackage = OpenTemplatePackage("OleObjectTest_Embed_ODP.xlsx");
            var odpOleWs = odpOlePackage.Workbook.Worksheets[0];
            var odpOle = odpOleWs.Drawings[0];
            isExcelOleObject = odpOle is ExcelOleObject;
            Assert.IsTrue(isExcelOleObject);
            embededOle = odpOle as ExcelOleObject;
            Assert.IsFalse(embededOle.IsExternalLink);

            //Read ODT Object
            var odtOlePackage = OpenTemplatePackage("OleObjectTest_Embed_ODT.xlsx");
            var odtOleWs = odtOlePackage.Workbook.Worksheets[0];
            var odtOle = odtOleWs.Drawings[0];
            isExcelOleObject = odtOle is ExcelOleObject;
            Assert.IsTrue(isExcelOleObject);
            embededOle = odtOle as ExcelOleObject;
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

            //Read ODT Object
            var odtOlePackage = OpenTemplatePackage("OleObjectTest_Link_ODT.xlsx");
            var odtOleWs = odtOlePackage.Workbook.Worksheets[0];
            var odtOle = odtOleWs.Drawings[0];
            isExcelOleObject = odtOle is ExcelOleObject;
            Assert.IsTrue(isExcelOleObject);
            linkedOle = odtOle as ExcelOleObject;
            Assert.IsTrue(linkedOle.IsExternalLink);

            //Read ODP Object
            var odpOlePackage = OpenTemplatePackage("OleObjectTest_Link_ODP.xlsx");
            var odpOleWs = odpOlePackage.Workbook.Worksheets[0];
            var odpOle = odpOleWs.Drawings[0];
            isExcelOleObject = odpOle is ExcelOleObject;
            Assert.IsTrue(isExcelOleObject);
            linkedOle = odpOle as ExcelOleObject;
            Assert.IsTrue(linkedOle.IsExternalLink);
        }

        [TestMethod]
        public void WriteEmbeddedOleObject()
        {
            //Write Generic Object
            using var genericOlePackage = OpenPackage("EpplusOleObject_Embed_Generic.xlsx", true);
            var generiWs = genericOlePackage.Workbook.Worksheets.Add("Sheet 1");
            var myFile = Properties.Resources.GetOLEObjectFullFileName("MyTextDocument.txt");
            var genericOle = generiWs.Drawings.AddOleObject("MyTextFile", myFile);
            Assert.IsTrue(genericOle._document.Storage.DataStreams.ContainsKey(Ole10Native.OLE10NATIVE_STREAM_NAME));
            Assert.IsTrue(genericOle._document.Storage.DataStreams.ContainsKey(CompObj.COMPOBJ_STREAM_NAME));
            Assert.IsFalse(genericOle.IsExternalLink);
            SaveAndCleanup(genericOlePackage);

            //Write PDF Object
            using var pdfOlePackage = OpenPackage("EpplusOleObject_Embed_PDF.xlsx", true);
            var pdfWs = pdfOlePackage.Workbook.Worksheets.Add("Sheet 1");
            myFile = Properties.Resources.GetOLEObjectFullFileName("MyPDF.pdf");
            var pdfOle = pdfWs.Drawings.AddOleObject("MyPDFFile", myFile);
            Assert.IsTrue(pdfOle._document.Storage.DataStreams.ContainsKey(Ole.OLE_STREAM_NAME));
            Assert.IsTrue(pdfOle._document.Storage.DataStreams.ContainsKey(CompObj.COMPOBJ_STREAM_NAME));
            Assert.IsTrue(pdfOle._document.Storage.DataStreams.ContainsKey(OleDataFile.CONTENTS_STREAM_NAME));
            Assert.IsFalse(pdfOle.IsExternalLink);
            SaveAndCleanup(pdfOlePackage);

            //Write DOCX Object
            using var docxOlePackage = OpenPackage("EpplusOleObject_Embed_DOCX.xlsx", true);
            var docxWs = docxOlePackage.Workbook.Worksheets.Add("Sheet 1");
            myFile = Properties.Resources.GetOLEObjectFullFileName("MyWord.docx");
            var docxOle = docxWs.Drawings.AddOleObject("MyWordFile", myFile);
            Assert.IsTrue(docxOle._oleObjectPart.Uri.ToString().Contains(".docx"));
            Assert.IsFalse(docxOle.IsExternalLink);
            SaveAndCleanup(docxOlePackage);

            //Write PPTX Object
            using var pptxOlePackage = OpenPackage("EpplusOleObject_Embed_PPTX.xlsx", true);
            var pptxWs = pptxOlePackage.Workbook.Worksheets.Add("Sheet 1");
            myFile = Properties.Resources.GetOLEObjectFullFileName("MyPresent.pptx");
            var pptxOle = pptxWs.Drawings.AddOleObject("MyPPFile", myFile);
            Assert.IsTrue(pptxOle._oleObjectPart.Uri.ToString().Contains(".pptx"));
            Assert.IsFalse(pptxOle.IsExternalLink);
            SaveAndCleanup(pptxOlePackage);

            //Write XLSX Object
            using var xlsxOlePackage = OpenPackage("EpplusOleObject_Embed_XLSX.xlsx", true);
            var xlsxWs = xlsxOlePackage.Workbook.Worksheets.Add("Sheet 1");
            myFile = Properties.Resources.GetOLEObjectFullFileName("MySheet.xlsx");
            var xlsxOle = xlsxWs.Drawings.AddOleObject("MyExcelFile", myFile);
            Assert.IsTrue(xlsxOle._oleObjectPart.Uri.ToString().Contains(".xlsx"));
            Assert.IsFalse(xlsxOle.IsExternalLink);
            SaveAndCleanup(xlsxOlePackage);

            //Write ODS Object
            using var odsOlePackage = OpenPackage("EpplusOleObject_Embed_ODS.xlsx", true);
            var odsWs = odsOlePackage.Workbook.Worksheets.Add("Sheet 1");
            myFile = Properties.Resources.GetOLEObjectFullFileName("MySheets.ods");
            var odsOle = odsWs.Drawings.AddOleObject("MySpreadsheetFile", myFile);
            Assert.IsTrue(odsOle._document.Storage.DataStreams.ContainsKey(Ole.OLE_STREAM_NAME));
            Assert.IsTrue(odsOle._document.Storage.DataStreams.ContainsKey(CompObj.COMPOBJ_STREAM_NAME));
            Assert.IsTrue(odsOle._document.Storage.DataStreams.ContainsKey(OleDataFile.EMBEDDEDODF_STREAM_NAME));
            Assert.IsFalse(odsOle.IsExternalLink);
            SaveAndCleanup(odsOlePackage);

            //Write ODT Object
            using var odtOlePackage = OpenPackage("EpplusOleObject_Embed_ODT.xlsx", true);
            var odtWs = odtOlePackage.Workbook.Worksheets.Add("Sheet 1");
            myFile = Properties.Resources.GetOLEObjectFullFileName("MyTextDoc.odt");
            var odtOle = odtWs.Drawings.AddOleObject("MyDocFile", myFile);
            Assert.IsTrue(odtOle._document.Storage.DataStreams.ContainsKey(Ole.OLE_STREAM_NAME));
            Assert.IsTrue(odtOle._document.Storage.DataStreams.ContainsKey(CompObj.COMPOBJ_STREAM_NAME));
            Assert.IsTrue(odtOle._document.Storage.DataStreams.ContainsKey(OleDataFile.EMBEDDEDODF_STREAM_NAME));
            Assert.IsFalse(odtOle.IsExternalLink);
            SaveAndCleanup(odtOlePackage);

            //Write ODP Object
            using var odpOlePackage = OpenPackage("EpplusOleObject_Embed_ODP.xlsx", true);
            var odpWs = odpOlePackage.Workbook.Worksheets.Add("Sheet 1");
            myFile = Properties.Resources.GetOLEObjectFullFileName("MyPresents.odp");
            var odpOle = odpWs.Drawings.AddOleObject("MyPresentFile", myFile);
            Assert.IsTrue(odpOle._document.Storage.DataStreams.ContainsKey(Ole.OLE_STREAM_NAME));
            Assert.IsTrue(odpOle._document.Storage.DataStreams.ContainsKey(CompObj.COMPOBJ_STREAM_NAME));
            Assert.IsTrue(odpOle._document.Storage.DataStreams.ContainsKey(OleDataFile.EMBEDDEDODF_STREAM_NAME));
            Assert.IsFalse(odpOle.IsExternalLink);
            SaveAndCleanup(odpOlePackage);
        }

        [TestMethod]
        public void WriteLinkedOleObject()
        {
            //Write Generic Object
            using var genericOlePackage = OpenPackage("EpplusOleObject_Link_Generic.xlsx", true);
            var generiWs = genericOlePackage.Workbook.Worksheets.Add("Sheet 1");
            var myFile = Properties.Resources.GetOLEObjectFullFileName("MyTextDocument.txt");
            var genericOle = generiWs.Drawings.AddOleObject("MyTextFile", myFile, o => o.LinkToFile = true);
            Assert.IsNotNull(genericOle._externalLink);
            Assert.IsTrue(genericOle.IsExternalLink);
            SaveAndCleanup(genericOlePackage);

            //Write PDF Object
            using var pdfOlePackage = OpenPackage("EpplusOleObject_Link_PDF.xlsx", true);
            var pdfWs = pdfOlePackage.Workbook.Worksheets.Add("Sheet 1");
            myFile = Properties.Resources.GetOLEObjectFullFileName("MyPDF.pdf");
            var pdfOle = pdfWs.Drawings.AddOleObject("MyPDFFile", myFile, o => o.LinkToFile = true);
            Assert.IsNotNull(pdfOle._externalLink);
            Assert.IsTrue(pdfOle.IsExternalLink);
            SaveAndCleanup(pdfOlePackage);

            //Write DOCX Object
            using var docxOlePackage = OpenPackage("EpplusOleObject_Link_DOCX.xlsx", true);
            var docxWs = docxOlePackage.Workbook.Worksheets.Add("Sheet 1");
            myFile = Properties.Resources.GetOLEObjectFullFileName("MyWord.docx");
            var docxOle = docxWs.Drawings.AddOleObject("MyWordFile", myFile, o => o.LinkToFile = true);
            Assert.IsNotNull(docxOle._externalLink);
            Assert.IsTrue(docxOle.IsExternalLink);
            SaveAndCleanup(docxOlePackage);

            //Write PPTX Object
            using var pptxOlePackage = OpenPackage("EpplusOleObject_Link_PPTX.xlsx", true);
            var pptxWs = pptxOlePackage.Workbook.Worksheets.Add("Sheet 1");
            myFile = Properties.Resources.GetOLEObjectFullFileName("MyPresent.pptx");
            var pptxOle = pptxWs.Drawings.AddOleObject("MyPPFile", myFile, o => o.LinkToFile = true);
            Assert.IsNotNull(pptxOle._externalLink);
            Assert.IsTrue(pptxOle.IsExternalLink);
            SaveAndCleanup(pptxOlePackage);

            //Write XLSX Object
            using var xlsxOlePackage = OpenPackage("EpplusOleObject_Link_XLSX.xlsx", true);
            var xlsxWs = xlsxOlePackage.Workbook.Worksheets.Add("Sheet 1");
            myFile = Properties.Resources.GetOLEObjectFullFileName("MySheet.xlsx");
            var xlsxOle = xlsxWs.Drawings.AddOleObject("MyExcelFile", myFile, o => o.LinkToFile = true);
            Assert.IsNotNull(xlsxOle._externalLink);
            Assert.IsTrue(xlsxOle.IsExternalLink);
            SaveAndCleanup(xlsxOlePackage);

            //Write ODS Object
            using var odsOlePackage = OpenPackage("EpplusOleObject_Link_ODS.xlsx", true);
            var odsWs = odsOlePackage.Workbook.Worksheets.Add("Sheet 1");
            myFile = Properties.Resources.GetOLEObjectFullFileName("MySheets.ods");
            var odsOle = odsWs.Drawings.AddOleObject("MySpreadsheetFile", myFile, o => o.LinkToFile = true);
            Assert.IsNotNull(odsOle._externalLink);
            Assert.IsTrue(odsOle.IsExternalLink);
            SaveAndCleanup(odsOlePackage);

            //Write ODT Object
            using var odtOlePackage = OpenPackage("EpplusOleObject_Link_ODT.xlsx", true);
            var odtWs = odtOlePackage.Workbook.Worksheets.Add("Sheet 1");
            myFile = Properties.Resources.GetOLEObjectFullFileName("MyTextDoc.odt");
            var odtOle = odtWs.Drawings.AddOleObject("MyDocFile", myFile, o => o.LinkToFile = true);
            Assert.IsNotNull(odtOle._externalLink);
            Assert.IsTrue(odtOle.IsExternalLink);
            SaveAndCleanup(odtOlePackage);

            //Write ODP Object
            using var odpOlePackage = OpenPackage("EpplusOleObject_Link_ODP.xlsx", true);
            var odpWs = odpOlePackage.Workbook.Worksheets.Add("Sheet 1");
            myFile = Properties.Resources.GetOLEObjectFullFileName("MyPresents.odp");
            var odpOle = odpWs.Drawings.AddOleObject("MyPresentFile", myFile, o => o.LinkToFile = true);
            Assert.IsNotNull(odpOle._externalLink);
            Assert.IsTrue(odpOle.IsExternalLink);
            SaveAndCleanup(odpOlePackage);
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
            Assert.IsTrue(ole._oleObjectPart.Uri.ToString().Contains(".docx"));
        }
        [TestMethod]
        public void CheckMsOff_PPTX()
        {
            var p = OpenTemplatePackage("OleObjectTest_Embed_PPTX.xlsx");
            var ws = p.Workbook.Worksheets[0];
            var ole = ws.Drawings[0] as ExcelOleObject;
            Assert.IsTrue(ole._oleObjectPart.Uri.ToString().Contains(".pptx"));
        }
        [TestMethod]
        public void CheckMsOff_XLSX()
        {
            var p = OpenTemplatePackage("OleObjectTest_Embed_XLSX.xlsx");
            var ws = p.Workbook.Worksheets[0];
            var ole = ws.Drawings[0] as ExcelOleObject;
            Assert.IsTrue(ole._oleObjectPart.Uri.ToString().Contains(".xlsx"));
        }

        [TestMethod]
        public void DisplayAsIconTest()
        {
            using var genericOlePackage = OpenPackage("EpplusOleObject_Link_Icon_Generic.xlsx", true);
            var myFile = Properties.Resources.GetOLEObjectFullFileName("MyTextDocument.txt");
            var generiWs = genericOlePackage.Workbook.Worksheets.Add("Sheet 1");
            var genericOle = generiWs.Drawings.AddOleObject("MyIcon", myFile, o =>
            {
                o.DisplayAsIcon = true;
                o.LinkToFile = true;
            });
            Assert.IsTrue(genericOle.DisplayAsIcon);
        }
        [TestMethod]
        public void ChangePictureTest()
        {
            using var genericOlePackage = OpenPackage("EpplusOleObject_Link_Icon_Picture_Generic.xlsx", true);
            var myFile = Properties.Resources.GetOLEObjectFullFileName("MyTextDocument.txt");
            var myIcon = Properties.Resources.GetOLEObjectFullFileName("SampleIcon.bmp");
            var generiWs = genericOlePackage.Workbook.Worksheets.Add("Sheet 1");
            var genericOle = generiWs.Drawings.AddOleObject("MyCustomPicture", myFile, o =>
            {
                o.DisplayAsIcon = true;
                o.LinkToFile = true;
                o.Icon = new ExcelImage(myIcon);
            });
            SaveAndCleanup(genericOlePackage);
            //Nothing To Assert just check the excel file and see if it has a different picture.
        }

        [TestMethod]
        public void DeleteEmbeddedOleObjectTest()
        {
            var p = OpenTemplatePackage("OleObjectTest_Embed_DeleteMe.xlsx");
            var ws = p.Workbook.Worksheets[0];
            var ole = ws.Drawings[0] as ExcelOleObject;
            Assert.AreEqual(1, ws.Drawings.Count);
            ws.Drawings.Remove(ole);
            Assert.AreEqual(0, ws.Drawings.Count);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void DeleteEmbeddedOleObjectTest2()
        {
            var p = OpenTemplatePackage("OleObjectTest_Embed_DeleteMe.xlsx");
            var ws = p.Workbook.Worksheets[0];
            Assert.AreEqual(1, ws.Drawings.Count);
            ws.Drawings.Remove("Object 1");
            Assert.AreEqual(0, ws.Drawings.Count);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void DeleteLinkedOleObjectTest()
        {
            var p = OpenTemplatePackage("OleObjectTest_Link_DeleteMe.xlsx");
            var ws = p.Workbook.Worksheets[0];
            var ole = ws.Drawings[0] as ExcelOleObject;
            Assert.AreEqual(1, ws.Drawings.Count);
            ws.Drawings.Remove(ole);
            Assert.AreEqual(0, ws.Drawings.Count);
            SaveAndCleanup(p);
        }

        [TestMethod]
        public void CopyEmbeddedOleObjectTestSameWorksheet()
        {
            var p = OpenTemplatePackage("OleObjectTest_Embed_CopyMe.xlsx");
            var ws = p.Workbook.Worksheets[0];
            var ole = ws.Drawings[0] as ExcelOleObject;
            ws.Drawings[0].Copy(ws, 5, 0);
            Assert.AreEqual(2, ws.Drawings.Count);
            Assert.IsTrue(ws.Drawings[1] is ExcelOleObject);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyEmbeddedOleObjectTestSameWorkbook()
        {
            var p = OpenTemplatePackage("OleObjectTest_Embed_CopyMe.xlsx");
            var ws = p.Workbook.Worksheets[0];
            var ole = ws.Drawings[0] as ExcelOleObject;
            var ws2 = p.Workbook.Worksheets.Add("Sheet 2");
            ws.Drawings[0].Copy(ws2, 5, 0);
            Assert.AreEqual(1, ws2.Drawings.Count);
            Assert.IsTrue(ws2.Drawings[0] is ExcelOleObject);
            SaveAndCleanup(p);
            //p.SaveAs(@"C:\epplusTest\Testoutput\OleObjectTest_Embed_CopyMe2.xlsx");
        }
        [TestMethod]
        public void CopyEmbeddedOleObjectTestOtherWorkbook()
        {
            var p = OpenTemplatePackage("OleObjectTest_Embed_CopyMe.xlsx");
            var ws = p.Workbook.Worksheets[0];
            var ole = ws.Drawings[0] as ExcelOleObject;
            var p2 = new ExcelPackage();
            var ws2 = p2.Workbook.Worksheets.Add("Sheet1");
            ws.Drawings[0].Copy(ws2, 5, 0);
            Assert.AreEqual(1, ws2.Drawings.Count);
            Assert.IsTrue(ws2.Drawings[0] is ExcelOleObject);
            SaveAndCleanup(p2);
            //p2.SaveAs(@"C:\epplusTest\Testoutput\OleObjectTest_Embed_CopyMe3.xlsx");
        }

        [TestMethod]
        public void CopyLinkedOleObjectTestSameWorksheet()
        {
            var p = OpenTemplatePackage("OleObjectTest_Link_CopyMe.xlsx");
            var ws = p.Workbook.Worksheets[0];
            var ole = ws.Drawings[0] as ExcelOleObject;
            ws.Drawings[0].Copy(ws, 5, 0);
            Assert.AreEqual(2, ws.Drawings.Count);
            Assert.IsTrue(ws.Drawings[1] is ExcelOleObject);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyLinkedOleObjectTestSameWorkbook()
        {
            var p = OpenTemplatePackage("OleObjectTest_Link_CopyMe.xlsx");
            var ws = p.Workbook.Worksheets[0];
            var ole = ws.Drawings[0] as ExcelOleObject;
            var ws2 = p.Workbook.Worksheets.Add("Sheet 2");
            ws.Drawings[0].Copy(ws2, 5, 0);
            Assert.AreEqual(1, ws2.Drawings.Count);
            Assert.IsTrue(ws2.Drawings[0] is ExcelOleObject);
            SaveAndCleanup(p);
            //p.SaveAs(@"C:\epplusTest\Testoutput\OleObjectTest_Link_CopyMe2.xlsx");
        }
        [TestMethod]
        public void CopyLinkedOleObjectTestOtherWorkbook()
        {
            var p = OpenTemplatePackage("OleObjectTest_Link_CopyMe.xlsx");
            var ws = p.Workbook.Worksheets[0];
            var ole = ws.Drawings[0] as ExcelOleObject;
            var p2 = new ExcelPackage();
            var ws2 = p2.Workbook.Worksheets.Add("Sheet1");
            ws.Drawings[0].Copy(ws2, 5, 0);
            Assert.AreEqual(1, ws2.Drawings.Count);
            Assert.IsTrue(ws2.Drawings[0] is ExcelOleObject);
            SaveAndCleanup(p2);
            //p2.SaveAs(@"C:\epplusTest\Testoutput\OleObjectTest_Link_CopyMe3.xlsx");
        }

        [TestMethod]
        public void CopyBigWorksheet()
        {
            var p = OpenTemplatePackage("OleObjects.xlsx");
            var ws = p.Workbook.Worksheets[0];
            List<ExcelOleObject> oleObjects = new List<ExcelOleObject>();
            foreach (var ole in ws.Drawings)
            {
                if (ole is ExcelOleObject)
                {
                    oleObjects.Add(ole as ExcelOleObject);
                }
            }
            //Copy to same worksheet
            foreach (var ole in oleObjects)
            {
                ole.Copy(ws, ole.From.Row, ole.From.Column + 10);
            }
            //Copy to new worksheet
            var ws2 = p.Workbook.Worksheets.Add("Copies");
            foreach (var ole in oleObjects)
            {
                ole.Copy(ws2, ole.From.Row, ole.From.Column + 10);
            }
            //Copy to new workbook
            var p1 = new ExcelPackage();
            var ws1 = p1.Workbook.Worksheets.Add("New Workbook");
            foreach (var ole in oleObjects)
            {
                ole.Copy(ws1, ole.From.Row, ole.From.Column + 10);
            }
            //Copy worksheet
            p.Workbook.Worksheets.Add("Worksheet Copy", ws);
            //Save
            SaveAndCleanup(p);
            p1.SaveAs(@"C:\epplusTest\Testoutput\NewOleObjects.xlsx");
        }

        [TestMethod]
        public void CreateLinkOLEFromFileInfo()
        {
            //Write Generic Object
            using var genericOlePackage = OpenPackage("EpplusOleObject_Embed_Generic.xlsx", true);
            var generiWs = genericOlePackage.Workbook.Worksheets.Add("Sheet 1");
            var myFile = Properties.Resources.GetOLEObjectFullFileName("MyTextDocument.txt");
            FileInfo fileInfo = new FileInfo(myFile);
            var genericOle = generiWs.Drawings.AddOleObject("MyTextFile", fileInfo, o => o.LinkToFile = true);
            Assert.IsTrue(genericOle.IsExternalLink);
            SaveAndCleanup(genericOlePackage);
        }
        [TestMethod]
        public void CreateEmbeddedOLEFromFileInfo()
        {
            //Write Generic Object
            using var genericOlePackage = OpenPackage("EpplusOleObject_Embed_Generic.xlsx", true);
            var generiWs = genericOlePackage.Workbook.Worksheets.Add("Sheet 1");
            var myFile = Properties.Resources.GetOLEObjectFullFileName("MyTextDocument.txt");
            FileInfo fileInfo = new FileInfo(myFile);
            var genericOle = generiWs.Drawings.AddOleObject("MyTextFile", fileInfo);
            Assert.IsTrue(genericOle._document.Storage.DataStreams.ContainsKey(Ole10Native.OLE10NATIVE_STREAM_NAME));
            Assert.IsTrue(genericOle._document.Storage.DataStreams.ContainsKey(CompObj.COMPOBJ_STREAM_NAME));
            Assert.IsFalse(genericOle.IsExternalLink);
            SaveAndCleanup(genericOlePackage);
        }
        [TestMethod]
        public void CreateEmbeddedOLEFromStream()
        {
            //Write Generic Object
            using var genericOlePackage = OpenPackage("EpplusOleObject_Embed_Generic.xlsx", true);
            var generiWs = genericOlePackage.Workbook.Worksheets.Add("Sheet 1");
            var myFile = Properties.Resources.GetOLEObjectFullFileName("MyTextDocument.txt");
            using (FileStream fileStream = new FileStream(myFile, FileMode.Open, FileAccess.Read))
            {
                var genericOle = generiWs.Drawings.AddOleObject("MyTextFile", fileStream, "MyTextFile.txt");
                Assert.IsTrue(genericOle._document.Storage.DataStreams.ContainsKey(Ole10Native.OLE10NATIVE_STREAM_NAME));
                Assert.IsTrue(genericOle._document.Storage.DataStreams.ContainsKey(CompObj.COMPOBJ_STREAM_NAME));
                Assert.IsFalse(genericOle.IsExternalLink);
                SaveAndCleanup(genericOlePackage);
            }
        }

        [TestMethod]
        public void SetPosition()
        {
            using (ExcelPackage pck = OpenPackage("OLETest.xlsx", true))
            {
                var wb = pck.Workbook;
                var ws = wb.Worksheets.Add("OleSheet");
                var myIcon = Properties.Resources.GetOLEObjectFullFileName("SampleIcon.bmp");
                var oleImage = ws.Drawings.AddOleObject("ObjectImage", myIcon);
                oleImage.From.Column = 5;
                oleImage.To.Column = 20;
                oleImage.SetSize(175);
                oleImage.UpdateXml();
                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void TestingOLEObjectCopy()
        {
            using (ExcelPackage pck = OpenPackage("OLETest.xlsx", true))
            {
                var wb = pck.Workbook;
                var ws = wb.Worksheets.Add("OleSheet");

                var myIcon = Properties.Resources.GetOLEObjectFullFileName("SampleIcon.bmp");

                var oleImage = ws.Drawings.AddOleObject("ObjectImage", myIcon);

                var ws2 = wb.Worksheets.Add("OleSheetOther");

                oleImage.Copy(ws2, 1, 5);

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void NewAdd()
        {
            //Write Generic Object
            using var genericOlePackage = OpenPackage("EpplusOleObject_Embed_Generic.xlsx", true);
            var generiWs = genericOlePackage.Workbook.Worksheets.Add("Sheet 1");
            var myFile = Properties.Resources.GetOLEObjectFullFileName("MyTextDocument.txt");
            var myIcon = Properties.Resources.GetOLEObjectFullFileName("SampleIcon.bmp");
            FileInfo fileInfo = new FileInfo(myFile);
            var genericOle = generiWs.Drawings.AddOleObject("MyTextFile", fileInfo, new ExcelOleObjectParameters { Icon = new ExcelImage(myIcon) });
            SaveAndCleanup(genericOlePackage);
        }

        [TestMethod]
        public void DeleteTest()
        {
            using var p = OpenPackage("EpplusOleObject_Linked_Deleted1.xlsx", true);
            using var ws = p.Workbook.Worksheets.Add("Sheet 1");
            var myFile = Properties.Resources.GetOLEObjectFullFileName("MyTextDocument.txt");
            var myIcon = Properties.Resources.GetOLEObjectFullFileName("SampleIcon.bmp");
            FileInfo fileInfo1 = new FileInfo(myFile);
            FileInfo fileInfo2 = new FileInfo(myIcon);
            var ole1 = ws.Drawings.AddOleObject("MyTextFile", fileInfo1, o => o.LinkToFile = true);
            var ole2 = ws.Drawings.AddOleObject("MyIconFile", fileInfo2, o => o.LinkToFile = true);
            ws.Drawings.Remove(ole1);
            SaveAndCleanup(p);
        }


        [TestMethod]
        public void DeleteExternalIndexingFixTest()
        {
            using var p = OpenTemplatePackage("ExternalOleLinks.xlsx");
            var ws = p.Workbook.Worksheets[0];
            ws.Drawings.Remove(1);
            SaveAndCleanup(p);
        }

        [TestMethod]
        public void OleAbsoluteAnchor()
        {
            using (ExcelPackage pck = OpenPackage("OLEAbsoluteAnchor.xlsx", true))
            {
                var wb = pck.Workbook;
                var ws = wb.Worksheets.Add("AnchorWs");

                var oleObj = ws.Drawings.AddOleObject("SomeObject", Properties.Resources.GetOLEObjectFullFileName("SampleIcon.bmp"));
                oleObj.ChangeCellAnchor(eEditAs.Absolute);

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void OleSetPosition()
        {
            using (ExcelPackage pck = OpenPackage("OLESetPosition.xlsx", true))
            {
                var wb = pck.Workbook;
                var ws = wb.Worksheets.Add("AnchorWs");
                var wsImage = wb.Worksheets.Add("ImageWs");

                var oleObj = ws.Drawings.AddOleObject("SomeObject", Properties.Resources.GetOLEObjectFullFileName("SampleIcon.bmp"));
                var pic = wsImage.Drawings.AddPicture("SomePicture", GetResourceFile("EPPlus.png").FullName);

                pic.SetPosition(100, 100);
                oleObj.SetPosition(100, 100);

                //See resulting file for the differences in position and scaling
                SaveAndCleanup(pck);
            }
        }


        [TestMethod]
        public void OleAbsoluteAnchorAndSetPosition()
        {
            using (ExcelPackage pck = OpenPackage("OleAbsoluteAnchorAndSetPosition.xlsx", true))
            {
                var wb = pck.Workbook;
                var ws = wb.Worksheets.Add("AnchorWs");
                var wsImage = wb.Worksheets.Add("ImageWs");

                var oleObj = ws.Drawings.AddOleObject("SomeObject", Properties.Resources.GetOLEObjectFullFileName("SampleIcon.bmp"));
                var pic = wsImage.Drawings.AddPicture("SomePicture", GetResourceFile("EPPlus.png").FullName);

                pic.ChangeCellAnchor(eEditAs.Absolute);
                //Changing cell-anchor to absolute does not actually change the type of anchor within worksheet
                oleObj.ChangeCellAnchor(eEditAs.Absolute);

                pic.SetPosition(100, 100);
                oleObj.SetPosition(100, 100);

                oleObj.UpdateXml();


                //For Ole object From and Two appear to remain both in objects created by Excel and Epplus
                //Seems we may need to update From and Two when SetPosition on OLE objects. Yet Picture doesn't have to...
                Assert.AreEqual(pic.Position.X, oleObj.Position.X);
                Assert.AreEqual(pic.Position.Y, oleObj.Position.Y);

                Assert.AreEqual(pic.From.Row, oleObj.From.Row);
                Assert.AreEqual(pic.From.Column, oleObj.From.Column);

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void RemoveMultipleObjectsWithSameFile()
        {
            using (var p = OpenPackage("EpplusOleObject_Linked_AndDeleteAll.xlsx", true))
            {
                using var ws = p.Workbook.Worksheets.Add("Sheet 1");
                var myFile = Properties.Resources.GetOLEObjectFullFileName("MyTextDocument.txt");
                var myIcon = Properties.Resources.GetOLEObjectFullFileName("SampleIcon.bmp");
                FileInfo fileInfo1 = new FileInfo(myFile);
                FileInfo fileInfo2 = new FileInfo(myIcon);
                var ole1 = ws.Drawings.AddOleObject("MyTextFile", fileInfo1, o => o.LinkToFile = true);
                var ole2 = ws.Drawings.AddOleObject("MyIconFile", fileInfo2, o => o.LinkToFile = true);
                var ole3 = ws.Drawings.AddOleObject("MyIconFile2", fileInfo2, o => o.LinkToFile = true);
                var ole4 = ws.Drawings.AddOleObject("MyIconFile3", fileInfo2, o => o.LinkToFile = true);

                ws.Drawings.Remove(ole2);
                ws.Drawings.Remove(ole3);
                ws.Drawings.Remove(ole4);

                SaveAndCleanup(p);
            }
        }
    }
}
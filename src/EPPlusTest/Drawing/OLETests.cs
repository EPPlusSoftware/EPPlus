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
            var genericOle = generiWs.Drawings.AddOleObject(@"C:\epplusTest\Workbooks\OleObjectFiles\MyTextDocument.txt");
            Assert.IsTrue(genericOle._document.Storage.DataStreams.ContainsKey(Ole10Native.OLE10NATIVE_STREAM_NAME));
            Assert.IsTrue(genericOle._document.Storage.DataStreams.ContainsKey(CompObj.COMPOBJ_STREAM_NAME));
            Assert.IsFalse(genericOle.IsExternalLink);
            SaveAndCleanup(genericOlePackage);

            //Write PDF Object
            using var pdfOlePackage = OpenPackage("EpplusOleObject_Embed_PDF.xlsx", true);
            var pdfWs = pdfOlePackage.Workbook.Worksheets.Add("Sheet 1");
            var pdfOle = pdfWs.Drawings.AddOleObject(@"C:\epplusTest\Workbooks\OleObjectFiles\MyPDF.pdf");
            Assert.IsTrue(pdfOle._document.Storage.DataStreams.ContainsKey(Ole.OLE_STREAM_NAME));
            Assert.IsTrue(pdfOle._document.Storage.DataStreams.ContainsKey(CompObj.COMPOBJ_STREAM_NAME));
            Assert.IsTrue(pdfOle._document.Storage.DataStreams.ContainsKey(OleDataFile.CONTENTS_STREAM_NAME));
            Assert.IsFalse(pdfOle.IsExternalLink);
            SaveAndCleanup(pdfOlePackage);

            //Write DOCX Object
            using var docxOlePackage = OpenPackage("EpplusOleObject_Embed_DOCX.xlsx", true);
            var docxWs = docxOlePackage.Workbook.Worksheets.Add("Sheet 1");
            var docxOle = docxWs.Drawings.AddOleObject(@"C:\epplusTest\Workbooks\OleObjectFiles\MyWord.docx");
            Assert.IsTrue(docxOle.oleObjectPart.Uri.ToString().Contains(".docx"));
            Assert.IsFalse(docxOle.IsExternalLink);
            SaveAndCleanup(docxOlePackage);

            //Write PPTX Object
            using var pptxOlePackage = OpenPackage("EpplusOleObject_Embed_PPTX.xlsx", true);
            var pptxWs = pptxOlePackage.Workbook.Worksheets.Add("Sheet 1");
            var pptxOle = pptxWs.Drawings.AddOleObject(@"C:\epplusTest\Workbooks\OleObjectFiles\MyPresent.pptx");
            Assert.IsTrue(pptxOle.oleObjectPart.Uri.ToString().Contains(".pptx"));
            Assert.IsFalse(pptxOle.IsExternalLink);
            SaveAndCleanup(pptxOlePackage);

            //Write XLSX Object
            using var xlsxOlePackage = OpenPackage("EpplusOleObject_Embed_XLSX.xlsx", true);
            var xlsxWs = xlsxOlePackage.Workbook.Worksheets.Add("Sheet 1");
            var xlsxOle = xlsxWs.Drawings.AddOleObject(@"C:\epplusTest\Workbooks\OleObjectFiles\MySheet.xlsx");
            Assert.IsTrue(xlsxOle.oleObjectPart.Uri.ToString().Contains(".xlsx"));
            Assert.IsFalse(xlsxOle.IsExternalLink);
            SaveAndCleanup(xlsxOlePackage);

            //Write ODS Object
            using var odsOlePackage = OpenPackage("EpplusOleObject_Embed_ODS.xlsx", true);
            var odsWs = odsOlePackage.Workbook.Worksheets.Add("Sheet 1");
            var odsOle = odsWs.Drawings.AddOleObject(@"C:\epplusTest\Workbooks\OleObjectFiles\MySheets.ods");
            Assert.IsTrue(odsOle._document.Storage.DataStreams.ContainsKey(Ole.OLE_STREAM_NAME));
            Assert.IsTrue(odsOle._document.Storage.DataStreams.ContainsKey(CompObj.COMPOBJ_STREAM_NAME));
            Assert.IsTrue(odsOle._document.Storage.DataStreams.ContainsKey(OleDataFile.EMBEDDEDODF_STREAM_NAME));
            Assert.IsFalse(odsOle.IsExternalLink);
            SaveAndCleanup(odsOlePackage);

            //Write ODT Object
            using var odtOlePackage = OpenPackage("EpplusOleObject_Embed_ODT.xlsx", true);
            var odtWs = odtOlePackage.Workbook.Worksheets.Add("Sheet 1");
            var odtOle = odtWs.Drawings.AddOleObject(@"C:\epplusTest\Workbooks\OleObjectFiles\MyTextDoc.odt");
            Assert.IsTrue(odtOle._document.Storage.DataStreams.ContainsKey(Ole.OLE_STREAM_NAME));
            Assert.IsTrue(odtOle._document.Storage.DataStreams.ContainsKey(CompObj.COMPOBJ_STREAM_NAME));
            Assert.IsTrue(odtOle._document.Storage.DataStreams.ContainsKey(OleDataFile.EMBEDDEDODF_STREAM_NAME));
            Assert.IsFalse(odtOle.IsExternalLink);
            SaveAndCleanup(odtOlePackage);

            //Write ODP Object
            using var odpOlePackage = OpenPackage("EpplusOleObject_Embed_ODP.xlsx", true);
            var odpWs = odpOlePackage.Workbook.Worksheets.Add("Sheet 1");
            var odpOle = odpWs.Drawings.AddOleObject(@"C:\epplusTest\Workbooks\OleObjectFiles\MyPresents.odp");
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
            var genericOle = generiWs.Drawings.AddOleObject(@"C:\epplusTest\Workbooks\OleObjectFiles\MyTextDocument.txt", true);
            Assert.IsNotNull(genericOle._externalLink);
            Assert.IsTrue(genericOle.IsExternalLink);
            SaveAndCleanup(genericOlePackage);

            //Write PDF Object
            using var pdfOlePackage = OpenPackage("EpplusOleObject_Link_PDF.xlsx", true);
            var pdfWs = pdfOlePackage.Workbook.Worksheets.Add("Sheet 1");
            var pdfOle = pdfWs.Drawings.AddOleObject(@"C:\epplusTest\Workbooks\OleObjectFiles\MyPDF.pdf", true);
            Assert.IsNotNull(pdfOle._externalLink);
            Assert.IsTrue(pdfOle.IsExternalLink);
            SaveAndCleanup(pdfOlePackage);

            //Write DOCX Object
            using var docxOlePackage = OpenPackage("EpplusOleObject_Link_DOCX.xlsx", true);
            var docxWs = docxOlePackage.Workbook.Worksheets.Add("Sheet 1");
            var docxOle = docxWs.Drawings.AddOleObject(@"C:\epplusTest\Workbooks\OleObjectFiles\MyWord.docx", true);
            Assert.IsNotNull(docxOle._externalLink);
            Assert.IsTrue(docxOle.IsExternalLink);
            SaveAndCleanup(docxOlePackage);

            //Write PPTX Object
            using var pptxOlePackage = OpenPackage("EpplusOleObject_Link_PPTX.xlsx", true);
            var pptxWs = pptxOlePackage.Workbook.Worksheets.Add("Sheet 1");
            var pptxOle = pptxWs.Drawings.AddOleObject(@"C:\epplusTest\Workbooks\OleObjectFiles\MyPresent.pptx", true);
            Assert.IsNotNull(pptxOle._externalLink);
            Assert.IsTrue(pptxOle.IsExternalLink);
            SaveAndCleanup(pptxOlePackage);

            //Write XLSX Object
            using var xlsxOlePackage = OpenPackage("EpplusOleObject_Link_XLSX.xlsx", true);
            var xlsxWs = xlsxOlePackage.Workbook.Worksheets.Add("Sheet 1");
            var xlsxOle = xlsxWs.Drawings.AddOleObject(@"C:\epplusTest\Workbooks\OleObjectFiles\MySheet.xlsx", true);
            Assert.IsNotNull(xlsxOle._externalLink);
            Assert.IsTrue(xlsxOle.IsExternalLink);
            SaveAndCleanup(xlsxOlePackage);

            //Write ODS Object
            using var odsOlePackage = OpenPackage("EpplusOleObject_Link_ODS.xlsx", true);
            var odsWs = odsOlePackage.Workbook.Worksheets.Add("Sheet 1");
            var odsOle = odsWs.Drawings.AddOleObject(@"C:\epplusTest\Workbooks\OleObjectFiles\MySheets.ods", true);
            Assert.IsNotNull(odsOle._externalLink);
            Assert.IsTrue(odsOle.IsExternalLink);
            SaveAndCleanup(odsOlePackage);

            //Write ODT Object
            using var odtOlePackage = OpenPackage("EpplusOleObject_Link_ODT.xlsx", true);
            var odtWs = odtOlePackage.Workbook.Worksheets.Add("Sheet 1");
            var odtOle = odtWs.Drawings.AddOleObject(@"C:\epplusTest\Workbooks\OleObjectFiles\MyTextDoc.odt", true);
            Assert.IsNotNull(odtOle._externalLink);
            Assert.IsTrue(odtOle.IsExternalLink);
            SaveAndCleanup(odtOlePackage);

            //Write ODS Object
            using var odpOlePackage = OpenPackage("EpplusOleObject_Link_ODP.xlsx", true);
            var odpWs = odpOlePackage.Workbook.Worksheets.Add("Sheet 1");
            var odpOle = odpWs.Drawings.AddOleObject(@"C:\epplusTest\Workbooks\OleObjectFiles\MyPresents.odp", true);
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

        [TestMethod]
        public void DisplayAsIconTest()
        {
            using var genericOlePackage = OpenPackage("EpplusOleObject_Link_Icon_Generic.xlsx", true);
            var generiWs = genericOlePackage.Workbook.Worksheets.Add("Sheet 1");
            var genericOle = generiWs.Drawings.AddOleObject(@"C:\epplusTest\Workbooks\OleObjectFiles\MyTextDocument.txt", true, true);
            Assert.IsTrue(genericOle.DisplayAsIcon);
        }
        [TestMethod]
        public void ChangePictureTest()
        {
            using var genericOlePackage = OpenPackage("EpplusOleObject_Link_Icon_Picture_Generic.xlsx", true);
            var generiWs = genericOlePackage.Workbook.Worksheets.Add("Sheet 1");
            var genericOle = generiWs.Drawings.AddOleObject(@"C:\epplusTest\Workbooks\OleObjectFiles\MyTextDocument.txt", true, true, @"C:\epplusTest\Workbooks\OleObjectFiles\TestIcon.bmp");
            //Nothing To Assert just check the excel file and see if it has a different picture.
        }
    }
}
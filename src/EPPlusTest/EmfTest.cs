using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.EMF;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using System;
using System.ComponentModel.DataAnnotations;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlusTest
{
    [TestClass]
    public class EmfTest : TestBase
    {
        [TestMethod]
        public void ReadWriteTest()
        {
            using (var package = OpenPackage("ReadEmf.xlsx", true))
            {
                var wb = package.Workbook;
                var ws = wb.Worksheets.Add("EmfSheet");

                //var path = "C:\\Users\\OssianEdström\\Downloads\\OG_image1.emf";
                var path = "C:\\epplusTest\\Workbooks\\UnsignedWithDescriptorsOrigBackup.emf";


                var emf = new EmfImage();
                emf.Read(path);

                var record = (EMR_EXTTEXTOUTW)emf.records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW).First();
                emf.Save("C:\\epplusTest\\Workbooks\\Generated.emf");
            }
        }

        [TestMethod]
        public void ReadWritePreviouslyGeneratedFile()
        {
            using (var package = OpenPackage("ReadEmf.xlsx", true))
            {
                var wb = package.Workbook;
                var ws = wb.Worksheets.Add("EmfSheet");

                var path = "C:\\epplusTest\\Workbooks\\Generated.emf";

                var emf = new EmfImage();
                emf.Read(path);

                emf.Save("C:\\epplusTest\\Workbooks\\GeneratedTwo.emf");
            }
        }

        [TestMethod]
        public void ReadEmfAlt()
        {
            var emfImage = new EmfImage();
            emfImage.Read("C:\\epplusTest\\Testoutput\\LongName.emf");

            var textRecordArr = emfImage.records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW).Skip(2);
            var arr = textRecordArr.ToArray();

            var longName = ((EMR_EXTTEXTOUTW)arr[0]);
            var suggestedSigner = ((EMR_EXTTEXTOUTW)arr[1]);

            var fontRecordArr = emfImage.records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTCREATEFONTINDIRECTW);

            var longIndex = emfImage.records.IndexOf(longName);
            var signerIndex = emfImage.records.IndexOf(suggestedSigner);

            //emfImage.records[140].data = new byte[] { 3, 0, 0, 0 };

            emfImage.Save("C:\\epplusTest\\Testoutput\\ChangeFontOutput.emf");
        }

        [TestMethod]
        public void ReadStampExcel()
        {
            var readExcelVersion = new EmfImage();
            readExcelVersion.Read("C:\\epplusTest\\Testoutput\\TemplateBmp.emf");
            readExcelVersion.Save("C:\\epplusTest\\Testoutput\\TemplateResaved.emf");
        }


        [TestMethod]
        public void ReadStamp2()
        {
            var emfImage = new EmfImage();
            emfImage.Read("C:\\epplusTest\\Testoutput\\ValidStamp.emf");

            var dibits = (EMR_STRETCHDIBITS)emfImage.records.Find(x => x.Type == RECORD_TYPES.EMR_STRETCHDIBITS);

            var fileBytes = File.ReadAllBytes("C:\\Users\\OssianEdström\\Pictures\\ResizedAsExcel.bmp");

            dibits.ChangeImage(fileBytes);

            emfImage.Save("C:\\epplusTest\\Testoutput\\ValidStampChangedImage.emf");
        }

        [TestMethod]
        public void ReadStamp()
        {
            var emfImage = new EmfImage();
            emfImage.Read("C:\\epplusTest\\Testoutput\\ValidStamp.emf");

            var readExcelVersion = new EmfImage();
            readExcelVersion.Read("C:\\epplusTest\\Testoutput\\TemplateBmp.emf");

            var templateRecords = readExcelVersion.records;
            var dibitsTemplate = (EMR_STRETCHDIBITS)readExcelVersion.records.Find(x => x.Type == RECORD_TYPES.EMR_STRETCHDIBITS);

            var records = emfImage.records;

            var dibits = (EMR_STRETCHDIBITS)emfImage.records.Find(x => x.Type == RECORD_TYPES.EMR_STRETCHDIBITS);

            var setWorld = emfImage.records.Find(x => x.Type == RECORD_TYPES.EMR_SETWORLDTRANSFORM);
            var modifyWorld = emfImage.records.Find(x => x.Type == RECORD_TYPES.EMR_MODIFYWORLDTRANSFORM);
            var brushOrgEx = emfImage.records.Find(x => x.Type == RECORD_TYPES.EMR_SETBRUSHORGEX);

            var setWorldTemplate = templateRecords.Find(x => x.Type == RECORD_TYPES.EMR_SETWORLDTRANSFORM);
            var modifyWorldTemplate = templateRecords.Find(x => x.Type == RECORD_TYPES.EMR_MODIFYWORLDTRANSFORM);

            setWorld.data = setWorldTemplate.data;
            modifyWorld.data = modifyWorldTemplate.data;

            brushOrgEx.data = new byte[] { 9, 0, 0, 0, 59, 0, 0, 0 };

            var intersectR = new EMR_INTERSECTCLIPRECT();
            intersectR.Clip = new RectLObject(0, 0, 128, 160);
            records.Insert(154, intersectR);

            var fileBytes = File.ReadAllBytes("C:\\Users\\OssianEdström\\Pictures\\ResizedAsExcel.bmp");

            var handler = new BitmapHandler(fileBytes);

            dibits.bitMapHeader = handler.informationHeader;
            dibits.cbBmiSrc = dibits.bitMapHeader.sizeOfHeader;
            dibits.Padding2 = handler.OptionalData;
            dibits.BitsSrc = handler.PixelArray;

            dibits.cxDest = 128;
            dibits.cxSrc = 128;
            dibits.cySrc = 53;
            dibits.cyDest = 53;

            dibits.Bounds = new RectLObject(9, 59, 117, 102);

            emfImage.Save("C:\\epplusTest\\Testoutput\\ValidStampAltered.emf");
        }

        [TestMethod]
        public void CheckOGImage()
        {
            var emfImage = new EmfImage();
            emfImage.Read("C:\\epplusTest\\Testoutput\\OG_image1.emf");

            var textRecordArr = emfImage.records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW);
            var fontRecordArr = emfImage.records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTCREATEFONTINDIRECTW);
        }

        [TestMethod]
        public void CheckValidTemplate()
        {
            var validTemplate = new SignatureLineTemplateValid();
            var records = validTemplate.records;

            validTemplate.timeStamp.Text = "TimeStamp";
            validTemplate.signTextObject.Text = "TemplateSignature";
            validTemplate.suggestedSignerObject.Text = "TemplateSigner";
            validTemplate.suggestedTitleObject.Text = "TemplateTitle";
            validTemplate.SignedBy = "TemplateName";

            validTemplate.Save("C:\\epplusTest\\Testoutput\\ValidSignatureTemplate2.emf");
        }
        [TestMethod]
        public void CheckInvalidTemplate()
        {
            var invalidTemplate = new SignatureLineTemplateInvalid();
            var records = invalidTemplate.records;

            invalidTemplate.signTextObject.Text = "TemplateSignature";
            invalidTemplate.suggestedSignerObject.Text = "TemplateSigner";
            invalidTemplate.suggestedTitleObject.Text = "TemplateTitle";
            invalidTemplate.SignedBy = "TemplateName";

            invalidTemplate.Save("C:\\epplusTest\\Testoutput\\InvalidSignatureTemplate2.emf");
        }
    }
}

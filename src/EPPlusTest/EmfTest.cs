using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Drawing.EMF;
using System;
using System.Linq;
using OfficeOpenXml.Utils;

namespace EPPlusTest
{
    [TestClass]
    public class EmfTest : TestBase
    {
        const string emfOutputFolder = "EmfOutputFolder";

        [TestMethod]
        public void ReadWriteTest()
        {
            using (var package = OpenPackage("ReadEmf.xlsx", true))
            {
                var wb = package.Workbook;
                var ws = wb.Worksheets.Add("EmfSheet");

                var path = GetTemplateFile("UnsignedWithDescriptorsOrigBackup.emf").FullName;

                var emf = new EmfImage();
                emf.Read(path);

                var record = (EMR_EXTTEXTOUTW)emf.records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW).First();

                var outputPath = GetOutputFile(emfOutputFolder, "Generated.emf").FullName;
                emf.Save(outputPath);
            }
        }

        [TestMethod]
        public void ReadWritePreviouslyGeneratedFile()
        {
            using (var package = OpenPackage("ReadEmf.xlsx", true))
            {
                var wb = package.Workbook;
                var ws = wb.Worksheets.Add("EmfSheet");

                var path = GetTemplateFile("Generated.emf").FullName;

                var emf = new EmfImage();
                emf.Read(path);


                var outputPath = GetOutputFile(emfOutputFolder, "GeneratedTwo.emf").FullName;
                emf.Save(outputPath);
            }
        }

        [TestMethod]
        public void ReadEmfAlt()
        {
            var emfImage = new EmfImage();

            var path = GetTemplateFile("LongName.emf").FullName;

            emfImage.Read(path);

            var textRecordArr = emfImage.records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW).Skip(2);
            var arr = textRecordArr.ToArray();

            var longName = ((EMR_EXTTEXTOUTW)arr[0]);
            var suggestedSigner = ((EMR_EXTTEXTOUTW)arr[1]);

            var fontRecordArr = emfImage.records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTCREATEFONTINDIRECTW);

            var longIndex = emfImage.records.IndexOf(longName);
            var signerIndex = emfImage.records.IndexOf(suggestedSigner);

            var outputPath = GetOutputFile(emfOutputFolder, "ChangeFontOutput.emf").FullName;

            emfImage.Save(outputPath);
        }

        [TestMethod]
        public void CheckOGImage()
        {
            var emfImage = new EmfImage();
            var path = GetTemplateFile("OG_image1.emf").FullName;

            emfImage.Read(path);

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

            var outputPath = GetOutputFile(emfOutputFolder, "ValidSignatureTemplate2.emf").FullName;

            validTemplate.Save(outputPath);
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

            var outputPath = GetOutputFile(emfOutputFolder, "InvalidSignatureTemplate2.emf").FullName;

            invalidTemplate.Save(outputPath);
        }
        [TestMethod]
        public void CompareTemplates()
        {
            var validTemplate = new SignatureLineTemplateValid();
            var records = validTemplate.records;

            var invalidTemplate = new SignatureLineTemplateInvalid();
            var invalidRecords = invalidTemplate.records;

            var saveRecord = new EMR_RECORD();
            saveRecord.Type = RECORD_TYPES.EMR_SAVEDC;
            saveRecord.Size = 8;
            saveRecord.data = new byte[0];

            var startRecordIndex = 63;

            invalidRecords.Insert(startRecordIndex - 1, saveRecord);

            invalidRecords.Insert(startRecordIndex, records[55]);
            invalidRecords.Insert(startRecordIndex + 1, records[56]);
            invalidRecords.Insert(startRecordIndex + 2, records[57]);
            invalidRecords.Insert(startRecordIndex + 3, records[58]);
            invalidRecords.Insert(startRecordIndex + 4, records[59]);
            invalidRecords.Insert(startRecordIndex + 5, records[60]);


            var restoreRecord = new EMR_RECORD();
            restoreRecord.Type = RECORD_TYPES.EMR_RESTOREDC;
            restoreRecord.Size = 12;
            int stackState = -1;
            restoreRecord.data = BitConverter.GetBytes(stackState);

            invalidRecords.Insert(startRecordIndex + 6, restoreRecord);

            var outputPath = GetOutputFile(emfOutputFolder, "InvalidSignatureTemplateOther.emf").FullName;
            invalidTemplate.Save(outputPath);
        }
        [TestMethod]
        public void ChangeImageTemplateForStamp()
        {
            var templateEmf = new EmfImage();
            var resourceFile = GetResourceFile("TemplateForStamp.emf");

            var path = resourceFile.FullName;
            templateEmf.Read(path);

            var templateImage = (EMR_STRETCHDIBITS)templateEmf.records.Find(x => x.Type == RECORD_TYPES.EMR_STRETCHDIBITS);

            templateImage.UpdateToImage(GetResourceFile("111By75.bmp").FullName);
            templateEmf.Save(GetOutputFile(emfOutputFolder, "111By75.emf").FullName);

            templateImage.UpdateToImage(GetResourceFile("MaxWidthSignatureStamp.bmp").FullName);
            templateEmf.Save(GetOutputFile(emfOutputFolder, "MaxWidthSignatureStamp.emf").FullName);

            templateImage.UpdateToImage(GetResourceFile("MaxHeightSignatureStamp.bmp").FullName);
            templateEmf.Save(GetOutputFile(emfOutputFolder, "MaxHeightSignatureStamp.emf").FullName);

            templateImage.UpdateToImage(GetResourceFile("5pxSignature.bmp").FullName);
            templateEmf.Save(GetOutputFile(emfOutputFolder, "5pxSignature.emf").FullName);
        }
    }
}

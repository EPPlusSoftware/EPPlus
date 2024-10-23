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
        public void ReadStamp()
        {
            var emfImage = new EmfImage();
            emfImage.Read("C:\\epplusTest\\Testoutput\\ValidStamp.emf");

            var records = emfImage.records;

            var dibits = (EMR_STRETCHDIBITS)emfImage.records.Find(x => x.Type == RECORD_TYPES.EMR_STRETCHDIBITS);

            var fileBytes = File.ReadAllBytes("C:\\Users\\OssianEdström\\Pictures\\ResizedAsExcel.bmp");

            var br = new BinaryReader(new MemoryStream(fileBytes));

            br.BaseStream.Seek(0, SeekOrigin.Begin);
            dibits.bitMapHeader = new BitmapHeader(br);
            dibits.cbBmiSrc = dibits.bitMapHeader.sizeOfHeader;
            //var sign = Encoding.ASCII.GetString(br.ReadBytes(2));    //BM for a Windows bitmap

            //var size = br.ReadInt32();
            //var reserved = br.ReadBytes(4);
            //var offsetData = br.ReadInt32();
            //var ihSize = br.ReadInt32();

            //dibits.cbBmiSrc = (uint)ihSize;
            //dibits.bitMapHeader = new BitmapHeader(br, (uint)ihSize);

            var nrOfColorEntries = dibits.bitMapHeader.nColors;
            if(nrOfColorEntries == 0)
            {
                nrOfColorEntries = (uint)Math.Pow(2, (uint)dibits.bitMapHeader.colorDepth);
            }

            //br.ReadBytes((int)nrOfColorEntries);
            dibits.Padding2 = new byte[0];
            dibits.Size -= (uint)dibits.Padding2.Length;
            dibits.Padding2 = br.ReadBytes(dibits.bitMapHeader.offset - (int)br.BaseStream.Position);
            dibits.Size += (uint)dibits.Padding2.Length;
            if (dibits.Size % 4 != 0)
            {
                int paddingBytes = (int)(4 - (dibits.Size % 4)) % 4;
                dibits.EndPadding = new byte[paddingBytes];
                dibits.Size += (uint)paddingBytes;
            }

            br.BaseStream.Position = dibits.bitMapHeader.offset;

            var test = fileBytes.Length - (int)br.BaseStream.Position;
            var length = fileBytes.Length - dibits.bitMapHeader.offset;

            var srcsBits = br.ReadBytes(length);

           //var srcsBits = br.ReadBytes(fileBytes.Length - (int)br.BaseStream.Position);

            dibits.BitsSrc = srcsBits;

            dibits.Bounds = new RectLObject(9, 59, 117, 102);

            emfImage.Save("C:\\epplusTest\\Testoutput\\ValidStampAltered.emf");

            var readImage = new EmfImage();
            readImage.Read("C:\\epplusTest\\Testoutput\\ValidStampAltered.emf");

            var readExcelVersion = new EmfImage();
            readExcelVersion.Read("C:\\epplusTest\\Testoutput\\TemplateBmp.emf");

            var records2 = readImage.records;

            var dibits1 = (EMR_STRETCHDIBITS)readImage.records.Find(x => x.Type == RECORD_TYPES.EMR_STRETCHDIBITS);

            var targetRecord = readExcelVersion.records;
            var dibits2 = (EMR_STRETCHDIBITS)readExcelVersion.records.Find(x => x.Type == RECORD_TYPES.EMR_STRETCHDIBITS);
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

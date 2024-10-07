using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.EMF;
using System;
using System.Collections.Generic;
using System.Linq;

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


                var emf = new EMF();
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

                var emf = new EMF();
                emf.Read(path);

                emf.Save("C:\\epplusTest\\Workbooks\\GeneratedTwo.emf");
            }
        }

        [TestMethod]
        public void ReadEmfAlt()
        {
            var emfImage = new EMF();
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
        public void CheckValidTemplateEnhanced()
        {
            var validTemplate = new SignatureLineTemplateValid();
            //var records = validTemplate.records;

            Dictionary<float, short> _fontWidthDefault = null;
            if (FontSize.FontWidths.ContainsKey(FontSize.NonExistingFont))
            {
                FontSize.LoadAllFontsFromResource();
                _fontWidthDefault = FontSize.FontWidths[FontSize.NonExistingFont];
            }

            var pck = new ExcelPackage();

            //pck.Settings.TextSettings.DefaultTextMeasurer.

            var width = FontSize.GetFontSize("SegoeUI", true);

            //validTemplate.timeStamp.Text = "TimeStamp";
            //validTemplate.signTextObject.Text = "TemplateSignature";
            //validTemplate.suggestedSignerObject.Text = "TemplateSigner";
            //validTemplate.suggestedTitleObject.Text = validTemplate.suggestedTitleObject.Text;
            //validTemplate.Signers = "TemplateName";

            validTemplate.Save("C:\\epplusTest\\Testoutput\\testTemp.emf");
        }


        [TestMethod]
        public void CheckFontStuff()
        {
            var validTemplate = new EMF();
            //validTemplate.Read("C:\\Users\\OssianEdström\\Documents\\presentationEmf.emf");
            //validTemplate.Read("C:\\Users\\OssianEdström\\Pictures\\Segoe_UI_pt10.emf");
            //validTemplate.Read("C:\\epplusTest\\Testoutput\\ValidImageOriginal.emf");
            validTemplate.Read(@"C:\Users\OssianEdström\Desktop\AlphabetEmf.emf");

            var records = validTemplate.records;

            //var clipRect = (EMR_INTERSECTCLIPRECT)records[121];
            //clipRect.Clip.Left = 41;
            //clipRect.Clip.Top = 51;
            //clipRect.Clip.Right = 242;
            //clipRect.Clip.Bottom = 72;
            var fontRecordArr = records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTCREATEFONTINDIRECTW);
            var textRecords = records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW).ToArray();
            ((EMR_EXTTEXTOUTW)textRecords[1]).Text = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
            ((EMR_EXTTEXTOUTW)textRecords[2]).Text = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            //var timeStamp = (EMR_EXTTEXTOUTW)textRecords[0];
            //var signTextObject = (EMR_EXTTEXTOUTW)textRecords[2];
            //var suggestedSignerObject = (EMR_EXTTEXTOUTW)textRecords[3];
            //var suggestedTitleObject = (EMR_EXTTEXTOUTW)textRecords[4];
            //var signedBy = (EMR_EXTTEXTOUTW)textRecords[5];
            var header = ((EMR_HEADER)records[0]);
            var cx = BitConverter.ToUInt32(header.Device, 0);
            var cy = BitConverter.ToUInt32(header.Device, 4);

            var cxMili = BitConverter.ToUInt32(header.Millimeters, 0);
            var cyMili = BitConverter.ToUInt32(header.Millimeters, 4);

            var inchesX = cxMili * 0.0393700787f;
            var inchesY = cyMili * 0.0393700787f;
            //timeStamp.Text = "TimeStamp";
            //signTextObject.Text = "TemplateSignature";
            //suggestedSignerObject.Text = "TemplateSigner";
            //suggestedTitleObject.Text = "TemplateTitle";
            //signedBy.Text = "Signed by: TemplateName";

            validTemplate.Save("C:\\epplusTest\\Testoutput\\testTempPresentation.emf");
        }

        [TestMethod]
        public void CheckFontStuff2()
        {
            var validTemplate = new EMF();

            validTemplate.Read("C:\\Users\\OssianEdström\\Pictures\\TestStuff.emf");
            var records = validTemplate.records;

            var fontRecordArr = records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTCREATEFONTINDIRECTW);
            var textRecords = records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW).ToArray();
            ((EMR_EXTTEXTOUTW)textRecords[1]).Text = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";

            validTemplate.Save("C:\\epplusTest\\Testoutput\\testTempPresentation.emf");
        }

        [TestMethod]
        public void CheckFontEnlarged()
        {
            var validTemplate = new EMF();
            //validTemplate.Read("C:\\Users\\OssianEdström\\Documents\\TestingEmf.emf");
            validTemplate.Read(@"C:\Users\OssianEdström\Pictures\Picture9.emf");
            var records = validTemplate.records;

            var bounds = ((EMR_HEADER)records[0]).Bounds;
            bounds.Bottom = 129;
            bounds.Right = 257;

            //records.Remove(records[69]);
            //records.Remove(records[68]);

            var fontRecordArr = records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTCREATEFONTINDIRECTW);
            var textRecords = records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW).ToArray();
            ((EMR_EXTTEXTOUTW)textRecords[0]).Text = "DifferentText";

            validTemplate.Save("C:\\epplusTest\\Testoutput\\TestingEmf.emf");
        }

        [TestMethod]
        public void CheckValidTemplate()
        {
            var validTemplate = new EMF();
            validTemplate.Read("C:\\epplusTest\\Testoutput\\ValidImage.emf");
            var records = validTemplate.records;

            var clipRect = (EMR_INTERSECTCLIPRECT)records[121];
            clipRect.Clip.Left = 41;
            clipRect.Clip.Top = 51;
            clipRect.Clip.Right = 242;
            clipRect.Clip.Bottom = 72;

            var textRecords = records.FindAll(x => x.Type == RECORD_TYPES.EMR_EXTTEXTOUTW).ToArray();
            var timeStamp = (EMR_EXTTEXTOUTW)textRecords[0];
            var signTextObject = (EMR_EXTTEXTOUTW)textRecords[2];
            var suggestedSignerObject = (EMR_EXTTEXTOUTW)textRecords[3];
            var suggestedTitleObject = (EMR_EXTTEXTOUTW)textRecords[4];
            var signedBy = (EMR_EXTTEXTOUTW)textRecords[5];

            timeStamp.Text = "TimeStamp";
            signTextObject.Text = "TemplateSignature";
            suggestedSignerObject.Text = "TemplateSigner";
            suggestedTitleObject.Text = "TemplateTitle";
            signedBy.Text = "Signed by: TemplateName";

            validTemplate.Save("C:\\epplusTest\\Testoutput\\testTemp.emf");
        }
    }
}

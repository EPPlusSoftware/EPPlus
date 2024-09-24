using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Drawing.EMF;
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

            emfImage.records[140].data = new byte[] { 3, 0, 0, 0 };

            emfImage.Save("C:\\epplusTest\\Testoutput\\ChangeFontOutput.emf");
        }
    }
}

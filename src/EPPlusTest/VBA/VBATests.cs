using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.VBA
{
    [TestClass]
    public class VBATests : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenPackage("Vba.xlsm", true);
            _pck.Workbook.CreateVBAProject();
        }

        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void Compression()
        {
            //Compression/Decompression
            string value = "#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa";

            byte[] compValue = VBACompression.CompressPart(Encoding.GetEncoding(1252).GetBytes(value));
            string decompValue = Encoding.GetEncoding(1252).GetString(VBACompression.DecompressPart(compValue));
            Assert.AreEqual(value, decompValue);

            value = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";

            compValue = VBACompression.CompressPart(Encoding.GetEncoding(1252).GetBytes(value));
            decompValue = Encoding.GetEncoding(1252).GetString(VBACompression.DecompressPart(compValue));
            Assert.AreEqual(value, decompValue);
        }
        [TestMethod]
        public void WriteVBA()
        {
            _pck.Workbook.Worksheets.Add("WriteVBA");
            _pck.Workbook.VbaProject.Modules["WriteVBA"].Code += "\r\nPrivate Sub Worksheet_SelectionChange(ByVal Target As Range)\r\nMsgBox(\"Test of the VBA Feature!\")\r\nEnd Sub\r\n";
            _pck.Workbook.VbaProject.Modules["WriteVBA"].Name = "Blad1";
            _pck.Workbook.CodeModule.Name = "DenHärArbetsboken";
            _pck.Workbook.Worksheets[0].Name = "FirstSheet";
            //_pck.Workbook.CodeModule.Code += "\r\nPrivate Sub Workbook_Open()\r\nBlad1.Cells(1,1).Value = \"VBA test\"\r\nMsgBox \"VBA is running!\"\r\nEnd Sub";
            //X509Store store = new X509Store(StoreLocation.CurrentUser);
            //store.Open(OpenFlags.ReadOnly);
            //package.Workbook.VbaProject.Signature.Certificate = store.Certificates[11];

            var m = _pck.Workbook.VbaProject.Modules.AddModule("Module1");
            m.Code += "Public Sub Test(param1 as string)\r\n\r\nEnd sub\r\nPublic Function functest() As String\r\n\r\nEnd Function\r\n";
            var c = _pck.Workbook.VbaProject.Modules.AddClass("Class1", false);
            c.Code += "Private Sub Class_Initialize()\r\n\r\nEnd Sub\r\nPrivate Sub Class_Terminate()\r\n\r\nEnd Sub";
            var c2 = _pck.Workbook.VbaProject.Modules.AddClass("Class2", true);
            c2.Code += "Private Sub Class_Initialize()\r\n\r\nEnd Sub\r\nPrivate Sub Class_Terminate()\r\n\r\nEnd Sub";

            _pck.Workbook.VbaProject.Protection.SetPassword("EPPlus");
        }
        [TestMethod]
        public void WriteLongVBAModule()
        {
            _pck.Workbook.Worksheets.Add("VBASetData");
            _pck.Workbook.CodeModule.Code = "Private Sub Workbook_Open()\r\nCreateData\r\nEnd Sub";
            var module = _pck.Workbook.VbaProject.Modules.AddModule("Code");

            StringBuilder code = new StringBuilder("Public Sub CreateData()\r\n");
            for (int row = 1; row < 30; row++)
            {
                for (int col = 1; col < 30; col++)
                {
                    code.AppendLine(string.Format("VBASetData.Cells({0},{1}).Value=\"Cell {2}\"", row, col, new ExcelAddressBase(row, col, row, col).Address));
                }
            }
            code.AppendLine("End Sub");
            module.Code = code.ToString();

            //X509Store store = new X509Store(StoreLocation.CurrentUser);
            //store.Open(OpenFlags.ReadOnly);
            //package.Workbook.VbaProject.Signature.Certificate = store.Certificates[19];
        }
        [TestMethod]
        public void CreateUnicodeWsName()
        {
            ExcelWorksheet worksheet = _pck.Workbook.Worksheets.Add("测试");

            var sb = new StringBuilder();
            sb.AppendLine("Sub GetData()");
            sb.AppendLine("MsgBox (\"Hello,World\")");
            sb.AppendLine("End Sub");

            ExcelWorksheet worksheet2 = _pck.Workbook.Worksheets.Add("Sheet1");
            var stringBuilder = new StringBuilder();
            stringBuilder.AppendLine("Private Sub Worksheet_Change(ByVal Target As Range)");
            stringBuilder.AppendLine("GetData");
            stringBuilder.AppendLine("End Sub");
            worksheet.CodeModule.Code = stringBuilder.ToString();
        }

        [TestMethod]
        public void ValidateName()
        {
            using (var p = new ExcelPackage())
            {
                p.Workbook.CreateVBAProject();
                p.Workbook.Worksheets.Add("Work!Sheet");
                p.Workbook.Worksheets.Add("Mod=ule1");

                Assert.AreEqual("ThisWorkbook", p.Workbook.VbaProject.Modules[0].Name);
                Assert.AreEqual("Sheet0", p.Workbook.VbaProject.Modules[1].Name);
                Assert.AreEqual("Sheet1", p.Workbook.VbaProject.Modules[2].Name);
            }
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void ModuleNameContainsInvalidCharacters()
        {
            using (var p = new ExcelPackage())
            {
                p.Workbook.Worksheets.Add("InvalidName");
                p.Workbook.CreateVBAProject();
                p.Workbook.VbaProject.Modules.AddModule("Mod%ule2");
            }
        }
    }
}

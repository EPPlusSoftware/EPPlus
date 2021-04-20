using EPPlusTest;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;


namespace EPPlusTest.Core
{
    [TestClass]
    public class ExternalReferencesTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            //_pck = OpenPackage("ExternalReferences.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            //var dirName = _pck.File.DirectoryName;
            //var fileName = _pck.File.FullName;

            //SaveAndCleanup(_pck);

            //if (File.Exists(fileName)) File.Copy(fileName, dirName + "\\ExternalReferencesRead.xlsx", true);
        }
        [TestMethod]
        public void OpenAndReadExternalReferences()
        {
            var p = OpenTemplatePackage("ExtRef.xlsx");

            Assert.AreEqual(1, p.Workbook.ExternalReferences.Count);

            Assert.AreEqual(1D, p.Workbook.ExternalReferences[0].Worksheets["sheet1"].CellValues["A2"].Value);
            Assert.AreEqual(12D, p.Workbook.ExternalReferences[0].Worksheets["sheet1"].CellValues["C3"].Value);

            var c = 0;
            foreach(var cell in p.Workbook.ExternalReferences[0].Worksheets["sheet1"].CellValues)
            {
                Assert.IsNotNull(cell.Value);
                c++;
            }
            Assert.AreEqual(11, c);
        }
        [TestMethod]
        public void OpenAndReadExternalReferences1()
        {
            var p = OpenTemplatePackage("ExternalReferencesText1.xlsx");

            Assert.AreEqual(62, p.Workbook.ExternalReferences.Count);

            var c = 0;
            foreach (var cell in p.Workbook.ExternalReferences[0].Worksheets[0].CellValues)
            {
                Assert.IsNotNull(cell.Value);
                c++;
            }
            Assert.AreEqual(104, c);
        }

        [TestMethod]
        public void OpenAndReadExternalReferences2()
        {
            var p = OpenTemplatePackage("ExternalReferencesText2.xlsx");

            Assert.AreEqual(204, p.Workbook.ExternalReferences.Count);

            var c = 0;
            foreach (var cell in p.Workbook.ExternalReferences[0].Worksheets[0].CellValues)
            {
                Assert.IsNotNull(cell.Value);
                c++;
            }
            Assert.AreEqual(104, c);
        }
        [TestMethod]
        public void OpenAndReadExternalReferences3()
        {
            var p = OpenTemplatePackage("ExternalReferencesText3.xlsx");

            Assert.AreEqual(63, p.Workbook.ExternalReferences.Count);

            var c = 0;
            foreach (var cell in p.Workbook.ExternalReferences[0].Worksheets[0].CellValues)
            {
                Assert.IsNotNull(cell.Value);
                c++;
            }
            Assert.AreEqual(104, c);
        }

    }
}

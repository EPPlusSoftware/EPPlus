using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
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
        }

        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void ValidateName()
        {
            _pck.Workbook.Worksheets.Add("Work!Sheet");
            _pck.Workbook.CreateVBAProject();
            _pck.Workbook.Worksheets.Add("Mod=ule1");

            Assert.AreEqual("ThisWorkbook", _pck.Workbook.VbaProject.Modules[0].Name);
            Assert.AreEqual("Sheet0", _pck.Workbook.VbaProject.Modules[1].Name);
            Assert.AreEqual("Sheet1", _pck.Workbook.VbaProject.Modules[2].Name);
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

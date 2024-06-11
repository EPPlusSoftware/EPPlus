using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class IterativeCalculationsTest
    {
        private ExcelPackage _package;
        private ExcelWorksheet _sheet;

        [TestInitialize]
        public void Initialize()
        {
            _package = new ExcelPackage();
            _sheet = _package.Workbook.Worksheets.Add("test");
        }

        [TestCleanup]
        public void Cleanup()
        {
            _sheet = null;
            _package.Dispose();
        }

        //[TestMethod]
        //public void A1andB1CircularRegShouldWork()
        //{
        //    _dateWs1.Cells["B1"].Value = 1;
        //    _dateWs1.Cells["A1"].Formula = "A2 + B1";
        //    _dateWs1.Cells["A2"].Formula = "A1 + B1";

        //    var options = new ExcelCalculationOption { AllowCircularReferences = true };
        //    _dateWs1.Calculate(options);

        //    Assert.AreEqual(1d, _dateWs1.Cells["A1"].Value);
        //    Assert.AreEqual(2d, _dateWs1.Cells["B1"].Value);
        //}

    }
}

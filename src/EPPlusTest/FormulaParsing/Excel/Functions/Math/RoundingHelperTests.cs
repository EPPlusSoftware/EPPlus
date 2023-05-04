using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
    [TestClass]
    public class RoundingHelperTests
    {
        [TestMethod]
        public void Below1AndBelowMinus1()
        {
            var n = -120253.87499999999;
            var result = RoundingHelper.RoundToSignificantFig(n, 15);
            Assert.AreEqual(-120253.875, result);

            n = 120253.87499999999;
            result = RoundingHelper.RoundToSignificantFig(n, 15);
            Assert.AreEqual(120253.875, result);
        }

        [TestMethod]
        public void FromMinus1To1()
        {
            var n = -0.0699999999999994;
            var result = RoundingHelper.RoundToSignificantFig(n, 15);
            Assert.AreEqual(-0.07, result);

            n = -0.06999999999999;
            result = RoundingHelper.RoundToSignificantFig(n, 15);
            Assert.AreEqual(-0.06999999999999, result);

            n = 0.0700000000000003;
            result = RoundingHelper.RoundToSignificantFig(n, 15);
            Assert.AreEqual(0.07, result);

            using (var p = new ExcelPackage())
            {
                var sheet = p.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = -0.0699999999999994;
                sheet.Cells["A2"].Formula = "ROUND(A1, 2)";
                sheet.Cells["B1"].Value = 0.0700000000000003;
                sheet.Cells["B2"].Formula = "ROUND(B1, 2)";
                sheet.Calculate(opt => opt.PrecisionAndRoundingStrategy = PrecisionAndRoundingStrategy.Excel);
                Assert.AreEqual(-0.07, sheet.Cells["A2"].Value);
                Assert.AreEqual(0.07, sheet.Cells["B2"].Value);
            }
        }
    }
}

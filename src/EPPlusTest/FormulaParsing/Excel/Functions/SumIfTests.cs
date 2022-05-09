using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions
{
    [TestClass]
    public class SumIfTests
    {
        [TestMethod]
        public void SumIf_SumThisRowWithoutCircularReferences()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet1 = pck.Workbook.Worksheets.Add("Sheet1");
                sheet1.Cells["A1"].Value = "SumResult";
                // This shouldn't be a circular reference, because the 1:1="SUMMABLE" condition should filter out A2
                sheet1.Cells["A2"].Formula = "SUMIF(1:1,\"SUMMABLE\",2:2)";

                sheet1.Cells["B2"].Value = 1;
                sheet1.Cells["C2"].Value = 2;
                sheet1.Cells["D2"].Value = 3;
                sheet1.Cells["E2"].Value = 4;
                sheet1.Cells["F2"].Value = 5;
                sheet1.Cells["G2"].Value = 6;

                sheet1.Cells["C1"].Value = "SUMMABLE";
                sheet1.Cells["D1"].Value = "SUMMABLE";
                sheet1.Cells["E1"].Value = "SUMMABLE";
                sheet1.Cells["G1"].Value = "SUMMABLE";

                pck.Workbook.Calculate(x => x.AllowCircularReferences = true);

                Assert.AreEqual(15, sheet1.Cells["A2"].GetValue<double>(), double.Epsilon);
            }
        }

        [TestMethod]
        public void SumIf_SumThisColWithoutCircularReferences()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet1 = pck.Workbook.Worksheets.Add("Sheet1");
                sheet1.Cells["A1"].Value = "SumResult";
                // This shouldn't be a circular reference, because the 1:1="SUMMABLE" condition should filter out A2
                sheet1.Cells["B1"].Formula = "SUMIF(A:A,\"SUMMABLE\",B:B)";

                sheet1.Cells["B2"].Value = 1;
                sheet1.Cells["B3"].Value = 2;
                sheet1.Cells["B4"].Value = 3;
                sheet1.Cells["B5"].Value = 4;
                sheet1.Cells["B6"].Value = 5;
                sheet1.Cells["B7"].Value = 6;

                sheet1.Cells["A3"].Value = "SUMMABLE";
                sheet1.Cells["A4"].Value = "SUMMABLE";
                sheet1.Cells["A5"].Value = "SUMMABLE";
                sheet1.Cells["A7"].Value = "SUMMABLE";

                pck.Workbook.Calculate(x => x.AllowCircularReferences = true);

                Assert.AreEqual(15, sheet1.Cells["B1"].GetValue<double>(), double.Epsilon);
            }
        }
    }
}

using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Core
{
    [TestClass]
    public class FullPrecisionTests : TestBase
    {
        [ClassInitialize]
        public static void Init(TestContext context)
        {
        }
        [TestMethod]
        public void FullPrecisionShouldRoundValuesOnSetTest()
        {
            using(var p=new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");

                ws.Cells["A1"].Value = "String Value";
                ws.Cells["A2"].Value = 123.456789;
                
                ws.Cells["A3"].Value = 123.456789;
                ws.Cells["A4"].Value = -123.456789;
                ws.Cells["A3:A4"].Style.Numberformat.Format = "#,##0.00;-#,##0.000;0.0";

                p.Workbook.FullPrecision = false;

                Assert.AreEqual("String Value", ws.Cells["A1"].Value);
                Assert.AreEqual(123.456789, ws.Cells["A2"].Value);
                Assert.AreEqual(123.46, ws.Cells["A3"].Value);
                Assert.AreEqual(-123.457, ws.Cells["A4"].Value);
            }
        }
        [TestMethod]
        public void ShouldRoundValuesWhenSetValueOnRangeWithFullPrecisionFalseTest()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");

                p.Workbook.FullPrecision = false;

                ws.Cells["A3:A4"].Style.Numberformat.Format = "#,##0.00;-#,##0.000;0.0";

                ws.Cells["A1"].Value = "String Value";
                ws.Cells["A2"].Value = 123.456789;

                ws.Cells["A3"].Value = 123.456789;
                ws.Cells["A4"].Value = -123.456789;

                Assert.AreEqual("String Value", ws.Cells["A1"].Value);
                Assert.AreEqual(123.456789, ws.Cells["A2"].Value);
                Assert.AreEqual(123.46, ws.Cells["A3"].Value);
                Assert.AreEqual(-123.457, ws.Cells["A4"].Value);
            }
        }
        [TestMethod]
        public void ShouldRoundValuesWhenSetNumberFormatOnRangeWithFullPrecisionFalseTest()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");

                p.Workbook.FullPrecision = false;

                ws.Cells["A1"].Value = "String Value";
                ws.Cells["A2"].Value = 123.456789;

                ws.Cells["A3"].Value = 123.456789;
                ws.Cells["A4"].Value = -123.456789;

                ws.Cells["A3:A4"].Style.Numberformat.Format = "#,##0.00;-#,##0.000;0.0";

                Assert.AreEqual("String Value", ws.Cells["A1"].Value);
                Assert.AreEqual(123.456789, ws.Cells["A2"].Value);
                Assert.AreEqual(123.46, ws.Cells["A3"].Value);
                Assert.AreEqual(-123.457, ws.Cells["A4"].Value);
            }
        }

    }
}

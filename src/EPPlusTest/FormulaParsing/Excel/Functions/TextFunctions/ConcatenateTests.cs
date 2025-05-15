using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.TextFunctions
{
    [TestClass]
    public class ConcatenateTests: TestBase
    {
        [TestMethod]
        public void ConcatenateDates()
        {
            using (var p = OpenTemplatePackage("s854.xlsx"))
            {
                var ws = p.Workbook.Worksheets[7];
                ws.Cells["A52"].Calculate();

                var val = ws.Cells["A52"].Value;

                Assert.AreEqual("201901", val);

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void ConcatenateDates2()
        {
            using (var p = OpenPackage("s854_epplusGenerated.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");

                ws.Cells["G1:G5"].Formula = "2000+ROW()";
                ws.Cells["H1:H5"].Value = "001";

                ws.Cells["J1"].Formula = "G1:G5 & H1:H5";
                ws.Cells["G1:Z59"].Calculate();

                var val1 = ws.Cells["J1"].GetValue<string>();

                Assert.AreEqual("2001001", val1);

                SaveAndCleanup(p);
            }
        }
    }
}

using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace EPPlusTest.FormulaParsing.Excel.Functions.TextFunctions
{
    [TestClass]
    public class SubstituteTests
    {
        [TestMethod]
        public void SubstituteShouldHandleRangeArguments()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet1");
            ws.Cells["E1"].Value = "a";
            ws.Cells["E2"].Value = "b";
            ws.Cells["E3"].Value = "c";
            ws.Cells["F1"].Value = "e";
            ws.Cells["F2"].Value = "f";
            ws.Cells["F3"].Value = "g";

            ws.Cells["H8"].Formula = "SUBSTITUTE(\"abc123\",E1:E3,F1:F3)";

            ws.Calculate();

            Assert.AreEqual("ebc123", ws.Cells["H8"].Value);
            Assert.AreEqual("afc123", ws.Cells["H9"].Value);
            Assert.AreEqual("abg123", ws.Cells["H10"].Value);

        }
    }
}

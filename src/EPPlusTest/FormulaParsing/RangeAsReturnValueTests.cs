using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class RangeAsReturnValueTests
    {
        [TestMethod]
        public void IndirectShouldHandleRangeInGroupExpression()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Formula = "MATCH(1;(OFFSET(INDIRECT(\"A1:A2\"),0,0)))";
                sheet.Calculate();
                var v = sheet.Cells["A3"].Value;
                Assert.AreEqual(1, v);
            }
        }
    }
}

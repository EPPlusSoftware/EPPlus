using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlusTest.Issues
{
    [TestClass]
    public class DefinedNameIssues : TestBase
    {
        [TestMethod]
        public void I1238()
        {
            using (var p = OpenTemplatePackage("I1238SlowWorkbook.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                ws.Cells["A1"].Value = 1;
                SaveAndCleanup(p);
            }
        }
		[TestMethod]
		public void s652()
        {
            using (var p = OpenTemplatePackage("s652.xlsm"))
            {
                using var p2 = new ExcelPackage();
                var ws = p.Workbook.Worksheets[0];
                p2.Workbook.Worksheets.Add("New ws", ws);
                SaveWorkbook("s652.xlsx", p2);
            }
		}
    }
}

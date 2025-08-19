using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.Issues
{
    [TestClass]
    public class ExportIssues : TestBase
    {
        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void ToDataTableDuplicateColumnNames()
        {
            var package = OpenTemplatePackage("ToDataTableDuplicateColumnNames.xlsx");
            var sheet = package.Workbook.Worksheets[0];
            var dt = sheet.Cells["A1:C3"].ToDataTable(x =>
            {
                x.FirstRowIsColumnNames = true;
            });
        }

        [TestMethod]
        public void ToDataTableDuplicateColumnNames2()
        {
            var package = OpenTemplatePackage("ToDataTableDuplicateColumnNames.xlsx");
            var sheet = package.Workbook.Worksheets[0];
            var dt = sheet.Cells["A1:C3"].ToDataTable(x =>
            {
                x.AllowDuplicateColumnNames = true;
            });
            // rows and column count
            Assert.AreEqual(2, dt.Rows.Count);
            Assert.AreEqual(3, dt.Columns.Count);
            // column names
            Assert.AreEqual("Id", dt.Columns[0].ColumnName);
            Assert.AreEqual("Name1", dt.Columns[1].ColumnName);
            Assert.AreEqual("Name2", dt.Columns[2].ColumnName);
            // row values
            Assert.AreEqual(1d, dt.Rows[0][0]);
            Assert.AreEqual("Name 1", dt.Rows[0][1]);
            Assert.AreEqual("Name 2", dt.Rows[0][2]);
        }
    }
}

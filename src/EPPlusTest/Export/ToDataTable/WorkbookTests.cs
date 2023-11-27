using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Export.ToDataTable
{
    [TestClass]
    public class WorkbookTests : TestBase
    {
        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void ShouldThrowInvalidOperationExceptionWhenNullNotAllowed()
        {
            using (var package = OpenTemplatePackage("ToDataTableNullValues.xlsx"))
            {
                var sheet1 = package.Workbook.Worksheets[0];
                var dt = sheet1.Cells["A1:B4"].ToDataTable(o => o.AlwaysAllowNull = false);
            }
        }

        [TestMethod]
        public void ShouldExportNullValues()
        {
            using (var package = OpenTemplatePackage("ToDataTableNullValues.xlsx"))
            {
                var sheet1 = package.Workbook.Worksheets[0];
                var dt = sheet1.Cells["A1:B4"].ToDataTable(o => o.AlwaysAllowNull = true);
                Assert.AreEqual(2, dt.Columns.Count);
                Assert.AreEqual(DBNull.Value, dt.Rows[1]["Id"]);
            }
        }
    }
}
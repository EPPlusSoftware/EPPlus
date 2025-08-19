using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;

namespace EPPlusTest
{
    [TestClass]
    public class MaxRowsTest : TestBase
    {
        [TestMethod]
        public void DeletingAtMaxRowsOfExcelSheetShouldNotThrow()
        {
            var package = new ExcelPackage();

            var sheet = package.Workbook.Worksheets.Add("DeletingAtMaxSheet");

            sheet.Cells["A1048576"].Value = 5;

            //Verify end and start of sheet are as expected.
            Assert.AreEqual(1048576, sheet.Dimension.End.Row);
            Assert.AreEqual(1, sheet.Dimension.Start.Column);
            Assert.AreEqual(5, sheet.Cells[sheet.Dimension.End.Row, sheet.Dimension.Start.Column].Value);

            sheet.DeleteRow(ExcelPackage.MaxRows);

            Assert.IsNull(sheet.Dimension);
            Assert.IsNull(sheet.Cells["A1048576"].Value);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void DeletingAtMaxRowsOfExcelSheetShouldThrow()
        {
            var package = new ExcelPackage();

            var sheet = package.Workbook.Worksheets.Add("DeletingAtMaxSheet");


            sheet.Cells["A1048576"].Value = 5;

            //Verify end and start of sheet are as expected.
            Assert.AreEqual(1048576, sheet.Dimension.End.Row);
            Assert.AreEqual(1, sheet.Dimension.Start.Column);
            Assert.AreEqual(5, sheet.Cells[sheet.Dimension.End.Row, sheet.Dimension.Start.Column].Value);

            sheet.DeleteRow(ExcelPackage.MaxRows, 2);
        }
    }
}

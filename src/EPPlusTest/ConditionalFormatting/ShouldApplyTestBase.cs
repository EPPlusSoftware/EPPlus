using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using System;

namespace EPPlusTest.ConditionalFormatting
{
    public class ShouldApplyTestBase : TestBase
    {
        protected ExcelWorksheet CreatePackageSheet(string name)
        {
            var pck = OpenPackage($"{name}.xlsx", true);
            return pck.Workbook.Worksheets.Add($"{name}");
        }

        protected void AssertConditionalFormat<T>(ExcelWorksheet sheet, DateTime date, T cfClass) where T : ExcelConditionalFormattingRule
        {
            sheet.Cells["A1"].Style.Numberformat.Format = "yyyy-mm-dd";
            sheet.Cells["A1"].Formula = $"=DATE({date.Year},{date.Month},{date.Day})";

            sheet.Cells["A2"].Style.Numberformat.Format = "yyyy-mm-dd";
            sheet.Cells["A2"].Formula = "=DATE(2023,10,5)";

            sheet.Cells["A3"].Style.Numberformat.Format = "yyyy-mm-dd";
            sheet.Cells["A3"].Formula = "=DATE(2023,6,5)";

            sheet.Cells["A1:A5"].Calculate();

            Assert.IsTrue(cfClass.ShouldApplyToCell(sheet.Cells["A1"]));
            Assert.IsFalse(cfClass.ShouldApplyToCell(sheet.Cells["A2"]));
            Assert.IsFalse(cfClass.ShouldApplyToCell(sheet.Cells["A3"]));
        }
    }
}

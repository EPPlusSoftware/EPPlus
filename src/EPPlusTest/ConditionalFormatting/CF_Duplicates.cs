using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;

namespace EPPlusTest.ConditionalFormatting
{
    [TestClass]
    public class CF_Duplicates : TestBase
    {
        [TestMethod]
        public void DuplicatesShouldApplyCorrectly()
        {
            using (var p = new ExcelPackage())
            {
                var sheet = p.Workbook.Worksheets.Add("duplicates");

                var range = sheet.Cells["A1:A15"];

                //Duplicates ignore case in excel.
                sheet.Cells["A1"].Value = "bye";
                sheet.Cells["A2"].Value = "Bye";
                sheet.Cells["A3"].Value = "buh-bye";
                sheet.Cells["A4"].Value = "someValue";
                sheet.Cells["A5"].Value = "bye";
                sheet.Cells["A6"].Value = 5;
                sheet.Cells["A7"].Value = "numbers";
                sheet.Cells["A8"].Value = 5;
                sheet.Cells["A9"].Value = 6;
                sheet.Cells["A10"].Value = 10;
                sheet.Cells["A11"].Value = "a";
                sheet.Cells["A12"].Value = "A";



                var cf = (ExcelConditionalFormattingDuplicateValues)range.ConditionalFormatting.AddDuplicateValues();

                Assert.IsTrue(cf.ShouldApplyToCell(sheet.Cells["A1"]));
                Assert.IsTrue(cf.ShouldApplyToCell(sheet.Cells["A2"]));
                Assert.IsFalse(cf.ShouldApplyToCell(sheet.Cells["A3"]));
                Assert.IsFalse(cf.ShouldApplyToCell(sheet.Cells["A4"]));
                Assert.IsTrue(cf.ShouldApplyToCell(sheet.Cells["A5"]));
                Assert.IsTrue(cf.ShouldApplyToCell(sheet.Cells["A6"]));
                Assert.IsFalse(cf.ShouldApplyToCell(sheet.Cells["A7"]));
                Assert.IsTrue(cf.ShouldApplyToCell(sheet.Cells["A8"]));
                Assert.IsFalse(cf.ShouldApplyToCell(sheet.Cells["A9"]));
                Assert.IsFalse(cf.ShouldApplyToCell(sheet.Cells["A10"]));
                Assert.IsTrue(cf.ShouldApplyToCell(sheet.Cells["A11"]));
                Assert.IsTrue(cf.ShouldApplyToCell(sheet.Cells["A12"]));
            }
        }
    }
}

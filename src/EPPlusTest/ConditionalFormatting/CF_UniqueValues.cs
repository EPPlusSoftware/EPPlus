using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.ConditionalFormatting
{
    [TestClass]
    public class CF_UniqueValues
    {
        [TestMethod]
        public void CF_UniqueValuesShouldApply()
        {
            using (var p = new ExcelPackage())
            {
                var sheet = p.Workbook.Worksheets.Add("uniqueSheet");

                var range = sheet.Cells["A1:A20"];
                range.Formula = "ROW()";
                range.Calculate();

                sheet.Cells["A5"].Value = 6;
                sheet.Cells["A10"].Value = 3;
                sheet.Cells["A11"].Value = "bye";
                sheet.Cells["A12"].Value = "bye";
                sheet.Cells["A15"].Value = "hi";

                var cf = (ExcelConditionalFormattingUniqueValues)range.ConditionalFormatting.AddUniqueValues();

                Assert.IsFalse(cf.ShouldApplyToCell(sheet.Cells["A11"]));
                Assert.IsFalse(cf.ShouldApplyToCell(sheet.Cells["A5"]));
                Assert.IsFalse(cf.ShouldApplyToCell(sheet.Cells["A10"]));

                Assert.IsTrue(cf.ShouldApplyToCell(sheet.Cells["A15"]));
                Assert.IsTrue(cf.ShouldApplyToCell(sheet.Cells["A1"]));
            }
        }
    }
}

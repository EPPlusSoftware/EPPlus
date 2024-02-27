using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using System;

namespace EPPlusTest.ConditionalFormatting
{
    [TestClass]
    public class CF_TopBottom : TestBase
    {
        [TestMethod]
        public void CF_TopBottomShouldApply()
        {
            using (var pck = OpenPackage("CF_TopBottomApply.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("topBottom");

                
                for (int i = 1; i <= 20; i++)
                {
                    sheet.Cells[i, 1].Value = i;
                }

                var top = sheet.Cells["A1:A20"].ConditionalFormatting.AddTop();
                var bottom = sheet.Cells["A1:A20"].ConditionalFormatting.AddBottom();
                var topPercent = sheet.Cells["A1:A20"].ConditionalFormatting.AddTopPercent();
                var bottomPercent = sheet.Cells["A1:A20"].ConditionalFormatting.AddBottomPercent();

                bottom.Rank = 10;
                top.Rank = 10;

                Assert.IsFalse(((ExcelConditionalFormattingTopBottomGroup)top).ShouldApplyToCell(sheet.Cells["A1"]));
                Assert.IsTrue(((ExcelConditionalFormattingTopBottomGroup)top).ShouldApplyToCell(sheet.Cells["A18"]));

                Assert.IsFalse(((ExcelConditionalFormattingTopBottomGroup)topPercent).ShouldApplyToCell(sheet.Cells["A1"]));
                Assert.IsTrue(((ExcelConditionalFormattingTopBottomGroup)topPercent).ShouldApplyToCell(sheet.Cells["A19"]));

                Assert.IsTrue(((ExcelConditionalFormattingTopBottomGroup)bottom).ShouldApplyToCell(sheet.Cells["A1"]));
                Assert.IsFalse(((ExcelConditionalFormattingTopBottomGroup)bottom).ShouldApplyToCell(sheet.Cells["A18"]));

                Assert.IsTrue(((ExcelConditionalFormattingTopBottomGroup)bottomPercent).ShouldApplyToCell(sheet.Cells["A1"]));
                Assert.IsFalse(((ExcelConditionalFormattingTopBottomGroup)bottomPercent).ShouldApplyToCell(sheet.Cells["A18"]));
            }
        }

        [TestMethod]
        public void ShouldApplyDoesNotChangeCellValue()
        {
            using (var pck = OpenPackage("CF_TopBottomApplyCellValue.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("topBottomValue");

                for (int i = 1; i <= 20; i++)
                {
                    sheet.Cells[i, 1].Value = $"Text{i}";
                }

                var top = sheet.Cells["A1:A20"].ConditionalFormatting.AddTop();
                ((ExcelConditionalFormattingTopBottomGroup)top).ShouldApplyToCell(sheet.Cells["A1"]);
                Assert.AreEqual("Text1", sheet.Cells["A1"].Value);
            }
        }
    }
}

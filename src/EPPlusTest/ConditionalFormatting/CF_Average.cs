using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.ConditionalFormatting;

namespace EPPlusTest.ConditionalFormatting
{
    [TestClass]
    public class CF_Average : TestBase
    {
        [TestMethod]
        public void CF_AverageGroupShoulApply()
        {
            using (var pck = OpenPackage("Averages.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("AveragesSheet");

                var aboveAverage = sheet.Cells["A1:A11"].ConditionalFormatting.AddAboveAverage();
                var aboveOrEqual = sheet.Cells["A1:A11"].ConditionalFormatting.AddAboveOrEqualAverage();
                var belowAverage = sheet.Cells["A1:A11"].ConditionalFormatting.AddBelowAverage();
                var belowOrEqual = sheet.Cells["A1:A11"].ConditionalFormatting.AddBelowOrEqualAverage();


                for (int i = 1; i < 12; i++)
                {
                    sheet.Cells[i,1].Value = i;
                }

                var aboveAverageClass = (ExcelConditionalFormattingAverageGroup)aboveAverage;
                var aboveOrEqualAverageClass = (ExcelConditionalFormattingAverageGroup)aboveOrEqual;
                var belowAverageClass = (ExcelConditionalFormattingAverageGroup)belowAverage;
                var belowOrQualAverageClass = (ExcelConditionalFormattingAverageGroup)belowOrEqual;


                Assert.IsTrue(aboveAverageClass.ShouldApplyToCell(sheet.Cells["A8"]));
                Assert.IsFalse(aboveAverageClass.ShouldApplyToCell(sheet.Cells["A1"]));
                Assert.IsFalse(aboveAverageClass.ShouldApplyToCell(sheet.Cells["A6"]));

                Assert.IsFalse(aboveOrEqualAverageClass.ShouldApplyToCell(sheet.Cells["A5"]));
                Assert.IsTrue(aboveOrEqualAverageClass.ShouldApplyToCell(sheet.Cells["A6"]));
                Assert.IsTrue(aboveOrEqualAverageClass.ShouldApplyToCell(sheet.Cells["A7"]));

                Assert.IsFalse(belowAverageClass.ShouldApplyToCell(sheet.Cells["A8"]));
                Assert.IsTrue(belowAverageClass.ShouldApplyToCell(sheet.Cells["A1"]));
                Assert.IsFalse(belowAverageClass.ShouldApplyToCell(sheet.Cells["A6"]));

                Assert.IsTrue(belowOrQualAverageClass.ShouldApplyToCell(sheet.Cells["A5"]));
                Assert.IsTrue(belowOrQualAverageClass.ShouldApplyToCell(sheet.Cells["A6"]));
                Assert.IsFalse(belowOrQualAverageClass.ShouldApplyToCell(sheet.Cells["A7"]));
            }
        }

        [TestMethod]
        public void CF_AverageGroupShoulApplyIfOnlyColumns()
        {
            using (var pck = OpenPackage("AveragesColumns.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("ColumnAverages");

                for (int i = 1; i < 15; i++)
                {
                    sheet.Cells[1, i].Value = i;
                }

                //Note only 11 columns. To ensure empty columns not a problem
                var aboveAverage = sheet.Cells["A1:E1,G1:H1,K1:N1"].ConditionalFormatting.AddAboveAverage();
                //Average is 80/11 = 7.272727...

                var aboveAverageClass = (ExcelConditionalFormattingAverageGroup)aboveAverage;

                Assert.IsTrue(aboveAverageClass.ShouldApplyToCell(sheet.Cells["K1"]));
                Assert.IsFalse(aboveAverageClass.ShouldApplyToCell(sheet.Cells["G1"]));
                Assert.IsFalse(aboveAverageClass.ShouldApplyToCell(sheet.Cells["B1"]));
            }
        }

    }
}

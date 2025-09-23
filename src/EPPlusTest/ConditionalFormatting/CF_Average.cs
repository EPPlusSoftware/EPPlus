using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using System.Drawing;

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

        [TestMethod]
        public void CF_STDEVShouldApply()
        {
            using (var pck = OpenPackage("StandardDevApply.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("StandardDevApply");

                sheet.Cells["A1"].Value = -55;
                sheet.Cells["A2"].Value = -35;
                sheet.Cells["A3"].Value = -24;
                sheet.Cells["A4"].Value = -10;
                sheet.Cells["A5"].Value = -10;
                sheet.Cells["A6"].Value = -9;
                sheet.Cells["A7"].Value = -6;
                sheet.Cells["A8"].Value = -3;
                sheet.Cells["A9"].Value = -3;
                sheet.Cells["A10"].Value = -3;
                sheet.Cells["A11"].Value = -2;
                sheet.Cells["A12"].Value = -2;
                sheet.Cells["A13"].Value = -2;
                sheet.Cells["A14"].Value = -1;
                sheet.Cells["A15"].Value = -1;
                sheet.Cells["A16"].Value = -1;
                sheet.Cells["A17"].Value = 0;
                sheet.Cells["A18"].Value = 0;
                sheet.Cells["A19"].Value = 0;
                sheet.Cells["A20"].Value = 1;
                sheet.Cells["A21"].Value = 1;
                sheet.Cells["A22"].Value = 1;
                sheet.Cells["A23"].Value = 2;
                sheet.Cells["A24"].Value = 2;
                sheet.Cells["A25"].Value = 2;
                sheet.Cells["A26"].Value = 3;
                sheet.Cells["A27"].Value = 3;
                sheet.Cells["A28"].Value = 3;
                sheet.Cells["A29"].Value = 6;
                sheet.Cells["A30"].Value = 9;
                sheet.Cells["A31"].Value = 10;
                sheet.Cells["A32"].Value = 10;
                sheet.Cells["A33"].Value = 24;
                sheet.Cells["A34"].Value = 35;
                sheet.Cells["A35"].Value = 55;

                var range = sheet.Cells["A1:A35"];

                var above3 = range.ConditionalFormatting.AddAboveStdDev();
                var above2 = range.ConditionalFormatting.AddAboveStdDev();
                var above1 = range.ConditionalFormatting.AddAboveStdDev();

                var below3 = range.ConditionalFormatting.AddBelowStdDev();
                var below2 = range.ConditionalFormatting.AddBelowStdDev();
                var below1 = range.ConditionalFormatting.AddBelowStdDev();

                above1.StdDev = 1;
                above2.StdDev = 2;
                above3.StdDev = 3;

                below1.StdDev = 1; 
                below2.StdDev = 2; 
                below3.StdDev = 3;

                var stdevList = new List<IExcelConditionalFormattingStdDevGroup>() { above1, above2, above3, below1, below2, below3 };

                foreach( var stdev in stdevList )
                {
                    stdev.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                }

                above1.Style.Fill.BackgroundColor.Color = Color.Red;
                above2.Style.Fill.BackgroundColor.Color = Color.MediumVioletRed;
                above3.Style.Fill.BackgroundColor.Color = Color.DarkRed;

                below1.Style.Fill.BackgroundColor.Color = Color.Blue;
                below2.Style.Fill.BackgroundColor.Color = Color.MediumBlue;
                below3.Style.Fill.BackgroundColor.Color = Color.DarkBlue;


                Assert.IsTrue(((ExcelConditionalFormattingStdDevGroup)above1).ShouldApplyToCell(sheet.Cells["A33"]));
                Assert.IsTrue(((ExcelConditionalFormattingStdDevGroup)above2).ShouldApplyToCell(sheet.Cells["A34"]));
                Assert.IsTrue(((ExcelConditionalFormattingStdDevGroup)above3).ShouldApplyToCell(sheet.Cells["A35"]));

                Assert.IsFalse(((ExcelConditionalFormattingStdDevGroup)above1).ShouldApplyToCell(sheet.Cells["A32"]));
                Assert.IsFalse(((ExcelConditionalFormattingStdDevGroup)above2).ShouldApplyToCell(sheet.Cells["A33"]));
                Assert.IsFalse(((ExcelConditionalFormattingStdDevGroup)above3).ShouldApplyToCell(sheet.Cells["A34"]));

                Assert.IsTrue(((ExcelConditionalFormattingStdDevGroup)below1).ShouldApplyToCell(sheet.Cells["A3"]));
                Assert.IsTrue(((ExcelConditionalFormattingStdDevGroup)below2).ShouldApplyToCell(sheet.Cells["A2"]));
                Assert.IsTrue(((ExcelConditionalFormattingStdDevGroup)below3).ShouldApplyToCell(sheet.Cells["A1"]));

                //SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void CF_STDEVShouldApplyWithAverage()
        {
            using (var pck = OpenPackage("StandardDevApplyWithAverage.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("StandardDevApplyWithAverage");

                sheet.Cells["A1"].Value = -15;
                sheet.Cells["A2"].Value = -10.5;
                sheet.Cells["A3"].Value = -7;
                sheet.Cells["A4"].Value = -4;
                sheet.Cells["A5"].Value = -4;
                sheet.Cells["A6"].Value = -3;
                sheet.Cells["A7"].Value = -3;
                sheet.Cells["A8"].Value = -3;
                sheet.Cells["A9"].Value = -2;
                sheet.Cells["A10"].Value = -2;
                sheet.Cells["A11"].Value = -1;
                sheet.Cells["A12"].Value = -1;
                sheet.Cells["A13"].Value = -1;
                sheet.Cells["A14"].Value = -1;
                sheet.Cells["A15"].Value = -1;
                sheet.Cells["A16"].Value = -1;
                for(int i = 1; i <= 8; i++)
                {
                    sheet.Cells[16 + i, 1].Value = 0;
                }
                for (int i = 1; i <=6; i++)
                {
                    sheet.Cells[24 + i, 1].Value = 1;
                }
                sheet.Cells["A31"].Value = 2;
                sheet.Cells["A32"].Value = 2;
                sheet.Cells["A33"].Value = 3;
                sheet.Cells["A34"].Value = 3;
                sheet.Cells["A35"].Value = 3;
                sheet.Cells["A36"].Value = 4;
                sheet.Cells["A37"].Value = 4;
                sheet.Cells["A38"].Value = 6;
                sheet.Cells["A39"].Value = 10.5;
                sheet.Cells["A40"].Value = 15;

                var range = sheet.Cells["A1:A42"];

                var above3 = range.ConditionalFormatting.AddAboveStdDev();
                var above2 = range.ConditionalFormatting.AddAboveStdDev();
                var above1 = range.ConditionalFormatting.AddAboveStdDev();

                var below3 = range.ConditionalFormatting.AddBelowStdDev();
                var below2 = range.ConditionalFormatting.AddBelowStdDev();
                var below1 = range.ConditionalFormatting.AddBelowStdDev();

                above1.StdDev = 1;
                above2.StdDev = 2;
                above3.StdDev = 3;

                below1.StdDev = 1;
                below2.StdDev = 2;
                below3.StdDev = 3;

                var stdevList = new List<IExcelConditionalFormattingStdDevGroup>() { above1, above2, above3, below1, below2, below3 };

                foreach (var stdev in stdevList)
                {
                    stdev.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                }

                above1.Style.Fill.BackgroundColor.Color = Color.Red;
                above2.Style.Fill.BackgroundColor.Color = Color.MediumVioletRed;
                above3.Style.Fill.BackgroundColor.Color = Color.DarkRed;

                below1.Style.Fill.BackgroundColor.Color = Color.Blue;
                below2.Style.Fill.BackgroundColor.Color = Color.MediumBlue;
                below3.Style.Fill.BackgroundColor.Color = Color.DarkBlue;

                Assert.IsTrue(((ExcelConditionalFormattingStdDevGroup)above1).ShouldApplyToCell(sheet.Cells["A38"]));
                Assert.IsTrue(((ExcelConditionalFormattingStdDevGroup)above2).ShouldApplyToCell(sheet.Cells["A39"]));
                Assert.IsTrue(((ExcelConditionalFormattingStdDevGroup)above3).ShouldApplyToCell(sheet.Cells["A40"]));

                Assert.IsFalse(((ExcelConditionalFormattingStdDevGroup)above1).ShouldApplyToCell(sheet.Cells["A37"]));
                Assert.IsFalse(((ExcelConditionalFormattingStdDevGroup)above2).ShouldApplyToCell(sheet.Cells["A38"]));
                Assert.IsFalse(((ExcelConditionalFormattingStdDevGroup)above3).ShouldApplyToCell(sheet.Cells["A39"]));

                Assert.IsTrue(((ExcelConditionalFormattingStdDevGroup)below1).ShouldApplyToCell(sheet.Cells["A3"]));
                Assert.IsTrue(((ExcelConditionalFormattingStdDevGroup)below2).ShouldApplyToCell(sheet.Cells["A2"]));
                Assert.IsTrue(((ExcelConditionalFormattingStdDevGroup)below3).ShouldApplyToCell(sheet.Cells["A1"]));

                //SaveAndCleanup(pck);
            }
        }
    }
}

using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Style;
using System.Drawing;


namespace EPPlusTest.ConditionalFormatting
{
    [TestClass]
    public class CF_Between : TestBase
    {
        [TestMethod]
        public void CF_BetweenShouldApplyMissmatchWorks()
        {
            using (var pck = OpenPackage("betweenConditionalFormattingMissMatch.xlsx", true))
            {
                var missMatch = pck.Workbook.Worksheets.Add("missMatchSheet");

                missMatch.Cells["A1"].Value = "Ape";
                missMatch.Cells["A2"].Value = "Baboon";
                missMatch.Cells["A3"].Value = "C";
                missMatch.Cells["A4"].Value = "Capuchin";
                missMatch.Cells["A5"].Value = "Drill";

                missMatch.Cells["A6"].Value = "Emperor Tamarin";
                missMatch.Cells["A7"].Value = 2;
                missMatch.Cells["A8"].Value = 30;
                missMatch.Cells["A9"].Value = 20;
                missMatch.Cells["A10"].Value = 18;
                missMatch.Cells["A11"].Value = 19;
                missMatch.Cells["A12"].Value = 567;
                missMatch.Cells["A13"].Value = 5677777;
                missMatch.Cells["A14"].Value = 5677777;

                var between = missMatch.Cells["A1:A14"].ConditionalFormatting.AddBetween();

                between.Style.Fill.PatternType = ExcelFillStyle.Solid;
                between.Style.Fill.BackgroundColor.Color = Color.AliceBlue;

                between.Formula = "\"C\"";
                between.Formula2 = "20";

                var betweenReal = (ExcelConditionalFormattingBetween)between;

                //Numerical values above or equal to formula2 should == true. strings <= formula should == true
                //That's how excel seems to resolve it.
                Assert.IsTrue(betweenReal.ShouldApplyToCell(missMatch.Cells["A1"]));
                Assert.IsTrue(betweenReal.ShouldApplyToCell(missMatch.Cells["A2"]));
                Assert.IsTrue(betweenReal.ShouldApplyToCell(missMatch.Cells["A3"]));

                Assert.IsFalse(betweenReal.ShouldApplyToCell(missMatch.Cells["A4"]));
                Assert.IsFalse(betweenReal.ShouldApplyToCell(missMatch.Cells["A5"]));
                Assert.IsFalse(betweenReal.ShouldApplyToCell(missMatch.Cells["A6"]));
                Assert.IsFalse(betweenReal.ShouldApplyToCell(missMatch.Cells["A7"]));

                Assert.IsTrue(betweenReal.ShouldApplyToCell(missMatch.Cells["A8"]));
                Assert.IsTrue(betweenReal.ShouldApplyToCell(missMatch.Cells["A9"]));
                Assert.IsFalse(betweenReal.ShouldApplyToCell(missMatch.Cells["A10"]));
                Assert.IsFalse(betweenReal.ShouldApplyToCell(missMatch.Cells["A11"]));

                Assert.IsTrue(betweenReal.ShouldApplyToCell(missMatch.Cells["A12"]));
                Assert.IsTrue(betweenReal.ShouldApplyToCell(missMatch.Cells["A13"]));
                Assert.IsTrue(betweenReal.ShouldApplyToCell(missMatch.Cells["A14"]));

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void CF_BetweenShouldApplyNumeric()
        {
            using (var pck = OpenPackage("CF_NumericBetween.xlsx", true))
            {
                var numBetween = pck.Workbook.Worksheets.Add("numericBetween");
                numBetween.Cells["A7"].Value = "2";
                numBetween.Cells["A8"].Value = "30";
                numBetween.Cells["A9"].Value = "20";
                numBetween.Cells["A10"].Value = "18";
                numBetween.Cells["A11"].Value = "19";
                numBetween.Cells["A12"].Value = "567";
                numBetween.Cells["A13"].Value = "5677777";
                numBetween.Cells["A14"].Value = "5677777";

                var between = numBetween.Cells["A1:A14"].ConditionalFormatting.AddBetween();

                between.Formula = "5";
                between.Formula2 = "25";

                var betweenCast = (ExcelConditionalFormattingBetween)between;

                Assert.IsFalse(betweenCast.ShouldApplyToCell(numBetween.Cells["A7"]));
                Assert.IsFalse(betweenCast.ShouldApplyToCell(numBetween.Cells["A8"]));
                Assert.IsTrue(betweenCast.ShouldApplyToCell(numBetween.Cells["A9"]));
                Assert.IsTrue(betweenCast.ShouldApplyToCell(numBetween.Cells["A10"]));
                Assert.IsTrue(betweenCast.ShouldApplyToCell(numBetween.Cells["A11"]));

                Assert.IsFalse(betweenCast.ShouldApplyToCell(numBetween.Cells["A12"]));
                Assert.IsFalse(betweenCast.ShouldApplyToCell(numBetween.Cells["A13"]));
                Assert.IsFalse(betweenCast.ShouldApplyToCell(numBetween.Cells["A14"]));
            }
        }

        [TestMethod]
        public void CF_BetweenShouldApplyStrings()
        {
            using (var pck = OpenPackage("CF_StringBetween.xlsx", true))
            {
                var strBetween = pck.Workbook.Worksheets.Add("stringBetween");
                strBetween.Cells["A1"].Value = "Abc";
                strBetween.Cells["A2"].Value = "Def";
                strBetween.Cells["A3"].Value = "Ghi";
                strBetween.Cells["A4"].Value = "jkl";
                strBetween.Cells["A5"].Value = "mno";
                strBetween.Cells["A6"].Value = "pqr";
                strBetween.Cells["A7"].Value = "stv";
                strBetween.Cells["A8"].Value = "wxyz";

                var between = strBetween.Cells["A1:A14"].ConditionalFormatting.AddBetween();

                between.Formula = "\"Def\"";
                between.Formula2 = "\"mno\"";

                var betweenCast = (ExcelConditionalFormattingBetween)between;

                Assert.IsFalse(betweenCast.ShouldApplyToCell(strBetween.Cells["A1"]));
                Assert.IsTrue(betweenCast.ShouldApplyToCell(strBetween.Cells["A2"]));
                Assert.IsTrue(betweenCast.ShouldApplyToCell(strBetween.Cells["A3"]));
                Assert.IsTrue(betweenCast.ShouldApplyToCell(strBetween.Cells["A4"]));
                Assert.IsTrue(betweenCast.ShouldApplyToCell(strBetween.Cells["A5"]));

                Assert.IsFalse(betweenCast.ShouldApplyToCell(strBetween.Cells["A6"]));
                Assert.IsFalse(betweenCast.ShouldApplyToCell(strBetween.Cells["A7"]));
                Assert.IsFalse(betweenCast.ShouldApplyToCell(strBetween.Cells["A8"]));
            }
        }
    }
}

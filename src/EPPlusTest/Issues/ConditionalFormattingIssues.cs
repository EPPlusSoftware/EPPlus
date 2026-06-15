using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Globalization;
using System.Threading;

namespace EPPlusTest.Issues
{
    [TestClass]
    public class ConditionalFormattingIssues : TestBase
    {
        [TestMethod]
        public void DatabarNegativesAndFormulasTest()
        {
            var package = OpenTemplatePackage("i1244Databars.xlsm");
            Assert.IsNotNull(package.Workbook);

            SaveAndCleanup(package);
        }

        //Contains blanks when address ref.
        [TestMethod]
        public void ContainsBlanksTest()
        {
            using (var p = OpenTemplatePackage("i1254.xlsx"))
            {

                var sheet = p.Workbook.Worksheets[0];

                sheet.Cells["Z1"].Value = 1;

                sheet.Calculate();

                SaveAndCleanup(p);
            }
        }
        public void Test1_Input_ExpectedOutput()
        {
            var currentCulture = Thread.CurrentThread.CurrentCulture;
            // if this is InvariantCulture, everything works fine:
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("de-DE");

            using var package = OpenPackage("i2054.xlsx", delete: true);

            var worksheet = package.Workbook.Worksheets.Add("test");
            int fromRow = 1;
            int toRow = 10;
            var range = worksheet.Cells[fromRow, 1, toRow, 1];

            for (int i = fromRow; i <= toRow; i++)
            {
                worksheet.Cells[i, 1].Value = i;
            }

            var iconSet =
                range.ConditionalFormatting.AddThreeIconSet(
                    eExcelconditionalFormatting3IconsSetType.Symbols2);

            // icons are counted bottom up, when compared to Excel UI, so Icon3 is the topmost one:
            iconSet.Icon1.Type = eExcelConditionalFormattingValueObjectType.Num;
            iconSet.Icon1.Value = 0;

            iconSet.Icon2.Type = eExcelConditionalFormattingValueObjectType.Num;
            iconSet.Icon2.Value = 1.5; // this is the problem: get's written to the XML as 1,5

            iconSet.Icon3.Type = eExcelConditionalFormattingValueObjectType.Num;
            iconSet.Icon3.Value = 3;
            iconSet.Icon3.GreaterThanOrEqualTo = false;

            iconSet.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.BlackCircle;
            iconSet.Icon2.CustomIcon = eExcelconditionalFormattingCustomIcon.BlackCircle;
            iconSet.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.GoldStar;

            SaveAndCleanup(package);
            Thread.CurrentThread.CurrentCulture = currentCulture;
        }

        [TestMethod]
        public void s1025()
        {
            using (var package = OpenTemplatePackage("s1025.xlsx"))
            {
                ExcelWorkbook wb = package.Workbook;
                ExcelWorksheet ws = wb.Worksheets[0];

                IExcelConditionalFormattingBetween condGreen = ws.ConditionalFormatting.AddBetween(ws.Cells[5, 2, 25, 2]);
                condGreen.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                condGreen.Style.Fill.BackgroundColor.Color = ColorTranslator.FromHtml("#A9D08E");
                condGreen.Formula = ws.Cells[6, 7].FullAddressAbsolute.ToString();
                condGreen.Formula2 = ws.Cells[6, 8].FullAddressAbsolute.ToString();

                IExcelConditionalFormattingBetween condYellow = ws.ConditionalFormatting.AddBetween(ws.Cells[5, 2, 25, 2]);
                condYellow.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                condYellow.Style.Fill.BackgroundColor.Color = ColorTranslator.FromHtml("#FFE699");
                condYellow.Formula = ws.Cells[7, 7].FullAddressAbsolute.ToString();
                condYellow.Formula2 = ws.Cells[7, 8].FullAddressAbsolute.ToString();

                IExcelConditionalFormattingGreaterThan condRed = ws.ConditionalFormatting.AddGreaterThan(ws.Cells[5, 2, 25, 2]);
                condRed.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                condRed.Style.Fill.BackgroundColor.Color = ColorTranslator.FromHtml("#FF7979");
                condRed.Formula = ws.Cells[9, 7].FullAddressAbsolute.ToString();

                Assert.AreEqual(6, condRed.Priority);

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void VerifyReadWritePriority()
        {
            using (var p = OpenPackage("CFReadWritePriority.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("CFReadWritePriorityWs");

                ws.Cells["A1:C10"].Formula = "ROW()+COLUMN()";

                ws.Calculate();

                var cf1 = ws.Cells["A1:A10"].ConditionalFormatting.AddAboveAverage();
                var cf2 = ws.Cells["A1:A10"].ConditionalFormatting.AddAboveAverage();

                cf1.Style.Fill.BackgroundColor.SetColor(Color.RosyBrown);

                cf2.Style.Fill.BackgroundColor.SetColor(Color.DarkGreen);

                var cfB1 = ws.Cells["B1:B10"].ConditionalFormatting.AddBetween();
                cfB1.Formula = "4";
                cfB1.Formula2 = "14";

                var cfB2 = ws.Cells["B1:B10"].ConditionalFormatting.AddBetween();

                cfB2.Formula = "14";
                cfB2.Formula2 = "35";

                cfB1.Style.Fill.BackgroundColor.SetColor(Color.BlueViolet);
                cfB2.Style.Fill.BackgroundColor.SetColor(Color.Chartreuse);

                cfB1.Priority = 1;

                SaveAndCleanup(p);
            }

            using (var p = OpenPackage("CFReadWritePriority.xlsx", false))
            {
                var ws = p.Workbook.Worksheets[0];

                var cfText = ws.Cells["C3:C5"].ConditionalFormatting.AddContainsText();

                cfText.Formula = "";
                cfText.Style.Fill.BackgroundColor.Color = Color.Red;

                Assert.AreEqual(5, cfText.Priority);

                var outFile = GetOutputFile("", "CFSamePriorityResave.xlsx");

                p.SaveAs(outFile);
                //SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void i2381()
        {
            using (var package = OpenTemplatePackage("Cabinet_template_test_clean.xlsx"))
            {
                var copyCount = package.Workbook.Worksheets.Count;
                var worksheets = package.Workbook.Worksheets;
                var ws = worksheets[0];
                var range1 = ws.Cells["A6:AL23"];

                var cfS = range1.ConditionalFormatting.GetConditionalFormattings();
                var cfStyle = cfS[0].Style;
                var borderBottom = cfS[0].Style.Border.HasValue;
                //for (int i = 0; i < copyCount; i++)
                //{
                //    package.Workbook.Worksheets.Copy(
                //    worksheets[i].Name, $"{worksheets[i].Name}_{i}");
                //}

                //ws.Workbook.Styles.Dxfs = null;

                var file = GetOutputFile("", "Cabinet_template_test_clean_Output.xlsx");
                package.SaveAs(file);
                //SaveAndCleanup(package);
            }

            using (var package = OpenPackage("Cabinet_template_test_clean_Output.xlsx"))
            {
                var copyCount = package.Workbook.Worksheets.Count;
                var worksheets = package.Workbook.Worksheets;
                var ws = worksheets[0];
                var range1 = ws.Cells["A6:AL23"];

                var cfS = range1.ConditionalFormatting.GetConditionalFormattings();
                var cfStyle = cfS[0].Style;
                var borderBottom = cfS[0].Style.Border.HasValue;

                var style = ws.Workbook.Styles.Dxfs[128];

                var test = "why";

                var file = GetOutputFile("", "Cabinet_template_test_clean_OutputReRead.xlsx");
                package.SaveAs(file);
            }
        }

        [TestMethod]
        public void i2381OnlyTheBug()
        {
            using (var package = OpenTemplatePackage("mergedCellCFBorder.xlsx"))
            {
                var ws = package.Workbook.Worksheets[0];
                var cf = ws.ConditionalFormatting[1];
                var border = cf.Style.Border;
                SaveAndCleanup(package);
            }

            using (var package = OpenPackage("mergedCellCFBorder.xlsx"))
            {
                var ws = package.Workbook.Worksheets[0];
                var cf = ws.ConditionalFormatting[1];
                var border = cf.Style.Border;
                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void i2381Excel()
        {
            using (var package = OpenTemplatePackage("Cabinet_template_test_clean_excelSaved.xlsx"))
            {
                var copyCount = package.Workbook.Worksheets.Count;
                var worksheets = package.Workbook.Worksheets;
                var ws = worksheets[0];
                var range1 = ws.Cells["A6:AL23"];

                var cfS = range1.ConditionalFormatting.GetConditionalFormattings();
                var cfStyle = cfS[0].Style;
                var borderBottom = cfS[0].Style.Border.HasValue;
                //for (int i = 0; i < copyCount; i++)
                //{
                //    package.Workbook.Worksheets.Copy(
                //    worksheets[i].Name, $"{worksheets[i].Name}_{i}");
                //}

                var file = GetOutputFile("", "Cabinet_template_test_clean_ExcelOutput.xlsx");
                package.SaveAs(file);
                //SaveAndCleanup(package);
            }
        }
    }
}

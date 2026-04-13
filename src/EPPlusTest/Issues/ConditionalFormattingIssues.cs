using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Compatibility.System.Drawing;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Style;
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
            using var package = OpenTemplatePackage("s1025.xlsx"); //attached template, ExampleWB.xlsx
            ExcelWorksheet ws = package.Workbook.Worksheets[0];

            IExcelConditionalFormattingBetween condGreen = ws.ConditionalFormatting.AddBetween(ws.Cells[5, 2, 25, 2]);
            condGreen.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            condGreen.Style.Fill.BackgroundColor.Color = ColorTranslator.FromHtml("#A9D08E");
            condGreen.Formula = ws.Cells[6, 8].FullAddressAbsolute.ToString();
            condGreen.Formula2 = ws.Cells[6, 9].FullAddressAbsolute.ToString();
            condGreen.Priority = 104;

            IExcelConditionalFormattingBetween condYellow = ws.ConditionalFormatting.AddBetween(ws.Cells[5, 2, 25, 2]);
            condYellow.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            condYellow.Style.Fill.BackgroundColor.Color = ColorTranslator.FromHtml("#FFE699");
            condYellow.Formula = ws.Cells[7, 8].FullAddressAbsolute.ToString();
            condYellow.Formula2 = ws.Cells[7, 9].FullAddressAbsolute.ToString();
            condYellow.Priority = 105;

            IExcelConditionalFormattingGreaterThan condRed = ws.ConditionalFormatting.AddGreaterThan(ws.Cells[5, 2, 25, 2]);
            condRed.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            condRed.Style.Fill.BackgroundColor.Color = ColorTranslator.FromHtml("#FF7979");
            condRed.Formula = ws.Cells[9, 8].FullAddressAbsolute.ToString();
            condRed.Priority = 106;

            SaveAndCleanup(package);
        }
    }
}

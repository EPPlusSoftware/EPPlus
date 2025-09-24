using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
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
    }
}

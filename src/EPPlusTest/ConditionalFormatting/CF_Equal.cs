using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Style;
using System.Drawing;

namespace EPPlusTest.ConditionalFormatting
{
    [TestClass]
    public class CF_Equal : TestBase
    {
        [TestMethod]
        public void CF_ShoulApply()
        {
            using (var pck = OpenPackage("CF_Equal.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("Equal");
                for(int i = 1; i <= 10; i++) 
                {
                    sheet.Cells[i, 2].Value = i * 5;
                }

                var equal = sheet.Cells["B1:B10"].ConditionalFormatting.AddEqual();

                equal.Formula = "ROW()*5";

                var equalCast = (ExcelConditionalFormattingEqual)equal;

                for (int i = 1; i <= 10; i++)
                {
                    Assert.IsTrue(equalCast.ShouldApplyToCell(sheet.Cells[i,2]));
                }
            }
        }
    }
}

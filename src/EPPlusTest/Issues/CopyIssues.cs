using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Issues
{
    [TestClass]
    public class CopyIssues : TestBase
    {
        [TestMethod]
        public void Issue1332()
        {
            // the error in this issue was that the intersect operator (SPACE)
            // was replaced with "isc" when a formulas was copied to a new destination
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "SUBTOTAL(109, _DATA _Quantity)";
            sheet.Cells["A1"].Copy(sheet.Cells["B1"]);
            Assert.AreEqual("SUBTOTAL(109,_DATA _Quantity)", sheet.Cells["B1"].Formula);
        }
        [TestMethod]
        public void s651()
        {
            using (var p = OpenTemplatePackage("s651.xlsx"))
            {
                using (var p2 = OpenPackage("s651-save.xlsx", true))
                {
                    ExcelWorksheet wOutsheet = p2.Workbook.Worksheets.Add("MergeSheet");
                    wOutsheet.Cells.Style.Font.Name = "ＭＳ Ｐゴシック";
                    wOutsheet.Cells.Style.Font.Size = 9;

                    var wAnswerSheet = p.Workbook.Worksheets["Answer Sheet"];

                    ExcelRange wAnswerCopyHeaderRange = wAnswerSheet.Cells[1, 1, 42, 39];
                    ExcelRange wOutHeaderRange = wOutsheet.Cells[1, 1, 42, 39];
                    wAnswerCopyHeaderRange.Copy(wOutHeaderRange);

                    p2.Workbook.Worksheets["MergeSheet"].Name = "Data Sheet";
                    //p2.Workbook.Calculate();
                    SaveAndCleanup(p2);
                }
            }
		}
		[TestMethod]
		public void s651_2()
		{
			using (var p = OpenTemplatePackage("s651-2.xlsx"))
			{
				using (var p2 = OpenPackage("s651-2-save.xlsx", true))
				{
					ExcelWorksheet wOutsheet = p2.Workbook.Worksheets.Add("MergeSheet");
					var ws = p.Workbook.Worksheets[0];

					ws.Cells["A1:B8"].Copy(wOutsheet.Cells["A1"]);
					wOutsheet.Calculate();
					SaveAndCleanup(p2);
				}
			}
		}
        [TestMethod]
        public void i1645()
        {
            using (var package = OpenTemplatePackage("i1645.xlsx"))
            {
                var syncSht = package.Workbook.Worksheets["syncSht"];
                var snapSht = package.Workbook.Worksheets["snapSht"];
                var address = "B7:K16";
                snapSht.Cells[address].Copy(syncSht.Cells[address]);
                SaveAndCleanup(package);
            }
        }
    }
}

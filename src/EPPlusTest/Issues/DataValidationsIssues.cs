using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Issues
{
	[TestClass]
	public class DataValidationsIssues : TestBase
	{
		[TestMethod]
		public void DatabarNegativesAndFormulasTest()
		{
			using (var package = OpenTemplatePackage("s621.xlsx"))
			{
				var Sheet1 = package.Workbook.Worksheets[$"Sheet1"];
				var Sheet2 = package.Workbook.Worksheets[$"Sheet2"];
				Sheet1.InsertColumn(1, 2);


				var startCell = Sheet1.Cells[4, 1];
				var endCell = Sheet1.Cells[6, 1];
				var fullRange = $"{startCell.AddressAbsolute}:{endCell.AddressAbsolute}";

				var from = Sheet2.Cells[2, 3].AddressAbsolute;
				var to = Sheet2.Cells[Sheet2.Dimension.End.Row, 3].AddressAbsolute;


				var wValidationList = Sheet1.DataValidations.AddListValidation(fullRange);
				wValidationList.Formula.ExcelFormula = "Sheet2" + "!" +
					from + ":" + to;


				var validations2 = Sheet1.DataValidations.ToList();

				SaveAndCleanup(package);
			}
		}
		[TestMethod]
		public void s798()
		{
			var template = "s798.xlsx";
			string dv = "";
			using (var p1 = OpenTemplatePackage(template))
			{
				var ws = p1.Workbook.Worksheets[1];
				dv = ws.DataValidations[2].As.ListValidation.Formula.Values[3];
				SaveAndCleanup(p1);
			}
			using (var p2 = OpenPackage(template))
			{
				var ws = p2.Workbook.Worksheets[1];
				Assert.AreEqual(dv, ws.DataValidations[2].As.ListValidation.Formula.Values[3]);
				SaveWorkbook("s798-saved.xlsx", p2);
			}
		}

		[TestMethod]
		//Removing and adding data validation after removing and adding rows.
		public void i2154()
		{
			using (var pck = OpenPackage("testValidations2154.xlsx", true))
			{
				// Add a worksheet
				var sheet1 = pck.Workbook.Worksheets.Add("Sheet1");

				// Next, add a data validation list to the sheet
				var dv = sheet1.Cells["A1:A10"].DataValidation.AddListDataValidation();
				dv.Formula.Values.Add("Option A");
				dv.Formula.Values.Add("Option B");

				// Delete all except the first row
				sheet1.DeleteRow(2, 9);

				// Remove the data validation
				sheet1.DataValidations.Remove(dv);
				Assert.AreEqual(0, sheet1.DataValidations.Count);

				// Now re-add the data validation
				// THIS SHOULDN'T THROW AN EXCEPTION
				sheet1.Cells["A1:A10"].DataValidation.AddListDataValidation();
				dv.Formula.Values.Add("Option C");
				dv.Formula.Values.Add("Option D");
				SaveAndCleanup(pck);
			}
		}

		[TestMethod]
		public void i2571()
		{
            using (var pck = new ExcelPackage())
            {
                // GIVEN a data validation over two areas, D16:D20 and D28
                var sheet1 = pck.Workbook.Worksheets.Add("Sheet1");
                var dv = sheet1.Cells["D16:D20 D28"].DataValidation.AddListDataValidation();
                dv.Formula.Values.Add("Option A");
                dv.Formula.Values.Add("Option B");

                // WHEN 4 rows are inserted between the two areas
                sheet1.InsertRow(24, 4);

                // THEN the second area is shifted down by 4 rows
                Assert.AreEqual("D16:D20 D32", dv.Address.Address);
            }
        }
	}
}

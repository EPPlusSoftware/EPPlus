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
    }
}

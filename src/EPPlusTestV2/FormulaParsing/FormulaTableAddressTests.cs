using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class FormulaTableAddressTests : TestBase
    {
        private static ExcelPackage _package;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _package = OpenPackage("FormulaTableAddress.xlsx", true);
		}

		[ClassCleanup]
        public static void Cleanup()
        {
			SaveAndCleanup(_package);
		}
		[TestMethod]
        public void CalculateFormulaWithNameReferencingTableRowColumn()
        {
			var ws = _package.Workbook.Worksheets.Add("Data1");
			LoadTestdata(ws);
			ws.Cells["E1"].Value = "Formula";
			var tbl = ws.Tables.Add(ws.Cells["A1:E101"], "Table1");
			tbl.Columns[4].CalculatedColumnFormula = "TableFormulaName";
			_package.Workbook.Names.AddFormula("TableFormulaName", "Table1[[#This Row],[NumValue]]");
			ws.Cells["F2:F101"].CreateArrayFormula("TableFormulaName");
			ws.Calculate();
        }
		[TestMethod]
		public void ValidateInsertKeepTableFormulaIntact()
		{
			var ws = _package.Workbook.Worksheets.Add("Data2");
			LoadTestdata(ws);
			ws.Cells["E1"].Value = "Formula";
			var tbl = ws.Tables.Add(ws.Cells["A1:E101"], "Table2");
			tbl.Columns[4].CalculatedColumnFormula = "Table1[[#This Row],[NumValue]]+1";
			tbl.InsertRow(5, 2);
			ws.Calculate();
			Assert.AreEqual("Table1[[#This Row],[NumValue]]+1", ws.Cells["E5"].Formula);
		}
	}
}

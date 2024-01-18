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
        private static ExcelWorksheet _ws;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _package = OpenPackage("FormulaTableAddress.xlsx", true);
            _ws = _package.Workbook.Worksheets.Add("Data");
            LoadTestdata(_ws);
            _ws.Cells["E1"].Value = "Formula";
            var tbl = _ws.Tables.Add(_ws.Cells["A1:E101"], "Table1");
            tbl.Columns[4].CalculatedColumnFormula = "TableFormulaName";            
            _package.Workbook.Names.AddFormula("TableFormulaName", "Table1[[#This Row],[NumValue]]");
            _ws.Cells["F2:F101"].CreateArrayFormula("TableFormulaName");            
        }

        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_package);
        }
        [TestMethod]
        public void CalculateFormulaWithNameReferencingTableRowColumn()
        {
            _ws.Calculate();
        }
    }
}

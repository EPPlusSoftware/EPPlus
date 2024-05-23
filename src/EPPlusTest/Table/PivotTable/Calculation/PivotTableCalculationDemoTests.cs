using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Table.PivotTable.Calculation;
using OfficeOpenXml.Table.PivotTable.Calculation.Functions;
using System.Collections.Generic;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
namespace EPPlusTest.Table.PivotTable.Calculation
{
    [TestClass]
    public class PivotTableCalculationDemoTests : TestBase
    {
        static ExcelPackage _package;
        static ExcelWorksheet _ptWs;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _package = OpenTemplatePackage("PivotTableWorkbook.xlsx");
            _ptWs = _package.Workbook.Worksheets["PivotTable"];
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            _package.Dispose();
        }
        [TestMethod]
        public void DemoTest()
        {
            var pt = _ptWs.PivotTables[0];
            pt.Calculate();

            _ptWs.Calculate();
            var i11 = _ptWs.Cells["I11"].Value;


            var canadaValue = pt.CalculatedData.
                SelectField("Country", "Canada").
                GetValue("OrderValue");
            var canadaTrueValueTrue = pt.CalculatedData.
                SelectField("Country", "Canada").
                SelectField("IsValid", true).
                GetValue("OrderValue");
            var brazilTrueValueTrue = pt.CalculatedData.
                SelectField("Country","Brazil").
                SelectField("IsValid",true).
                GetValue("Rating");
            var grandTotal = pt.CalculatedData.GetValue("OrderValue");
        }
    }
}

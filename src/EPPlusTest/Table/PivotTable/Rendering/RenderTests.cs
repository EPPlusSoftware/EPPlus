using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;
namespace EPPlusTest.Table.PivotTable.Rendering
{
    [TestClass]
    public class PivotTableRenderTests : TestBase
    {
        //static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            //InitBase();
            //_pck = OpenPackage("PivotTableAutoSort.xlsx", true);
            //var ws = _pck.Workbook.Worksheets.Add("Data1");
            //var r = LoadItemData(ws);
            //ws.Tables.Add(r, "Table1");
            //ws = _pck.Workbook.Worksheets.Add("Data2");
            //r = LoadItemData(ws);
            //ws.Tables.Add(r, "Table2");
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            //SaveAndCleanup(_pck);
        }
    }
}

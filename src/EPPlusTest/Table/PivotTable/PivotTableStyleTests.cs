using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Table.PivotTable
{
    [TestClass]
    public class PivotTableStyleTests : TestBase
    {
        static ExcelPackage _pck;
        static ExcelWorksheet _ws;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenPackage("PivotTableStyle.xlsx", true);
            _ws = _pck.Workbook.Worksheets.Add("Data1");
            LoadItemData(_ws);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void SetPivotAllStyle()
        {
            var ws = _pck.Workbook.Worksheets.Add("StyleAll");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], _ws.Cells[_ws.Dimension.Address], "PivotTable1");
            pt.RowFields.Add(pt.Fields[0]);
            pt.DataFields.Add(pt.Fields[3]);
            var style = pt.PivotAreaStyles.Add(ePivotAreaType.All);
            style.DataOnly = false;
            style.LabelOnly = false;
            style.GrandRow = true;
            style.GrandColumn = true;
            style.Style.Font.Color.SetColor(OfficeOpenXml.Drawing.eThemeSchemeColor.Accent1);            
        }
    }
}


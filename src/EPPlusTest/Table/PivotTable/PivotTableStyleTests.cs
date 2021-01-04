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
            pt.Styling.All.Style.Font.Name="Bauhaus 93";
            pt.Styling.GrandRowData.Style.Font.Italic = true;
            pt.Styling.GrandRowData.Style.Font.Bold = false;
            //pt.Styling.Data.Style.Font.Bold = true;
            //pt.Styling.Origin.Style.Font.Color.SetColor(OfficeOpenXml.Drawing.eThemeSchemeColor.Accent4);
            //pt.Styling.ColumnHeaders.Style.Font.Color.SetColor(OfficeOpenXml.Drawing.eThemeSchemeColor.Accent5);
            //pt.Styling.GrandColumnHeaders.Style.Font.Color.SetColor(OfficeOpenXml.Drawing.eThemeSchemeColor.Accent5);
        }
    }
}


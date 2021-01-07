using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Drawing;
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
            var pt = ws.PivotTables.Add(ws.Cells["A3"], _ws.Cells[_ws.Dimension.Address], "PivotTable1");
            pt.RowFields.Add(pt.Fields[0]);
            pt.ColumnFields.Add(pt.Fields[1]);
            pt.DataFields.Add(pt.Fields[3]);
            pt.DataFields.Add(pt.Fields[2]);
            pt.PageFields.Add(pt.Fields[4]);
            pt.Styling.All.Style.Font.Name="Bauhaus 93";
            //pt.Styling.GrandRowData.Style.Font.Italic = true;
            //pt.Styling.GrandRowData.Style.Font.Bold = false;
            //pt.Styling.GrandColumnData.Style.Font.Underline = OfficeOpenXml.Style.ExcelUnderLineType.Single;
            pt.Styling.Labels.Style.Font.Color.SetColor(Color.Green);
            pt.Styling.Data.Style.Font.Color.SetColor(Color.Blue);
            //pt.Styling.GrandRowData.Style.Font.Color.SetColor(Color.Green);
            //pt.Styling.GrandColumnData.Style.Font.Color.SetColor(Color.Yellow);
            //pt.Styling.ColumnLabels.Style.Font.Underline = OfficeOpenXml.Style.ExcelUnderLineType.Double;
            //pt.Styling.ColumnLabels.Style.Font.Color.SetColor(Color.Red);

            pt.Styling.Origin.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            pt.Styling.Origin.Style.Fill.BackgroundColor.SetColor(Color.Red);

            //pt.Styling.ColumnLabels.FieldIndex = 1;
        }
    }
}


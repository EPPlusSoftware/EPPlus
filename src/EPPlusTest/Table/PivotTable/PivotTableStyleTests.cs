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
            var s=pt.Styling.Areas.AddWholeTable();
            s.Style.Font.Name = "Bauhaus 93";

            //s1.References.Add(pt.Fields[0]);
            //s1.LabelOnly = true;
            //s1.Style.Fill.Style = OfficeOpenXml.Style.eDxfFillStyle.PatternFill;
            //s1.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            //s1.Style.Fill.BackgroundColor.SetColor(Color.Black);
            //s1.FieldIndex = 0;
            //s1.FieldPosition = 0;
            //pt.Styling.GrandRowData.Style.Font.Italic = true;
            //pt.Styling.GrandRowData.Style.Font.Bold = false;
            //pt.Styling.GrandColumnData.Style.Font.Underline = OfficeOpenXml.Style.ExcelUnderLineType.Single;
            s = pt.Styling.Areas.AddWholeTable(true);
            s.Style.Font.Color.SetColor(Color.Green);
            s = pt.Styling.Areas.AddWholeTable(false, true);
            s.Style.Font.Color.SetColor(Color.Blue);
            //pt.Styling.GrandRowData.Style.Font.Color.SetColor(Color.Green);
            //pt.Styling.GrandColumnData.Style.Font.Color.SetColor(Color.Yellow);
            //pt.Styling.ColumnLabels.Style.Font.Underline = OfficeOpenXml.Style.ExcelUnderLineType.Double;
            //pt.Styling.ColumnLabels.Style.Font.Color.SetColor(Color.Red);
            var s1 = pt.Styling.Areas.AddButtonField(pt.Fields[4]);
            s1.Style.Font.Color.SetColor(Color.Pink);

            var s2 = pt.Styling.Areas.AddButton(ePivotTableAxis.RowAxis);
            s1.Style.Font.Underline = OfficeOpenXml.Style.ExcelUnderLineType.DoubleAccounting;

            var s3 = pt.Styling.Areas.AddButton(ePivotTableAxis.ColumnAxis);
            s3.Style.Font.Italic=true;

            var s4 = pt.Styling.Areas.AddButton(ePivotTableAxis.PageAxis);
            s4.Style.Font.Color.SetColor(Color.ForestGreen);

            var s5 = pt.Styling.Areas.AddButtonField(pt.Fields[4]);
            s4.Style.Font.Color.SetColor(Color.AliceBlue);

            //Top Left cells 
            var styleTopLeft = pt.Styling.Areas.AddTopStart();
            styleTopLeft.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            styleTopLeft.Style.Fill.BackgroundColor.SetColor(Color.Red);

            //Top right cells 
            var styleTopRight = pt.Styling.Areas.AddTopEnd();
            styleTopRight.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            styleTopRight.Style.Fill.BackgroundColor.SetColor(Color.Green);

            var styleTopRight2 = pt.Styling.Areas.AddTopEnd(1);
            styleTopRight2.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            styleTopRight2.Style.Fill.BackgroundColor.SetColor(Color.Yellow);


            //pt.Styling.ColumnLabels.FieldIndex = 1;
        }
    }
}


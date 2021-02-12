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
        internal ExcelPivotTable CreatePivotTable(ExcelWorksheet ws)
        {
            var pt = ws.PivotTables.Add(ws.Cells["A3"], _ws.Cells[_ws.Dimension.Address], "PivotTable1");
            pt.RowFields.Add(pt.Fields[0]);
            pt.ColumnFields.Add(pt.Fields[1]);
            pt.DataFields.Add(pt.Fields[3]);
            pt.DataFields.Add(pt.Fields[2]);
            pt.PageFields.Add(pt.Fields[4]);
            return pt;
        }
        [TestMethod]
        public void SetPivotAllStyle()
        {
            var ws = _pck.Workbook.Worksheets.Add("StyleAll");
            var pt = CreatePivotTable(ws);
            var s=pt.Styling.Areas.AddWholeTable();
            s.Style.Font.Name = "Bauhaus 93";

            Assert.IsTrue(s.Style.HasValue);
            Assert.AreEqual("Bauhaus 93", s.Style.Font.Name);
        }
        [TestMethod]
        public void AddPivotLabels()
        {
            var ws = _pck.Workbook.Worksheets.Add("StyleAllLabels");
            var pt = CreatePivotTable(ws);

            var s = pt.Styling.Areas.AddAllLabels();
            s.Style.Font.Color.SetColor(Color.Green);
        }
        [TestMethod]
        public void AddPivotLabelPageField()
        {
            var ws = _pck.Workbook.Worksheets.Add("StylePageFieldLabel");
            var pt = CreatePivotTable(ws);

            var s = pt.Styling.Areas.AddLabel(pt.PageFields[0]);
            s.Style.Font.Color.SetColor(Color.Green);
        }
        [TestMethod]
        public void AddPivotLabelColumnField()
        {
            var ws = _pck.Workbook.Worksheets.Add("StyleColumnFieldLabel");
            var pt = CreatePivotTable(ws);

            var s = pt.Styling.Areas.AddLabel(pt.ColumnFields[0]);
            s.Style.Font.Color.SetColor(Color.Indigo);
        }
        [TestMethod]
        public void AddPivotLabelRowColumnField()
        {
            var ws = _pck.Workbook.Worksheets.Add("StyleRowFieldLabel");
            var pt = CreatePivotTable(ws);

            var s = pt.Styling.Areas.AddLabel(pt.RowFields[0]);

            s.Style.Font.Italic = true;
            s.Style.Font.Strike = true;
            s.Style.Font.Name = "Times New Roman";
        }
        [TestMethod]
        public void AddPivotDataRowColumnField()
        {
            var ws = _pck.Workbook.Worksheets.Add("StyleRowFieldData");
            var pt = CreatePivotTable(ws);

            var s = pt.Styling.Areas.AddData(pt.RowFields[0]);

            s.Style.Font.Italic = true;
            s.Style.Font.Strike = true;
            s.Style.Font.Name = "Times New Roman";
        }
        [TestMethod]
        public void AddPivotDataRowColumnTotal()
        {
            var ws = _pck.Workbook.Worksheets.Add("StyleRowFieldDataTot");
            var pt = CreatePivotTable(ws);

            var s = pt.Styling.Areas.AddDataTotal(pt.RowFields[0]);

            s.Style.Font.Italic = true;
            s.Style.Font.Strike = true;
            s.Style.Font.Name = "Times New Roman";
        }
        [TestMethod]
        public void AddPivotLabelRowColumnTotal()
        {
            var ws = _pck.Workbook.Worksheets.Add("StyleRowFieldLabelTot");
            var pt = CreatePivotTable(ws);

            var s = pt.Styling.Areas.AddLabelTotal(pt.RowFields[0]);

            s.Style.Font.Italic = true;
            s.Style.Font.Strike = true;
            s.Style.Font.Name = "Times New Roman";
        }


        [TestMethod]
        public void AddPivotData()
        {
            var ws = _pck.Workbook.Worksheets.Add("StyleAllData");
            var pt = CreatePivotTable(ws);

            var s = pt.Styling.Areas.AddAllData();
            s.Style.Font.Color.SetColor(Color.Blue);
        }

        [TestMethod]
        public void AddButtonField()
        {
            var ws = _pck.Workbook.Worksheets.Add("StyleFieldPage");
            var pt = CreatePivotTable(ws);

            var s = pt.Styling.Areas.AddButtonField(pt.Fields[4]);
            s.Style.Font.Color.SetColor(Color.Pink);
        }

        [TestMethod]
        public void AddButtonRowAxis()
        {
            var ws = _pck.Workbook.Worksheets.Add("StyleButtonRowAxis");
            var pt = CreatePivotTable(ws);

            var s = pt.Styling.Areas.AddButtonField(ePivotTableAxis.RowAxis);
            s.Style.Font.Underline = OfficeOpenXml.Style.ExcelUnderLineType.DoubleAccounting;
        }
        [TestMethod]
        public void AddButtonColumnAxis()
        {
            var ws = _pck.Workbook.Worksheets.Add("StyleButtonColumnAxis");
            var pt = CreatePivotTable(ws);

            var s3 = pt.Styling.Areas.AddButtonField(ePivotTableAxis.ColumnAxis);
            s3.Style.Font.Italic = true;
        }
        [TestMethod]
        public void AddButtonPageAxis()
        {
            var ws = _pck.Workbook.Worksheets.Add("StyleButtonPageAxis");
            var pt = CreatePivotTable(ws);

            var s4 = pt.Styling.Areas.AddButtonField(ePivotTableAxis.PageAxis);
            s4.Style.Font.Color.SetColor(Color.ForestGreen);
        }


        [TestMethod]
        public void AddTopStart()
        {
            var ws = _pck.Workbook.Worksheets.Add("StyleTopStart");
            var pt = CreatePivotTable(ws);

            //Top Left cells 
            var styleTopLeft = pt.Styling.Areas.AddTopStart();
            styleTopLeft.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            styleTopLeft.Style.Fill.BackgroundColor.SetColor(Color.Red);
        }
        [TestMethod]
        public void AddTopStartOffset0()
        {
            var ws = _pck.Workbook.Worksheets.Add("StyleTopStartOffset0");
            var pt = CreatePivotTable(ws);

            //Top Left cells 
            var styleTopLeft = pt.Styling.Areas.AddTopStart(0);
            styleTopLeft.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            styleTopLeft.Style.Fill.BackgroundColor.SetColor(Color.Blue);
        }

        [TestMethod]
        public void AddTopEnd()
        {
            var ws = _pck.Workbook.Worksheets.Add("StyleTopEnd");
            var pt = CreatePivotTable(ws);

            var styleTopRight2 = pt.Styling.Areas.AddTopEnd();
            styleTopRight2.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            styleTopRight2.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
        }
        [TestMethod]
        public void AddTopEndOffset1()
        {
            var ws = _pck.Workbook.Worksheets.Add("StyleTopEndOffset1");
            var pt = CreatePivotTable(ws);

            var styleTopRight2 = pt.Styling.Areas.AddTopEnd(1);
            styleTopRight2.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            styleTopRight2.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
        }

    }
}


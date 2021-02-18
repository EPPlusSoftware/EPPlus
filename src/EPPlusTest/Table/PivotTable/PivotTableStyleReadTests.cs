using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Style;
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
    public class PivotTableStyleReadTests : TestBase
    {
        static ExcelPackage _pck;
        static ExcelWorksheet _ws;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenTemplatePackage("PivotTableReadStyle.xlsx");
        }
        [ClassCleanup]
        public static void Cleanup()
        {
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
        public void ReadPivotAllStyle()
        {
            var ws = TryGetWorksheet(_pck, "StyleAll");
            var pt = ws.PivotTables[0];
            Assert.AreEqual(1, pt.Styles.Count);

            var s = pt.Styles[0];

            Assert.AreEqual(ePivotAreaType.All, s.PivotAreaType);
            Assert.IsTrue(s.Style.HasValue);
            Assert.AreEqual("Bauhaus 93", s.Style.Font.Name);
        }
        [TestMethod]
        public void ReadPivotLabels()
        {
            var ws = TryGetWorksheet(_pck, "StyleAllLabels");
            var pt = ws.PivotTables[0];
            Assert.AreEqual(1, pt.Styles.Count);

            var s = pt.Styles[0];

            Assert.AreEqual(ePivotAreaType.All, s.PivotAreaType);
            Assert.IsTrue(s.LabelOnly);
            Assert.IsFalse(s.DataOnly);

            Assert.AreEqual(Color.Green.ToArgb(), s.Style.Font.Color.Color.Value.ToArgb());
        }
        [TestMethod]
        public void ReadPivotAllData()
        {
            var ws = TryGetWorksheet(_pck, "StyleAllData");
            var pt = ws.PivotTables[0];
            Assert.AreEqual(1, pt.Styles.Count);

            var s = pt.Styles[0];

            Assert.AreEqual(ePivotAreaType.All, s.PivotAreaType);
            Assert.IsTrue(s.DataOnly);
            Assert.IsFalse(s.LabelOnly);

            Assert.AreEqual(Color.Blue.ToArgb(), s.Style.Font.Color.Color.Value.ToArgb());
        }

        [TestMethod]
        public void ReadPivotLabelPageField()
        {
            var ws = TryGetWorksheet(_pck, "StylePageFieldLabel");
            var pt = ws.PivotTables[0];
            Assert.AreEqual(1, pt.Styles.Count);
            Assert.AreEqual(1, pt.Styles.Count);

            var s = pt.Styles[0];

            Assert.AreEqual(ePivotAreaType.Normal, s.PivotAreaType);
            Assert.IsTrue(s.LabelOnly);
            Assert.IsFalse(s.DataOnly);
            Assert.AreEqual(1, s.Conditions.Fields.Count);
            Assert.AreEqual(pt.PageFields[0].Index, s.Conditions.Fields[0].FieldIndex);

            Assert.AreEqual(Color.Green.ToArgb(), s.Style.Font.Color.Color.Value.ToArgb());
        }
        [TestMethod]
        public void ReadPivotLabelColumnField()
        {
            var ws = TryGetWorksheet(_pck, "StyleColumnFieldLabel");
            var pt = ws.PivotTables[0];

            var s = pt.Styles[0];

            Assert.IsTrue(s.LabelOnly);
            Assert.IsFalse(s.DataOnly);
            Assert.AreEqual(1, s.Conditions.Fields.Count);
            Assert.AreEqual(pt.ColumnFields[0].Index, s.Conditions.Fields[0].FieldIndex);

            Assert.AreEqual(Color.Indigo.ToArgb(), s.Style.Font.Color.Color.Value.ToArgb());
        }
        [TestMethod]
        public void AddPivotLabelColumnFieldSingleCell()
        {
            var ws = TryGetWorksheet(_pck, "StyleColumnFieldLabelCell");
            var pt = ws.PivotTables[0];

            var s = pt.Styles[0];

            Assert.IsFalse(pt.DataOnRows);
            Assert.IsTrue(s.LabelOnly);
            Assert.IsFalse(s.DataOnly);

            Assert.AreEqual(1, s.Conditions.DataFields.Count);
            Assert.AreEqual(1, s.Conditions.Fields.Count);
            Assert.AreEqual("Price", s.Conditions.DataFields[0].Name);

            Assert.AreEqual(1, s.Conditions.Fields.Count);
            Assert.AreEqual(2, s.Conditions.Fields[0].CacheItems.Count);
            Assert.AreEqual(0, s.Conditions.Fields[0].CacheItems[0].Index);
            Assert.AreEqual(1, s.Conditions.Fields[0].CacheItems[1].Index);

            Assert.AreEqual(Color.Indigo.ToArgb(), s.Style.Font.Color.Color.Value.ToArgb());
        }

        [TestMethod]
        public void AddPivotLabelRowColumnField()
        {
            var ws = TryGetWorksheet(_pck, "StyleRowFieldLabel");
            var pt = ws.PivotTables[0];

            var s = pt.Styles[0];

            Assert.AreEqual(ePivotAreaType.Normal, s.PivotAreaType);
            Assert.IsTrue(s.LabelOnly);
            Assert.IsFalse(s.DataOnly);
            Assert.AreEqual(1, s.Conditions.Fields.Count);
            Assert.AreEqual(pt.RowFields[0].Index, s.Conditions.Fields[0].FieldIndex);

            Assert.IsTrue(s.Style.Font.Italic.Value);
            Assert.IsTrue(s.Style.Font.Strike.Value);
            Assert.AreEqual("Times New Roman", s.Style.Font.Name);
        }
        [TestMethod]
        public void ReadPivotDataRowColumnField()
        {
            var ws = TryGetWorksheet(_pck, "StyleRowFieldData");
            var pt = ws.PivotTables[0];

            var s = pt.Styles[0];

            Assert.AreEqual(ePivotAreaType.Data, s.PivotAreaType);
            Assert.IsTrue(s.DataOnly);
            Assert.IsFalse(s.LabelOnly);
            Assert.AreEqual(1, s.Conditions.Fields.Count);
            Assert.AreEqual(pt.RowFields[0].Index, s.Conditions.Fields[0].FieldIndex);

            Assert.IsTrue(s.Style.Font.Italic.Value);
            Assert.IsTrue(s.Style.Font.Strike.Value);
            Assert.AreEqual("Times New Roman", s.Style.Font.Name);
        }
        [TestMethod]
        public void ReadPivotData()
        {
            var ws = TryGetWorksheet(_pck, "StyleData");
            var pt = ws.PivotTables[0];

            var s = pt.Styles[0];

            Assert.AreEqual(2, s.Conditions.Fields.Count);
            Assert.AreEqual(pt.Fields[0].Index, s.Conditions.Fields[0].FieldIndex);
            Assert.AreEqual(pt.Fields[1].Index, s.Conditions.Fields[1].FieldIndex);

            Assert.AreEqual(s.Style.Fill.Style, eDxfFillStyle.PatternFill);
            Assert.AreEqual(Color.Red.ToArgb(), s.Style.Fill.BackgroundColor.Color.Value.ToArgb());
            Assert.IsTrue(s.Style.Font.Italic.Value);
            Assert.IsTrue(s.Style.Font.Strike.Value);
            Assert.AreEqual("Times New Roman", s.Style.Font.Name);
        }
        [TestMethod]
        public void AddPivotDataGrandColumn()
        {
            var ws = TryGetWorksheet(_pck, "StyleDataGrandColumn");
            var pt = ws.PivotTables[0];

            var s = pt.Styles[0];

            Assert.AreEqual(2, s.Conditions.Fields.Count);
            Assert.AreEqual(pt.Fields[0].Index, s.Conditions.Fields[0].FieldIndex);
            Assert.AreEqual(pt.Fields[1].Index, s.Conditions.Fields[1].FieldIndex);
            Assert.IsTrue(s.GrandColumn);
            Assert.AreEqual(s.Style.Fill.Style, OfficeOpenXml.Style.eDxfFillStyle.PatternFill);
            Assert.AreEqual(Color.LightGray.ToArgb(), s.Style.Fill.BackgroundColor.Color.Value.ToArgb());

            Assert.AreEqual(ExcelUnderLineType.Single, s.Style.Font.Underline);
            Assert.AreEqual("Times New Roman", s.Style.Font.Name);
        }
        [TestMethod]
        public void AddPivotDataGrandRow()
        {
            var ws = TryGetWorksheet(_pck, "StyleDataGrandRow");
            var pt = ws.PivotTables[0];

            var s = pt.Styles[0];

            Assert.IsTrue(s.GrandRow);
            Assert.AreEqual(s.Style.Fill.Style, OfficeOpenXml.Style.eDxfFillStyle.PatternFill);
            Assert.AreEqual(Color.LightGray.ToArgb(), s.Style.Fill.BackgroundColor.Color.Value.ToArgb());

            Assert.AreEqual(ExcelUnderLineType.Single, s.Style.Font.Underline);
            Assert.AreEqual("Times New Roman", s.Style.Font.Name);
        }

        [TestMethod]
        public void AddPivotLabelRow()
        {
            var ws = TryGetWorksheet(_pck, "StyleRowFieldLabelTot");
            var pt = ws.PivotTables[0];

            var s = pt.Styles[0];

            Assert.IsTrue(s.LabelOnly);
            Assert.IsFalse(s.DataOnly);
            Assert.AreEqual(pt.RowFields[0].Index, s.Conditions.Fields[0].FieldIndex);

            Assert.IsTrue(s.GrandRow);
            Assert.IsTrue(s.Style.Font.Italic.Value);
            Assert.IsTrue(s.Style.Font.Strike.Value);
            Assert.AreEqual("Times New Roman", s.Style.Font.Name);
        }
        [TestMethod]
        public void AddPivotLabelRowDf1()
        {
            var ws = TryGetWorksheet(_pck, "StyleRowFieldLabelTotDf1");
            var pt = ws.PivotTables[0];

            var s = pt.Styles[0];

            Assert.AreEqual(1, s.Conditions.DataFields.Count);
            Assert.AreEqual(1, s.Conditions.Fields.Count);
            Assert.AreEqual(pt.RowFields[0].Index, s.Conditions.Fields[0].FieldIndex);
            Assert.AreEqual("Stock", s.Conditions.DataFields[0].Name);

            Assert.IsTrue(s.GrandRow);
            Assert.IsTrue(s.Style.Font.Italic.Value);
            Assert.IsTrue(s.Style.Font.Strike.Value);
            Assert.AreEqual("Times New Roman", s.Style.Font.Name);
        }

        [TestMethod]
        public void ReadPivotLabelRowDataField2()
        {
            var ws = TryGetWorksheet(_pck, "StyleRowFieldDf2");
            var pt = ws.PivotTables[0];

            var s = pt.Styles[0];

            Assert.AreEqual(1, s.Conditions.DataFields.Count);
            Assert.AreEqual(1, s.Conditions.Fields.Count);
            Assert.AreEqual(pt.RowFields[0].Index, s.Conditions.Fields[0].FieldIndex);
            Assert.AreEqual("Stock", s.Conditions.DataFields[0].Name);

            Assert.IsTrue(s.Style.Font.Italic.Value);
            Assert.IsTrue(s.Style.Font.Strike.Value);
            Assert.AreEqual("Times New Roman", s.Style.Font.Name);
        }
        [TestMethod]
        public void ReadPivotLabelRowDataField2AndValue()
        {
            var ws = TryGetWorksheet(_pck, "StyleRowFieldDf2Value");
            var pt = ws.PivotTables[0];

            var s = pt.Styles[0];

            Assert.AreEqual(1, s.Conditions.DataFields.Count);
            Assert.AreEqual(1, s.Conditions.Fields.Count);
            Assert.AreEqual(1, s.Conditions.Fields[0].CacheItems.Count);
            Assert.AreEqual(pt.RowFields[0].Index, s.Conditions.Fields[0].FieldIndex);
            Assert.AreEqual("Stock", s.Conditions.DataFields[0].Name);

            Assert.AreEqual("Stock", s.Conditions.DataFields[0].Name);

            Assert.AreEqual("Screwdriver", s.Conditions.Fields[0].CacheItems[0].Value);

            Assert.IsTrue(s.Style.Font.Italic.Value);
            Assert.IsTrue(s.Style.Font.Strike.Value);
            Assert.AreEqual("Times New Roman", s.Style.Font.Name);

        }
        //[TestMethod]
        //public void AddPivotDataItemByIndex()
        //{
        //    var ws = _pck.Workbook.Worksheets.Add("PivotDataItemIndex");
        //    var pt = CreatePivotTable(ws);

        //    pt.CacheDefinition.Refresh();
        //    var s = pt.Styles.AddDataForCellReference(true, pt.Fields[0], pt.Fields[1]);
        //    s.References[0].AddReferenceByIndex(1);
        //    s.References[1].AddReferenceByIndex(0);
        //    s.References[2].AddReferenceByIndex(0);
        //    s.Style.Fill.Style = OfficeOpenXml.Style.eDxfFillStyle.PatternFill;
        //    s.Style.Fill.BackgroundColor.SetColor(Color.Red);
        //    s.Outline = true;
        //    //s.Axis = ePivotTableAxis.RowAxis;
        //    s.Style.Font.Color.SetColor(Color.Blue);
        //}
        //[TestMethod]
        //public void AddPivotDataItemByValue()
        //{
        //    var ws = _pck.Workbook.Worksheets.Add("PivotDataItemValue");
        //    var pt = CreatePivotTable(ws);

        //    pt.CacheDefinition.Refresh();
        //    var s = pt.Styles.AddDataForCellReference(true, pt.Fields[0], pt.Fields[1]);
        //    s.References[0].AddReferenceByValue("Stock");
        //    s.References[1].AddReferenceByValue("Apple");
        //    s.References[2].AddReferenceByValue("Groceries");
        //    s.Style.Fill.Style = OfficeOpenXml.Style.eDxfFillStyle.PatternFill;
        //    s.Style.Fill.BackgroundColor.SetColor(Color.Red);
        //    s.Outline = true;
        //    //s.Axis = ePivotTableAxis.RowAxis;
        //    s.Style.Font.Color.SetColor(Color.Blue);
        //}

        //[TestMethod]
        //public void AddButtonField()
        //{
        //    var ws = _pck.Workbook.Worksheets.Add("StyleFieldPage");
        //    var pt = CreatePivotTable(ws);

        //    var s = pt.Styles.AddButtonField(pt.Fields[4]);
        //    s.Style.Font.Color.SetColor(Color.Pink);
        //}

        //[TestMethod]
        //public void AddButtonRowAxis()
        //{
        //    var ws = _pck.Workbook.Worksheets.Add("StyleButtonRowAxis");
        //    var pt = CreatePivotTable(ws);

        //    var s = pt.Styles.AddButtonField(ePivotTableAxis.RowAxis);
        //    s.Style.Font.Underline = OfficeOpenXml.Style.ExcelUnderLineType.DoubleAccounting;
        //}
        //[TestMethod]
        //public void AddButtonColumnAxis()
        //{
        //    var ws = _pck.Workbook.Worksheets.Add("StyleButtonColumnAxis");
        //    var pt = CreatePivotTable(ws);

        //    var s3 = pt.Styles.AddButtonField(ePivotTableAxis.ColumnAxis);
        //    s3.Style.Font.Italic = true;
        //}
        //[TestMethod]
        //public void AddButtonPageAxis()
        //{
        //    var ws = _pck.Workbook.Worksheets.Add("StyleButtonPageAxis");
        //    var pt = CreatePivotTable(ws);

        //    var s4 = pt.Styles.AddButtonField(ePivotTableAxis.PageAxis);
        //    s4.Style.Font.Color.SetColor(Color.ForestGreen);
        //}


        //[TestMethod]
        //public void AddTopStart()
        //{
        //    var ws = _pck.Workbook.Worksheets.Add("StyleTopStart");
        //    var pt = CreatePivotTable(ws);

        //    //Top Left cells 
        //    var styleTopLeft = pt.Styles.AddTopStart();
        //    styleTopLeft.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //    styleTopLeft.Style.Fill.BackgroundColor.SetColor(Color.Red);
        //}
        //[TestMethod]
        //public void AddTopStartOffset0()
        //{
        //    var ws = _pck.Workbook.Worksheets.Add("StyleTopStartOffset0");
        //    var pt = CreatePivotTable(ws);

        //    //Top Left cells 
        //    var styleTopLeft = pt.Styles.AddTopStart(0);
        //    styleTopLeft.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //    styleTopLeft.Style.Fill.BackgroundColor.SetColor(Color.Blue);
        //}

        //[TestMethod]
        //public void AddTopEnd()
        //{
        //    var ws = _pck.Workbook.Worksheets.Add("StyleTopEnd");
        //    var pt = CreatePivotTable(ws);

        //    var styleTopRight2 = pt.Styles.AddTopEnd();
        //    styleTopRight2.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //    styleTopRight2.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
        //}
        //[TestMethod]
        //public void AddTopEndOffset1()
        //{
        //    var ws = _pck.Workbook.Worksheets.Add("StyleTopEndOffset1");
        //    var pt = CreatePivotTable(ws);

        //    var styleTopRight2 = pt.Styles.AddTopEnd(1);
        //    styleTopRight2.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //    styleTopRight2.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
        //}

    }
}
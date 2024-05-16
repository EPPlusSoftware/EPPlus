/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * Required Notice: Copyright (C) EPPlus Software AB. 
 * https://epplussoftware.com
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "" as is "" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
  Date               Author                       Change
 *******************************************************************************
  02/11/2020         EPPlus Software AB       Initial release EPPlus 5
 *******************************************************************************/
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Xml;

namespace EPPlusTest.Table.PivotTable
{
    [TestClass]
    public class PivotTableTests : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenPackage("PivotTable.xlsx", true);
            var ws = _pck.Workbook.Worksheets.Add("Data");
            LoadItemData(ws);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void ValidateLoadSaveTableSource()
        {
            using (ExcelPackage p1 = new ExcelPackage())
            {
                var tblName = "Table1";
                var tblAddress = "A1:D4";
                var wsData = p1.Workbook.Worksheets.Add("TableData");
                wsData.Cells["A1"].Value = "Column1";
                wsData.Cells["B1"].Value = "Column2";
                wsData.Cells["C1"].Value = "Column3";
                wsData.Cells["D1"].Value = "Column4";
                var wsPivot = p1.Workbook.Worksheets.Add("PivotSimple");
                var Table1 = wsData.Tables.Add(wsData.Cells[tblAddress], tblName);
                var pivotTable1 = wsPivot.PivotTables.Add(wsPivot.Cells["A1"], wsData.Cells[Table1.Address.Address], "PivotTable1");

                pivotTable1.RowFields.Add(pivotTable1.Fields[0]);
                pivotTable1.DataFields.Add(pivotTable1.Fields[1]);
                pivotTable1.ColumnFields.Add(pivotTable1.Fields[2]);

                Assert.AreEqual(tblAddress, wsPivot.PivotTables[0].CacheDefinition.SourceRange.Address);
                Assert.AreEqual(Table1.Columns.Count, pivotTable1.Fields.Count);
                Assert.AreEqual(1, pivotTable1.RowFields.Count);
                Assert.AreEqual(1, pivotTable1.DataFields.Count);
                Assert.AreEqual(1, pivotTable1.ColumnFields.Count);

                p1.Save();

                using (var p2 = new ExcelPackage(p1.Stream))
                {
                    wsData = p2.Workbook.Worksheets[0];
                    wsPivot = p2.Workbook.Worksheets[1];

                    pivotTable1 = wsPivot.PivotTables[0];
                    Assert.AreEqual(tblAddress, pivotTable1.CacheDefinition.SourceRange.Address);
                    Assert.AreEqual(Table1.Columns.Count, pivotTable1.Fields.Count);
                    Assert.AreEqual(1, pivotTable1.RowFields.Count);
                    Assert.AreEqual(1, pivotTable1.DataFields.Count);
                    Assert.AreEqual(1, pivotTable1.ColumnFields.Count);
                }
            }
        }
        [TestMethod]
        public void ValidateLoadSaveAddressSource()
        {
            using (ExcelPackage p1 = new ExcelPackage())
            {
                var address = "A1:D4";
                var wsData = p1.Workbook.Worksheets.Add("TableData");
                wsData.Cells["A1"].Value = "Column1";
                wsData.Cells["B1"].Value = "Column2";
                wsData.Cells["C1"].Value = "Column3";
                wsData.Cells["D1"].Value = "Column4";
                var wsPivot = p1.Workbook.Worksheets.Add("PivotSimple");
                var pivotTable1 = wsPivot.PivotTables.Add(wsPivot.Cells["A1"], wsData.Cells[address], "PivotTable1");
                pivotTable1.RowFields.Add(pivotTable1.Fields[0]);
                pivotTable1.DataFields.Add(pivotTable1.Fields[1]);
                pivotTable1.ColumnFields.Add(pivotTable1.Fields[2]);

                Assert.AreEqual(address, wsPivot.PivotTables[0].CacheDefinition.SourceRange.Address);
                Assert.AreEqual(4, pivotTable1.Fields.Count);
                Assert.AreEqual(1, pivotTable1.RowFields.Count);
                Assert.AreEqual(1, pivotTable1.DataFields.Count);
                Assert.AreEqual(1, pivotTable1.ColumnFields.Count);

                p1.Save();

                using (var p2 = new ExcelPackage(p1.Stream))
                {
                    wsData = p2.Workbook.Worksheets[0];
                    wsPivot = p2.Workbook.Worksheets[1];

                    pivotTable1 = wsPivot.PivotTables[0];
                    Assert.AreEqual(address, pivotTable1.CacheDefinition.SourceRange.Address);
                    Assert.AreEqual(4, pivotTable1.Fields.Count);
                    Assert.AreEqual(1, pivotTable1.RowFields.Count);
                    Assert.AreEqual(1, pivotTable1.DataFields.Count);
                    Assert.AreEqual(1, pivotTable1.ColumnFields.Count);
                }
            }
        }

        [TestMethod]
        public void CreatePivotTableAddressSource()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotSourceAddress");
            LoadTestdata(ws);

            var pivotTable1 = ws.PivotTables.Add(ws.Cells["G1"], ws.Cells["A1:D100"], "PivotTable1");

            pivotTable1.RowFields.Add(pivotTable1.Fields[0]);
            pivotTable1.RowFields.Add(pivotTable1.Fields[2]);
            pivotTable1.DataFields.Add(pivotTable1.Fields[1]);
            pivotTable1.DataFields.Add(pivotTable1.Fields[3]);
        }
        [TestMethod]
        public void CreatePivotTableTableSource()
        {
            var ws = _pck.Workbook.Worksheets.Add("PivotSourceTable");
            LoadTestdata(ws);
            var table = ws.Tables.Add(ws.Cells["A1:D100"], "table1");
            var pivotTable1 = ws.PivotTables.Add(ws.Cells["G1"], table, "PivotTable1");

            pivotTable1.RowFields.Add(pivotTable1.Fields[0]);
            pivotTable1.RowFields.Add(pivotTable1.Fields[2]);
            pivotTable1.DataFields.Add(pivotTable1.Fields[1]);
            pivotTable1.DataFields.Add(pivotTable1.Fields[3]);
        }
        [TestMethod]
        public void RowsDataOnColumns()
        {
            var wsData = _pck.Workbook.Worksheets["Data"];
            var ws = _pck.Workbook.Worksheets.Add("Rows-Data on columns");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable1");
            pt.GrandTotalCaption = "Total amount";
            pt.RowFields.Add(pt.Fields[1]);
            pt.RowFields.Add(pt.Fields[0]);
            pt.DataFields.Add(pt.Fields[3]);
            pt.DataFields.Add(pt.Fields[2]);
            pt.DataFields[0].Function = DataFieldFunctions.Product;
            pt.DataOnRows = false;
        }
        [TestMethod]
        public void RowsDataOnRow()
        {
            var wsData = _pck.Workbook.Worksheets["Data"];
            var ws = _pck.Workbook.Worksheets.Add("Rows-Data on rows");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[1]);
            pt.RowFields.Add(pt.Fields[0]);
            pt.DataFields.Add(pt.Fields[3]);
            pt.DataFields.Add(pt.Fields[2]);
            pt.DataFields[0].Function = DataFieldFunctions.Average;
            pt.DataOnRows = true;
        }
        [TestMethod]
        public void ColumnsDataOnColumns()
        {
            var wsData = _pck.Workbook.Worksheets["Data"];
            var ws = _pck.Workbook.Worksheets.Add("Columns-Data on columns");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable3");
            pt.ColumnFields.Add(pt.Fields[1]);
            pt.ColumnFields.Add(pt.Fields[0]);
            pt.DataFields.Add(pt.Fields[3]);
            pt.DataFields.Add(pt.Fields[2]);
            pt.DataOnRows = false;
        }
        [TestMethod]
        public void ColumnsDataOnRows()
        {
            var wsData = _pck.Workbook.Worksheets["Data"];
            var ws = _pck.Workbook.Worksheets.Add("Columns-Data on rows");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable4");
            pt.ColumnFields.Add(pt.Fields[1]);
            pt.ColumnFields.Add(pt.Fields[0]);
            pt.DataFields.Add(pt.Fields[3]);
            pt.DataFields.Add(pt.Fields[2]);
            pt.DataOnRows = true;
        }
        [TestMethod]
        public void ColumnsRows_DataOnColumns()
        {
            var wsData = _pck.Workbook.Worksheets["Data"];
            var ws = _pck.Workbook.Worksheets.Add("Columns/Rows-Data on columns");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable5");
            pt.ColumnFields.Add(pt.Fields[1]);
            pt.RowFields.Add(pt.Fields[0]);
            pt.DataFields.Add(pt.Fields[3]);
            pt.DataFields.Add(pt.Fields[2]);
            pt.DataOnRows = false;
        }
        [TestMethod]
        public void ColumnsRows_DataOnRows()
        {
            var wsData = _pck.Workbook.Worksheets["Data"];
            var ws = _pck.Workbook.Worksheets.Add("Columns/Rows-Data on rows");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable6");
            pt.ColumnFields.Add(pt.Fields[1]);
            pt.RowFields.Add(pt.Fields[0]);
            pt.DataFields.Add(pt.Fields[3]);
            pt.DataFields.Add(pt.Fields[2]);
            pt.DataOnRows = true;
            ws.Drawings.AddChart("Pivotchart6", OfficeOpenXml.Drawing.Chart.eChartType.BarStacked3D, pt);
        }
        [TestMethod]
        public void RowsPage_DataOnColumns()
        {
            var wsData = _pck.Workbook.Worksheets["Data"];
            var ws = _pck.Workbook.Worksheets.Add("Rows/Page-Data on Columns");

            var pt = ws.PivotTables.Add(ws.Cells["A3"], wsData.Cells["K1:N11"], "Pivottable7");
            pt.PageFields.Add(pt.Fields[1]);
            pt.RowFields.Add(pt.Fields[0]);
            pt.DataFields.Add(pt.Fields[3]);
            pt.DataFields.Add(pt.Fields[2]);
            pt.DataOnRows = false;

            pt.Fields[0].SubTotalFunctions = eSubTotalFunctions.Sum | eSubTotalFunctions.Max;
            Assert.AreEqual(pt.Fields[0].SubTotalFunctions, eSubTotalFunctions.Sum | eSubTotalFunctions.Max);

            pt.Fields[0].SubTotalFunctions = eSubTotalFunctions.Sum | eSubTotalFunctions.Product | eSubTotalFunctions.StdDevP;
            Assert.AreEqual(pt.Fields[0].SubTotalFunctions, eSubTotalFunctions.Sum | eSubTotalFunctions.Product | eSubTotalFunctions.StdDevP);

            pt.Fields[0].SubTotalFunctions = eSubTotalFunctions.None;
            Assert.AreEqual(pt.Fields[0].SubTotalFunctions, eSubTotalFunctions.None);

            pt.Fields[0].SubTotalFunctions = eSubTotalFunctions.Default;
            Assert.AreEqual(pt.Fields[0].SubTotalFunctions, eSubTotalFunctions.Default);

            pt.Fields[0].Sort = eSortType.Descending;
            pt.TableStyle = OfficeOpenXml.Table.TableStyles.Medium14;
        }
        [TestMethod]
        public void Pivot_GroupDate()
        {
            var wsData = _pck.Workbook.Worksheets["Data"];
            var ws = _pck.Workbook.Worksheets.Add("Pivot-Group Date");

            var pt = ws.PivotTables.Add(ws.Cells["A3"], wsData.Cells["K1:O11"], "Pivottable8");
            pt.RowFields.Add(pt.Fields[1]);
            pt.RowFields.Add(pt.Fields[4]);
            pt.Fields[4].AddDateGrouping(eDateGroupBy.Years | eDateGroupBy.Months | eDateGroupBy.Days | eDateGroupBy.Quarters, new DateTime(2010, 01, 31), new DateTime(2010, 11, 30));
            pt.RowHeaderCaption = "År";
            pt.Fields[4].Name = "Dag";
            pt.Fields[4].Items[0].Hidden = true;
            pt.Fields[5].Name = "Månad";
            pt.Fields[5].Items[0].Hidden = true;
            pt.Fields[6].Name = "Kvartal";
            pt.Fields[6].Items[0].Hidden = true;
            pt.GrandTotalCaption = "Totalt";

            pt.DataFields.Add(pt.Fields[3]);
            pt.DataFields.Add(pt.Fields[2]);
            pt.DataOnRows = true;

            pt = ws.PivotTables.Add(ws.Cells["H3"], wsData.Cells["K1:O11"], "Pivottable10");
            pt.RowFields.Add(pt.Fields[1]);
            pt.RowFields.Add(pt.Fields[4]);
            pt.Fields[4].AddDateGrouping(7, new DateTime(2010, 01, 31), new DateTime(2010, 11, 30));
            pt.RowHeaderCaption = "Veckor";
            pt.GrandTotalCaption = "Totalt";

            pt = ws.PivotTables.Add(ws.Cells["A60"], wsData.Cells["K1:O11"], "Pivottable11");
            pt.RowFields.Add(pt.Fields["Category"]);
            pt.RowFields.Add(pt.Fields["Item"]);
            pt.RowFields.Add(pt.Fields[4]);

            pt.DataFields.Add(pt.Fields[3]);
            pt.DataFields.Add(pt.Fields[2]);

            pt.DataOnRows = true;

        }
        [TestMethod]
        public void Pivot_GroupNumber()
        {
            var wsData = _pck.Workbook.Worksheets["Data"];
            var ws = _pck.Workbook.Worksheets.Add("Pivot-Group Number");
            var pt = ws.PivotTables.Add(ws.Cells["A3"], wsData.Cells["K1:N11"], "Pivottable9");
            pt.PageFields.Add(pt.Fields[1]);
            pt.RowFields.Add(pt.Fields[3]);
            pt.RowFields[0].AddNumericGrouping(-3.3, 5.5, 4.0);
            pt.DataFields.Add(pt.Fields[2]);
            pt.RowFields[0].Items[0].Hidden = true;
            pt.RowFields[0].Items[1].Hidden = true;
            pt.RowFields[0].Items[2].Hidden = true;
            pt.RowFields[0].Items[3].Hidden = true;
            pt.DataOnRows = false;
            pt.TableStyle = OfficeOpenXml.Table.TableStyles.Medium14;
        }
        [TestMethod]
        public void Pivot_ManyRowFields()
        {
            var wsData = _pck.Workbook.Worksheets["Data"];
            var ws = _pck.Workbook.Worksheets.Add("Pivot-Many RowFields");

            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:O11"], "Pivottable10");
            pt.ColumnFields.Add(pt.Fields[1]);
            pt.RowFields.Add(pt.Fields[0]);
            pt.RowFields.Add(pt.Fields[3]);
            pt.RowFields.Add(pt.Fields[2]);
            pt.RowFields.Add(pt.Fields[4]);
            pt.DataOnRows = true;
            pt.ColumnHeaderCaption = "Column Caption";
            pt.RowHeaderCaption = "Row Caption";
        }
        [TestMethod]
        public void Pivot_Blank()
        {
            var wsData = _pck.Workbook.Worksheets["Data"];
            var ws = _pck.Workbook.Worksheets.Add("Pivot-Blank");

            wsData.Cells["A1"].Value = "Column1";
            wsData.Cells["B1"].Value = "Column2";
            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["A1:B2"], "Pivottable11");
            pt.ColumnFields.Add(pt.Fields[1]);
            var rf = pt.RowFields.Add(pt.Fields[0]);
            rf.SubTotalFunctions = eSubTotalFunctions.None;
            pt.DataOnRows = true;
        }
        [TestMethod]
        public void Pivot_SaveDataFalse()
        {
            var wsData = _pck.Workbook.Worksheets["Data"];
            var ws = _pck.Workbook.Worksheets.Add("Pivot-NoRecord");

            wsData.Cells["A1"].Value = "Column1";
            wsData.Cells["B1"].Value = "Column2";
            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["A1:B2"], "Pivottable11");
            pt.ColumnFields.Add(pt.Fields[1]);
            var rf = pt.RowFields.Add(pt.Fields[0]);
            rf.SubTotalFunctions = eSubTotalFunctions.None;
            pt.DataOnRows = true;
            pt.CacheDefinition.SaveData = false;
        }
        [TestMethod]
        public void Pivot_SavedDataTrue()
        {
            var wsData = _pck.Workbook.Worksheets["Data"];
            var ws = _pck.Workbook.Worksheets.Add("Pivot-WithRecord");

            wsData.Cells["A1"].Value = "Column1";
            wsData.Cells["B1"].Value = "Column2";
            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["A1:B2"], "Pivottable11");
            pt.ColumnFields.Add(pt.Fields[1]);
            var rf = pt.RowFields.Add(pt.Fields[0]);
            rf.SubTotalFunctions = eSubTotalFunctions.None;
            pt.DataOnRows = true;
            pt.CacheDefinition.SaveData = false;    //Remove the record xml
            pt.CacheDefinition.SaveData = true;     //Add the record xml
        }
        [TestMethod]
        public void Pivot_ManyPageFields()
        {
            var wsData = _pck.Workbook.Worksheets["Data"];
            var ws = _pck.Workbook.Worksheets.Add("Pivot-Many PageFields");

            var pt = ws.PivotTables.Add(ws.Cells["A3"], wsData.Cells["K1:O11"], "Pivottable12");
            pt.ColumnFields.Add(pt.Fields[1]);
            pt.RowFields.Add(pt.Fields[0]);
            var pf1 = pt.PageFields.Add(pt.Fields[2]);
            pf1.Items.Refresh();
            pf1.Items[1].Hidden = true;
            pf1.Items[8].Hidden = true;


            var pf2 = pt.PageFields.Add(pt.Fields[4]);
            pf2.Items.Refresh();
            pf2.Items[1].Hidden = true;
            pf1.MultipleItemSelectionAllowed = true;
            pf2.MultipleItemSelectionAllowed = true;
            pt.DataFields.Add(pt.Fields[3]);
            pt.DataOnRows = true;
            pt.ColumnHeaderCaption = "Column Caption";
            pt.RowHeaderCaption = "Row Caption";

            Assert.AreEqual(1, pt.ColumnFields.Count);
            Assert.AreEqual(2, pt.PageFields.Count);
            Assert.AreEqual(1, pt.RowFields.Count);
            Assert.AreEqual(1, pt.DataFields.Count);
            Assert.IsTrue(pf1.MultipleItemSelectionAllowed);
        }
        [TestMethod]
        public void Pivot_StylingFieldsFalse()
        {
            var wsData = _pck.Workbook.Worksheets["Data"];
            var ws = _pck.Workbook.Worksheets.Add("Pivot-StylingFieldsFalse");

            var pt = ws.PivotTables.Add(ws.Cells["A3"], wsData.Cells["K1:O11"], "Pivottable12");
            pt.ColumnFields.Add(pt.Fields[1]);
            pt.RowFields.Add(pt.Fields[0]);
            var df = pt.DataFields.Add(pt.Fields[3]);
            pt.DataOnRows = true;
            pt.ColumnHeaderCaption = "Column Caption";
            pt.RowHeaderCaption = "Row Caption";

            Assert.IsTrue(pt.ShowColumnHeaders);
            Assert.IsFalse(pt.ShowColumnStripes);
            Assert.IsTrue(pt.ShowRowHeaders);
            Assert.IsFalse(pt.ShowRowStripes);
            Assert.IsTrue(pt.ShowLastColumn);

            pt.ShowColumnHeaders = false;
            pt.ShowColumnStripes = true;
            pt.ShowRowHeaders = false;
            pt.ShowRowStripes = true;
            pt.ShowLastColumn = false;

            Assert.IsFalse(pt.ShowColumnHeaders);
            Assert.IsTrue(pt.ShowColumnStripes);
            Assert.IsFalse(pt.ShowRowHeaders);
            Assert.IsTrue(pt.ShowRowStripes);
            Assert.IsFalse(pt.ShowLastColumn);

            Assert.AreEqual(1, pt.ColumnFields.Count);
            Assert.AreEqual(1, pt.RowFields.Count);
            Assert.AreEqual(1, pt.DataFields.Count);

        }
        [TestMethod]
        public void RowsDataOnRow_WithNumberFormat()
        {
            var wsData = _pck.Workbook.Worksheets["Data"];
            var ws = _pck.Workbook.Worksheets.Add("PivotTable with numberformat");
            var pt = ws.PivotTables.Add(ws.Cells["A1"], wsData.Cells["K1:N11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[1]);
            pt.RowFields.Add(pt.Fields[0]);
            pt.DataFields.Add(pt.Fields[3]);
            pt.DataFields.Add(pt.Fields[2]);

            pt.Fields[3].Format = "#,##0";
            pt.Fields[3].Cache.Format = "#,##0.000";
            ws.Workbook.Styles.UpdateXml();

            Assert.AreEqual(3, pt.Fields[3].NumFmtId);
            Assert.AreEqual(165, pt.Fields[3].Cache.NumFmtId);
        }
        [TestMethod]
        public void AddCalculatedField()
        {
            var ws = _pck.Workbook.Worksheets.Add("CalculatedField");

            LoadTestdata(ws);
            var formula = "NumValue*2";
            var tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTable1");
            var f = tbl.Fields.AddCalculatedField("NumValueX2", formula);
            f.Format = "#,##0";
            var rf = tbl.RowFields.Add(tbl.Fields[1]);
            var df1 = tbl.DataFields.Add(tbl.Fields[3]);
            var df2 = tbl.DataFields.Add(tbl.Fields["NumValueX2"]);
            df1.Function = DataFieldFunctions.Sum;
            df2.Function = DataFieldFunctions.Sum;
            tbl.DataOnRows = false;
            Assert.AreEqual("NumValue*2", tbl.Fields[4].Cache.Formula);
        }
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void ShouldThrowExceptionOnAddingCalculatedFieldToColumns()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("RowArgExcep");
                LoadTestdata(ws);
                var formula = "NumValue*2";
                var tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTable1");
                tbl.Fields.AddCalculatedField("NumValueX2", formula);
                var rf = tbl.ColumnFields.Add(tbl.Fields[4]);
            }
        }
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void ShouldThrowExceptionOnAddingCalculatedFieldToRow()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("RowArgExcep");
                LoadTestdata(ws);
                var formula = "NumValue*2";
                var tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTable1");
                tbl.Fields.AddCalculatedField("NumValueX2", formula);
                var rf = tbl.RowFields.Add(tbl.Fields[4]);
            }
        }
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void ShouldThrowExceptionOnAddingCalculatedFieldToPage()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("RowArgExcep");
                LoadTestdata(ws);
                var formula = "NumValue*2";
                var tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTable1");
                tbl.Fields.AddCalculatedField("NumValueX2", formula);
                var rf = tbl.PageFields.Add(tbl.Fields[4]);
            }
        }
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void ShouldThrowExceptionOnAddingCalculatedFieldWithBlankFormula()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("RowArgExcep");
                LoadTestdata(ws);
                var tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTable1");
                tbl.Fields.AddCalculatedField("NumValueX2", "");
            }
        }
        [TestMethod]
        public void PivotTableStyleTests()
        {
            var wsData = _pck.Workbook.Worksheets["Data"];
            var ws = _pck.Workbook.Worksheets.Add("StyleTests");

            var pt = ws.PivotTables.Add(ws.Cells["A3"], wsData.Cells["K1:N11"], "Pivottable8");
            pt.PivotTableStyle = PivotTableStyles.None;
            Assert.AreEqual(PivotTableStyles.None, pt.PivotTableStyle);
            Assert.AreEqual(TableStyles.None, pt.TableStyle);

            pt.PivotTableStyle = PivotTableStyles.Medium28;
            Assert.AreEqual(PivotTableStyles.Medium28, pt.PivotTableStyle);
            Assert.AreEqual(TableStyles.Medium28, pt.TableStyle);

            pt.PivotTableStyle = PivotTableStyles.Dark28;
            Assert.AreEqual(PivotTableStyles.Dark28, pt.PivotTableStyle);
            Assert.AreEqual(TableStyles.Custom, pt.TableStyle);
            Assert.AreEqual("PivotStyleDark28", pt.StyleName);

            pt.TableStyle = TableStyles.Light15;
            Assert.AreEqual(PivotTableStyles.Light15, pt.PivotTableStyle);
            Assert.AreEqual(TableStyles.Light15, pt.TableStyle);
            Assert.AreEqual("PivotStyleLight15", pt.StyleName);


            pt.PivotTableStyle = PivotTableStyles.Light28;
            Assert.AreEqual(PivotTableStyles.Light28, pt.PivotTableStyle);
            Assert.AreEqual(TableStyles.Custom, pt.TableStyle);
            Assert.AreEqual("PivotStyleLight28", pt.StyleName);
        }

        [TestMethod]
        public void HideValuesRow()
        {
            var ws = _pck.Workbook.Worksheets.Add("HideValuesRow");

            LoadTestdata(ws);
            var tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTable1");
            var rf = tbl.RowFields.Add(tbl.Fields[1]);
            var df = tbl.DataFields.Add(tbl.Fields[3]);
            df.Function = DataFieldFunctions.Sum;
            tbl.DataOnRows = false;
            tbl.GridDropZones = false;
            Assert.IsTrue(tbl.ShowValuesRow);
            tbl.ShowValuesRow = false;
            Assert.IsFalse(tbl.ShowValuesRow);
            tbl.ShowValuesRow = true;
            Assert.IsTrue(tbl.ShowValuesRow);
            tbl.ShowValuesRow = false;
        }
        [TestMethod]
        public void ValidateSharedItemsAreCaseInsensitive()
        {
            var ws = _pck.Workbook.Worksheets.Add("CaseInsentitive");

            ws.Cells["A1"].Value = "Column1";
            ws.Cells["B1"].Value = "Column2";
            ws.Cells["A2"].Value = "Value1";
            ws.Cells["B2"].Value = 1;
            ws.Cells["A3"].Value = "value1";
            ws.Cells["B3"].Value = 2;
            var tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:B3"], "CIPivotTable");
            var rf = tbl.RowFields.Add(tbl.Fields[0]);
            var df = tbl.DataFields.Add(tbl.Fields[1]);
            rf.Cache.Refresh();
            Assert.AreEqual(1, rf.Cache.SharedItems.Count);
            Assert.AreEqual("Value1", rf.Cache.SharedItems[0]);
        }
        [TestMethod]
        public void ValidateAttributesWhenNumbericAndMissing()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("NumericAndNull");
                ws.Cells["A1"].Value = "Int";
                ws.Cells["A2"].Value = 1;
                ws.Cells["A3"].Value = 2;
                ws.Cells["A4"].Value = 2;

                ws.Cells["B1"].Value = "Float";
                ws.Cells["B2"].Value = 1.3;
                ws.Cells["B3"].Value = 2.4;
                ws.Cells["B4"].Value = 5.6;

                ws.Cells["C1"].Value = "IntFloat";
                ws.Cells["C2"].Value = 3;
                ws.Cells["C3"].Value = 2.4;
                ws.Cells["C4"].Value = 2;

                ws.Cells["D1"].Value = "IntNull";
                ws.Cells["D2"].Value = 3;
                ws.Cells["D4"].Value = 3;

                ws.Cells["E1"].Value = "FloatNull";
                ws.Cells["E3"].Value = 4.2;
                ws.Cells["E4"].Value = 5.7;

                ws.Cells["F1"].Value = "IntFloatNull";
                ws.Cells["F2"].Value = 5;
                ws.Cells["F4"].Value = 6.2;

                ws.Cells["G1"].Value = "StringNull";
                ws.Cells["G2"].Value = "Value 1";
                ws.Cells["G4"].Value = "Value 3";

                ws.Cells["H1"].Value = "MixedIntBool";
                ws.Cells["H2"].Value = 1;
                ws.Cells["H4"].Value = true;

                ws.Cells["I1"].Value = "Mixed float";
                ws.Cells["I3"].Value = 3.3;
                ws.Cells["I4"].Value = "Value 3";


                var tbl = ws.PivotTables.Add(ws.Cells["K3"], ws.Cells["A1:I4"], "ptNumberMissing");
                var pf1 = tbl.PageFields.Add(tbl.Fields[0]);
                var pf2 = tbl.PageFields.Add(tbl.Fields[1]);
                var pf3 = tbl.PageFields.Add(tbl.Fields[2]);
                var pf4 = tbl.PageFields.Add(tbl.Fields[3]);
                var pf5 = tbl.PageFields.Add(tbl.Fields[4]);
                var pf6 = tbl.PageFields.Add(tbl.Fields[5]);
                var pf7 = tbl.PageFields.Add(tbl.Fields[6]);
                var pf8 = tbl.PageFields.Add(tbl.Fields[7]);
                var pf9 = tbl.PageFields.Add(tbl.Fields[8]);

                tbl.CacheDefinition.Refresh();

                p.Save();

                AssertShartedItemsAttributes(pf1.Cache.TopNode.FirstChild, 4, true, true, false, false, false);
                AssertShartedItemsAttributes(pf2.Cache.TopNode.FirstChild, 3, true, false, false, false, false);
                AssertShartedItemsAttributes(pf3.Cache.TopNode.FirstChild, 3, true, false, false, false, false);
                AssertShartedItemsAttributes(pf4.Cache.TopNode.FirstChild, 4, true, true, true, false, false);
                AssertShartedItemsAttributes(pf5.Cache.TopNode.FirstChild, 3, true, false, true, false, false);
                AssertShartedItemsAttributes(pf6.Cache.TopNode.FirstChild, 3, true, false, true, false, false);
                AssertShartedItemsAttributes(pf7.Cache.TopNode.FirstChild, 1, false, false, true, false, false);

                AssertShartedItemsAttributes(pf8.Cache.TopNode.FirstChild, 4, true, true, true, false, true);
                AssertShartedItemsAttributes(pf9.Cache.TopNode.FirstChild, 3, true, false, true, false, true);
            }
        }

        private void AssertShartedItemsAttributes(XmlNode node, int count, bool numberValues, bool intValues, bool containsBlanks, bool semiMixedValues, bool mixedValues)
        {
            if (node.Attributes.Count != count)
            {
                Assert.Fail("Wrong attrib Count");
            }
            AssertContains(node, "containsNumber", numberValues);
            AssertContains(node, "containsInteger", intValues);
            AssertContains(node, "containsBlank", containsBlanks);
            AssertContains(node, "containsSemiMixedTypes", semiMixedValues);
            AssertContains(node, "containsMixedTypes", mixedValues);

            //containsInteger = "1" containsNumber = "1" containsString = "0" containsSemiMixedTypes = "0"
        }

        private void AssertContains(XmlNode node, string attrName, bool value)
        {
            var a = node.Attributes[attrName];
            if (a == null)
            {
                if (value)
                {
                    Assert.Fail($"{attrName} value not false");
                }
            }
            else
            {
                if (value && a.Value != "1")
                {
                    Assert.Fail($"{attrName} value not true");
                }
            }
        }
        [TestMethod]
        public void CopyPivotTableToExternalPackageSameWorksheetAsData()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("data");
                LoadTestdata(ws);

                var pivotTable1 = ws.PivotTables.Add(ws.Cells["G1"], ws.Cells["A1:D100"], "PivotTable1");
                pivotTable1.RowFields.Add(pivotTable1.Fields[0]);
                pivotTable1.RowFields.Add(pivotTable1.Fields[2]);
                pivotTable1.DataFields.Add(pivotTable1.Fields[1]);
                pivotTable1.DataFields.Add(pivotTable1.Fields[3]);

                using (var p2 = new ExcelPackage())
                {
                    var wsNew = p2.Workbook.Worksheets.Add("PivotCopy", ws);
                    SaveWorkbook("copiedPivot.xlsx", p2);
                }
                SaveWorkbook("Pivot.xlsx", p);
            }
        }

        [TestMethod]
        public void PivotTableCalculation()
        {
            using (var p = OpenTemplatePackage("PivotTestPackage.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];

                var pivotWs = p.Workbook.Worksheets[1];

                var ogPT = pivotWs.PivotTables[0];

                var pt = pivotWs.PivotTables.Add(pivotWs.Cells["F2"], ws.Cells["A1:E15"], "PivotTableNew");
                pt.RowFields.Add(pt.Fields["Category"]);
                pt.RowFields.Add(pt.Fields["Product"]);

                pt.DataFields.Add(pt.Fields["Year"]);
                pt.DataFields.Add(pt.Fields["Sales"]);

                pt.DataOnRows = false;

                pivotWs.Cells["B28"].Formula = "GETPIVOTDATA(\"Sales\",$B$2)";
                pivotWs.Cells["B29"].Formula = "GETPIVOTDATA(\"Sales\",$B$2,\"Category\",\"Clothing\")";
                pivotWs.Cells["B30"].Formula = "GETPIVOTDATA(\"Sales\",$B$2,\"Category\",\"Clothing\",\"Product\",\"Shirts\")";

                pt.Calculate(true);
                ogPT.Calculate(true);

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void GetPivotDataStressTest()
        {
            using (var p = OpenTemplatePackage("MonsterTable.xlsx"))
            {
                var ws = p.Workbook.Worksheets.GetByName("MonsterPT");
                var dataWs = p.Workbook.Worksheets.GetByName("MonsterData");

                dataWs.Cells["A12"].Value = "Wolf";
                dataWs.Cells["B12"].Value = "Beast";
                dataWs.Cells["K12"].Value = 10;

                var table = ws.PivotTables[0];

                var dataField = table.DataFields["Sum of Hit Points"];
                dataField.Function = DataFieldFunctions.Average;
                dataField.ShowDataAs.SetPercentOfColumn();

                ws.Cells["F40"].Formula = "GETPIVOTDATA(CONCAT(\"\",$B$3,\"\"),$A$3,\"Name\",\"Abominable Yeti\",\"Type\",\"Monstrosity\",\"Source\",\"Monster Manual\",\"Size\",\"Large\")";
                ws.Cells["F41"].Formula = "GETPIVOTDATA(\"Sum of Hit Points\",$A$3,\"Type\",\"Monstrosity\",\"Source\",\"Monster Manual\")";
                ws.Cells["F42"].Formula = "GETPIVOTDATA(\"Sum of Speed\",$A$3,\"Name\",\"Winter Wolf\",\"Type\",\"Monstrosity\",\"Source\",\"Monster Manual\",\"Size\",\"Large\")";
                ws.Cells["F43"].Formula = "GETPIVOTDATA(\"Sum of Speed\",$A$3,\"Type\",\"Ooze\",\"Source\",\"Monster Manual\",\"Size\",\"Large\")";

                //table.Calculate(true);
                ws.Calculate();

                Assert.AreEqual(ws.Cells["E40"].Value, ws.Cells["F40"].Value);
                Assert.AreEqual(ws.Cells["E41"].Value, ws.Cells["F41"].Value);
                Assert.AreEqual(ws.Cells["E42"].Value, ws.Cells["F42"].Value);
                Assert.AreEqual(ws.Cells["E43"].Value, ws.Cells["F43"].Value);

                SaveAndCleanup(p);
            }
        }

        void SetShowDataAs(eShowDataAs showAs, ExcelPivotTableDataField dt, ExcelPivotTableField compareField, ePrevNextPivotItem item)
        {
            switch (showAs)
            {
                case eShowDataAs.Difference:
                    dt.ShowDataAs.SetDifference(compareField, ePrevNextPivotItem.Previous);
                    break;
                case eShowDataAs.Index:
                    dt.ShowDataAs.SetIndex();
                    break;
                case eShowDataAs.Normal:
                    dt.ShowDataAs.SetNormal();
                    break;
                case eShowDataAs.Percent:
                    dt.ShowDataAs.SetPercent(compareField, ePrevNextPivotItem.Next);
                    break;
                case eShowDataAs.PercentDifference:
                    dt.ShowDataAs.SetPercentageDifference(compareField, ePrevNextPivotItem.Next);
                    break;
                case eShowDataAs.PercentOfColumn:
                    dt.ShowDataAs.SetPercentOfColumn();
                    break;
                case eShowDataAs.PercentOfRow:
                    dt.ShowDataAs.SetPercentOfRow();
                    break;
                case eShowDataAs.PercentOfTotal:
                    dt.ShowDataAs.SetPercentOfTotal();
                    break;
                case eShowDataAs.RunningTotal:
                    dt.ShowDataAs.SetPercentOfRunningTotal(compareField);
                    break;
                case eShowDataAs.PercentOfParentRow:
                    dt.ShowDataAs.SetPercentParentRow();
                    break;
                case eShowDataAs.PercentOfParentColumn:
                    dt.ShowDataAs.SetPercentParentColumn();
                    break;
                case eShowDataAs.PercentOfParent:
                    dt.ShowDataAs.SetPercentParent(compareField);
                    break;
                case eShowDataAs.RankAscending:
                    dt.ShowDataAs.SetRankAscending(compareField);
                    break;
                case eShowDataAs.RankDescending:
                    dt.ShowDataAs.SetRankDescending(compareField);
                    break;
                case eShowDataAs.PercentOfRunningTotal:
                    dt.ShowDataAs.SetPercentOfRunningTotal(compareField);
                    break;
            }
        }

        [TestMethod]
        public void RunningTotalMissingColumnRowFieldsShouldNotThrow()
        {
            using (var p = OpenPackage("DivZeroShowValuesAs.xlsx", true))
            {
                var sheet = p.Workbook.Worksheets.Add("DataSheet");

                var table = sheet.Tables.Add(new ExcelAddress("A1:B5"), "TestTable");
                table.ShowHeader = true;

                sheet.Cells["A1"].Value = "Space";
                sheet.Cells["A2"].Value = 5;
                sheet.Cells["A3"].Value = 10;
                sheet.Cells["A4"].Value = 3;
                sheet.Cells["A5"].Value = 27;

                sheet.Cells["B1"].Value = "Time";
                sheet.Cells["B2"].Value = 120;
                sheet.Cells["B4"].Value = 90;
                sheet.Cells["B3"].Value = 150;
                sheet.Cells["B5"].Value = 3500;

                var ptWs = p.Workbook.Worksheets.Add("PTSheet");

                var pt = ptWs.PivotTables.Add(ptWs.Cells["A1"], table, "Pivot1");

                var dt = pt.DataFields.Add(pt.Fields["Time"]);
                var field2 = pt.Fields["Space"];

                dt.ShowDataAs.SetRunningTotal(field2);
                dt.Function = DataFieldFunctions.StdDev;

                pt.Calculate(true);

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void RunningTotalPercentMissingColumnRowFieldsShouldNotThrow()
        {
            using (var p = OpenPackage("DivZeroShowValuesAs.xlsx", true))
            {
                var sheet = p.Workbook.Worksheets.Add("DataSheet");

                var table = sheet.Tables.Add(new ExcelAddress("A1:B5"), "TestTable");
                table.ShowHeader = true;

                sheet.Cells["A1"].Value = "Space";
                sheet.Cells["A2"].Value = 5;
                sheet.Cells["A3"].Value = 10;
                sheet.Cells["A4"].Value = 3;
                sheet.Cells["A5"].Value = 27;

                sheet.Cells["B1"].Value = "Time";
                sheet.Cells["B2"].Value = 120;
                sheet.Cells["B4"].Value = 90;
                sheet.Cells["B3"].Value = 150;
                sheet.Cells["B5"].Value = 3500;

                var ptWs = p.Workbook.Worksheets.Add("PTSheet");

                var pt = ptWs.PivotTables.Add(ptWs.Cells["A1"], table, "Pivot1");

                var dt = pt.DataFields.Add(pt.Fields["Time"]);
                var field2 = pt.Fields["Space"];

                dt.ShowDataAs.SetPercentOfRunningTotal(field2);
                dt.Function = DataFieldFunctions.StdDev;

                pt.Calculate(true);

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void RunningTotalDivisonByZeroStdDev()
        {
            using (var p = OpenPackage("DivZeroShowValuesAsOther.xlsx", true))
            {
                var sheet = p.Workbook.Worksheets.Add("DataSheet");

                var table = sheet.Tables.Add(new ExcelAddress("A1:D5"), "TestTable");
                table.ShowHeader = true;
                sheet.Cells["A1"].Value = "Names";
                sheet.Cells["A2"].Value = "Michael Phelps";
                sheet.Cells["A3"].Value = "Kristin Otto";
                sheet.Cells["A4"].Value = "Aleksandr Popov";
                sheet.Cells["A5"].Value = "Mark Spitz";

                sheet.Cells["B1"].Value = "Swim Length(m)";
                sheet.Cells["B2"].Value = 200;
                sheet.Cells["B3"].Value = 300;
                sheet.Cells["B4"].Value = 100;
                sheet.Cells["B5"].Value = 500;

                sheet.Cells["C1"].Value = "Time(s)";
                sheet.Cells["C2"].Value = 120;
                sheet.Cells["C3"].Value = 150;
                sheet.Cells["C4"].Value = 90;
                sheet.Cells["C5"].Value = 3500;

                sheet.Cells["D1"].Value = "Age(yrs)";
                sheet.Cells["D2"].Formula = "2024-1985";
                sheet.Cells["D3"].Formula = "2024-1966";
                sheet.Cells["D4"].Formula = "2024-1971";
                sheet.Cells["D5"].Formula = "2024-1950";

                var ptWs = p.Workbook.Worksheets.Add("PTSheet");

                var pt = ptWs.PivotTables.Add(ptWs.Cells["A1"], table, "SwimmerPivotTable");

                pt.RowFields.Add(pt.Fields["Names"]);
                var age = pt.RowFields.Add(pt.Fields["Age(yrs)"]);

                var length = pt.DataFields.Add(pt.Fields["Swim Length(m)"]);
                var dt = pt.DataFields.Add(pt.Fields["Time(s)"]);

                pt.DataOnRows = false;

                sheet.Calculate();
                ptWs.Calculate();

                //Setting running total and then setting normal causes the issue
                //As basefield becomes 0
                dt.ShowDataAs.SetPercentOfRunningTotal(age);
                dt.ShowDataAs.SetNormal();

                dt.Function = DataFieldFunctions.StdDev;

                pt.Calculate(true);
                ptWs.Calculate();

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void TestAllDatafieldsWithAllShowDataOptions()
        {
            using (var p = OpenPackage("SwimmersPivotData.xlsx", true))
            {
                var sheet = p.Workbook.Worksheets.Add("DataSheet");

                var table = sheet.Tables.Add(new ExcelAddress("A1:D5"), "TestTable");
                table.ShowHeader = true;
                sheet.Cells["A1"].Value = "Names";
                sheet.Cells["A2"].Value = "Michael Phelps";
                sheet.Cells["A3"].Value = "Kristin Otto";
                sheet.Cells["A4"].Value = "Aleksandr Popov";
                sheet.Cells["A5"].Value = "Mark Spitz";

                sheet.Cells["B1"].Value = "Swim Length(m)";
                sheet.Cells["B2"].Value = 200;
                sheet.Cells["B3"].Value = 300;
                sheet.Cells["B4"].Value = 100;
                sheet.Cells["B5"].Value = 500;

                sheet.Cells["C1"].Value = "Time(s)";
                sheet.Cells["C2"].Value = 120;
                sheet.Cells["C3"].Value = 150;
                sheet.Cells["C4"].Value = 90;
                sheet.Cells["C5"].Value = 3500;

                sheet.Cells["D1"].Value = "Age(yrs)";
                sheet.Cells["D2"].Formula = "2024-1985";
                sheet.Cells["D3"].Formula = "2024-1966";
                sheet.Cells["D4"].Formula = "2024-1971";
                sheet.Cells["D5"].Formula = "2024-1950";

                var ptWs = p.Workbook.Worksheets.Add("PTSheet");

                var pt = ptWs.PivotTables.Add(ptWs.Cells["A1"], table, "SwimmerPivotTable");

                pt.RowFields.Add(pt.Fields["Names"]);
                var age = pt.RowFields.Add(pt.Fields["Age(yrs)"]);

                var length = pt.DataFields.Add(pt.Fields["Swim Length(m)"]);
                var dt = pt.DataFields.Add(pt.Fields["Time(s)"]);

                pt.DataOnRows = false;

                var functionValues = Enum.GetValues(typeof(DataFieldFunctions));
                var showAsValues = Enum.GetValues(typeof(eShowDataAs));

                for (int h = 0; h < showAsValues.Length; h++)
                {
                    var dataOption = (eShowDataAs)showAsValues.GetValue(h);

                    if (Enum.IsDefined(typeof(eShowDataAs), h) == false)
                    {
                        throw new InvalidCastException($"eShowDataAs has no value that matches {h}");
                    }

                    SetShowDataAs(dataOption, dt, age, ePrevNextPivotItem.Previous);

                    for (int i = 0; i < functionValues.Length; i++)
                    {
                        var tmpFunction = (DataFieldFunctions)functionValues.GetValue(i);

                        if (Enum.IsDefined(typeof(DataFieldFunctions), i) == false)
                        {
                            throw new InvalidCastException($"DataFieldFunctions has no value that matches {i}");
                        }

                        dt.Function = tmpFunction;

                        pt.Calculate(true);
                        ptWs.Calculate();

                        var addresses = pt.CalculatedItems[1].Index;
                        var values = pt.CalculatedItems[1].Values;

                        var range = pt.Address;

                        var num = 5 + 1 * (i + 1);
                        for (int j = 0; j < values.Count; j++)
                        {
                            ptWs.Cells[(j + 1) + 10 * h, num].Value = values[j];
                        }
                    }
                }
                SaveAndCleanup(p);
            }
        }
    }
}
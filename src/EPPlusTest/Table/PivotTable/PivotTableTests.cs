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
using System.Collections.Generic;
using System.IO;

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
            CreateDataSheet();
        }

        private static void CreateDataSheet()
        {
            var ws = _pck.Workbook.Worksheets.Add("Data");
            ws.Cells["K1"].Value = "Item";
            ws.Cells["L1"].Value = "Category";
            ws.Cells["M1"].Value = "Stock";
            ws.Cells["N1"].Value = "Price";
            ws.Cells["O1"].Value = "Date for grouping";

            ws.Cells["K2"].Value = "Crowbar";
            ws.Cells["L2"].Value = "Hardware";
            ws.Cells["M2"].Value = 12;
            ws.Cells["N2"].Value = 85.2;
            ws.Cells["O2"].Value = new DateTime(2010, 1, 31);

            ws.Cells["K3"].Value = "Crowbar";
            ws.Cells["L3"].Value = "Hardware";
            ws.Cells["M3"].Value = 15;
            ws.Cells["N3"].Value = 12.2;
            ws.Cells["O3"].Value = new DateTime(2010, 2, 28);

            ws.Cells["K4"].Value = "Hammer";
            ws.Cells["L4"].Value = "Hardware";
            ws.Cells["M4"].Value = 550;
            ws.Cells["N4"].Value = 72.7;
            ws.Cells["O4"].Value = new DateTime(2010, 3, 31);

            ws.Cells["K5"].Value = "Hammer";
            ws.Cells["L5"].Value = "Hardware";
            ws.Cells["M5"].Value = 120;
            ws.Cells["N5"].Value = 11.3;
            ws.Cells["O5"].Value = new DateTime(2010, 4, 30);

            ws.Cells["K6"].Value = "Crowbar";
            ws.Cells["L6"].Value = "Hardware";
            ws.Cells["M6"].Value = 120;
            ws.Cells["N6"].Value = 173.2;
            ws.Cells["O6"].Value = new DateTime(2010, 5, 31);

            ws.Cells["K7"].Value = "Hammer";
            ws.Cells["L7"].Value = "Hardware";
            ws.Cells["M7"].Value = 1;
            ws.Cells["N7"].Value = 4.2;
            ws.Cells["O7"].Value = new DateTime(2010, 6, 30);

            ws.Cells["K8"].Value = "Saw";
            ws.Cells["L8"].Value = "Hardware";
            ws.Cells["M8"].Value = 4;
            ws.Cells["N8"].Value = 33.12;
            ws.Cells["O8"].Value = new DateTime(2010, 6, 28);

            ws.Cells["K9"].Value = "Screwdriver";
            ws.Cells["L9"].Value = "Hardware";
            ws.Cells["M9"].Value = 1200;
            ws.Cells["N9"].Value = 45.2;
            ws.Cells["O9"].Value = new DateTime(2010, 8, 31);

            ws.Cells["K10"].Value = "Apple";
            ws.Cells["L10"].Value = "Groceries";
            ws.Cells["M10"].Value = 807;
            ws.Cells["N10"].Value = 1.2;
            ws.Cells["O10"].Value = new DateTime(2010, 9, 30);

            ws.Cells["K11"].Value = "Butter";
            ws.Cells["L11"].Value = "Groceries";
            ws.Cells["M11"].Value = 52;
            ws.Cells["N11"].Value = 7.2;
            ws.Cells["O11"].Value = new DateTime(2010, 10, 31);
            ws.Cells["O2:O11"].Style.Numberformat.Format = "yyyy-MM-dd";
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
            var ws=_pck.Workbook.Worksheets.Add("PivotSourceAddress");
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
            var pivotTable1 = ws.PivotTables.Add(ws.Cells["G1"], table , "PivotTable1");

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
            pt.Fields[5].Name = "Månad";
            pt.Fields[6].Name = "Kvartal";
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
            pt.RowFields.Add(pt.Fields["Date for grouping"]);

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
            var rf=pt.RowFields.Add(pt.Fields[0]);
            rf.SubTotalFunctions = eSubTotalFunctions.None;
            pt.DataOnRows = true;
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
            var pf2 = pt.PageFields.Add(pt.Fields[4]);
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
            pt.DataFields.Add(pt.Fields[3]);
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


    }
}

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
    public class PivotTableDataFieldShowDataAsTests : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenPackage("PivotTableShowDataAs.xlsx", true);
            var ws = _pck.Workbook.Worksheets.Add("Data");
            LoadItemData(ws);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }        
        [TestMethod]
        public void ShowAsPercentOfTotal()
        {
            var ws = _pck.Workbook.Worksheets.Add("ShowDataAsPercTot");

            LoadTestdata(ws);
            var tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTablePercTot");
            var rf = tbl.RowFields.Add(tbl.Fields[1]);
            var df = tbl.DataFields.Add(tbl.Fields[3]);
            df.Function = DataFieldFunctions.Sum;
            df.ShowDataAs.SetPercentOfTotal();
            tbl.DataOnRows = false;
            tbl.GridDropZones = false;

            Assert.AreEqual(eShowDataAs.PercentOfTotal, df.ShowDataAs.Value);

            tbl.Calculate();

        }
        [TestMethod]
        public void ShowAsPercentOfRow()
        {
            var ws = _pck.Workbook.Worksheets.Add("ShowDataAsPercRow");

            LoadTestdata(ws);
            var tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTablePercRow");
            var rf = tbl.RowFields.Add(tbl.Fields[1]);
            var df = tbl.DataFields.Add(tbl.Fields[3]);
            df.Function = DataFieldFunctions.Sum;
            df.ShowDataAs.SetPercentOfRow();
            tbl.DataOnRows = false;
            tbl.GridDropZones = false;

            Assert.AreEqual(eShowDataAs.PercentOfRow, df.ShowDataAs.Value);

            tbl.Calculate();

        }
        [TestMethod]
        public void ShowAsPercentOfCol()
        {
            var ws = _pck.Workbook.Worksheets.Add("ShowDataAsPercCol");

            LoadTestdata(ws);
            var tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTablePercCol");
            var rf = tbl.RowFields.Add(tbl.Fields[1]);
            var df = tbl.DataFields.Add(tbl.Fields[3]);
            df.Function = DataFieldFunctions.Sum;
            df.ShowDataAs.SetPercentOfColumn();
            tbl.DataOnRows = false;
            tbl.GridDropZones = false;

            Assert.AreEqual(eShowDataAs.PercentOfColumn, df.ShowDataAs.Value);

            tbl.Calculate();

        }
        [TestMethod]
        public void ShowAsPercent()
        {
            var ws = _pck.Workbook.Worksheets.Add("ShowDataAsPerc");

            LoadTestdata(ws);
            var tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTablePerc");
            var rf = tbl.RowFields.Add(tbl.Fields[1]);
            var df = tbl.DataFields.Add(tbl.Fields[3]);
            df.Function = DataFieldFunctions.Sum;
            rf.Items.Refresh();
            df.ShowDataAs.SetPercent(rf, 50);
            tbl.DataOnRows = false;
            tbl.GridDropZones = false;

            Assert.AreEqual(eShowDataAs.Percent, df.ShowDataAs.Value);
            Assert.AreEqual(rf.Index, df.BaseField);
            Assert.AreEqual(50, df.BaseItem);

            tbl.Calculate();

        }
        [TestMethod]
        public void ShowAsIndex()
        {
            var ws = _pck.Workbook.Worksheets.Add("ShowDataAsIndex");

            LoadTestdata(ws);
            var tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTableIndex");
            var rf = tbl.RowFields.Add(tbl.Fields[1]);
            var df = tbl.DataFields.Add(tbl.Fields[3]);
            df.Function = DataFieldFunctions.Sum;
            rf.Items.Refresh();
            df.ShowDataAs.SetIndex();
            tbl.DataOnRows = false;
            tbl.GridDropZones = false;

            Assert.AreEqual(eShowDataAs.Index, df.ShowDataAs.Value);

            tbl.Calculate();

        }
        [TestMethod]
        public void ShowAsDifference()
        {
            var ws = _pck.Workbook.Worksheets.Add("ShowDataAsDifference");

            LoadTestdata(ws);
            var tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTableDifference");
            var rf = tbl.RowFields.Add(tbl.Fields[1]);
            var df = tbl.DataFields.Add(tbl.Fields[3]);
            df.Function = DataFieldFunctions.Sum;
            rf.Items.Refresh();
            df.ShowDataAs.SetDifference(rf, ePrevNextPivotItem.Previous);
            tbl.DataOnRows = false;
            tbl.GridDropZones = false;

            Assert.AreEqual(eShowDataAs.Difference, df.ShowDataAs.Value);
            Assert.AreEqual(rf.Index, df.BaseField);
            Assert.AreEqual((int)ePrevNextPivotItem.Previous, df.BaseItem);

            tbl.Calculate();

            Assert.AreEqual(0D, tbl.CalculatedData.SelectField("NumValue", 2).GetValue("NumFormattedValue"));
            Assert.AreEqual(33D, tbl.CalculatedData.SelectField("NumValue", 3).GetValue("NumFormattedValue"));
            Assert.AreEqual(33D, tbl.CalculatedData.SelectField("NumValue", 100).GetValue("NumFormattedValue"));
        }
        [TestMethod]
        public void ShowAsPercentageDifference()
        {
            var ws = _pck.Workbook.Worksheets.Add("ShowDataAsPercDiff");

            LoadTestdata(ws);
            var tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTablePercDiff");
            var rf = tbl.RowFields.Add(tbl.Fields[1]);
            var df = tbl.DataFields.Add(tbl.Fields[3]);
            df.Function = DataFieldFunctions.Sum;
            rf.Items.Refresh();
            df.ShowDataAs.SetPercentageDifference(rf, ePrevNextPivotItem.Previous);
            tbl.DataOnRows = false;
            tbl.GridDropZones = false;

            Assert.AreEqual(eShowDataAs.PercentDifference, df.ShowDataAs.Value);
            Assert.AreEqual(rf.Index, df.BaseField);
            Assert.AreEqual((int)ePrevNextPivotItem.Previous, df.BaseItem);

            tbl.Calculate();

            Assert.AreEqual(0D, tbl.CalculatedData.SelectField("NumValue", 2).GetValue("NumFormattedValue"));
            Assert.AreEqual(0.5, tbl.CalculatedData.SelectField("NumValue", 3).GetValue("NumFormattedValue"));
            Assert.AreEqual(0.010101D, (double)tbl.CalculatedData.SelectField("NumValue", 100).GetValue("NumFormattedValue"), 0.00001);
        }

        [TestMethod]
        public void ShowAsRunningTotal()
        {
            var ws = _pck.Workbook.Worksheets.Add("ShowDataAsRunningTotal");

            LoadTestdata(ws);
            var tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTableRunningTotal");
            var rf = tbl.RowFields.Add(tbl.Fields[1]);
            var df = tbl.DataFields.Add(tbl.Fields[3]);
            df.Function = DataFieldFunctions.Sum;
            df.ShowDataAs.SetRunningTotal(rf);
            tbl.DataOnRows = false;
            tbl.GridDropZones = false;

            Assert.AreEqual(eShowDataAs.RunningTotal, df.ShowDataAs.Value);
            Assert.AreEqual(rf.Index, df.BaseField);

            tbl.Calculate();
        }

        [TestMethod]
        public void ShowAsPercentOfParent()
        {
            var ws = _pck.Workbook.Worksheets.Add("ShowDataAsPercentOfParent");

            LoadTestdata(ws);
            var tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTablePercParent");
            var rf = tbl.RowFields.Add(tbl.Fields[1]);
            var df = tbl.DataFields.Add(tbl.Fields[3]);
            df.Function = DataFieldFunctions.Sum;
            df.ShowDataAs.SetPercentParent(rf);
            tbl.DataOnRows = false;
            tbl.GridDropZones = false;

            Assert.AreEqual(eShowDataAs.PercentOfParent, df.ShowDataAs.Value);

            tbl.Calculate();
        }
        [TestMethod]
        public void ShowAsPercentOfParentRow()
        {
            var ws = _pck.Workbook.Worksheets.Add("ShowDataAsPercentOfParentRow");

            LoadTestdata(ws);
            var tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTablePercentOfParentRow");
            var rf = tbl.RowFields.Add(tbl.Fields[1]);
            var df = tbl.DataFields.Add(tbl.Fields[3]);
            df.Function = DataFieldFunctions.Sum;
            df.ShowDataAs.SetPercentParentRow();
            tbl.DataOnRows = false;
            tbl.GridDropZones = false;

            Assert.AreEqual(eShowDataAs.PercentOfParentRow, df.ShowDataAs.Value);

            tbl.Calculate();
        }
        [TestMethod]
        public void ShowAsPercentOfParentCol()
        {
            var ws = _pck.Workbook.Worksheets.Add("ShowDataAsPercentOfParentCol");

            LoadTestdata(ws);
            var tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTablePercentOfParentRow");
            var rf = tbl.RowFields.Add(tbl.Fields[1]);
            var df = tbl.DataFields.Add(tbl.Fields[3]);
            df.Function = DataFieldFunctions.Sum;
            df.ShowDataAs.SetPercentParentColumn();
            tbl.DataOnRows = false;
            tbl.GridDropZones = false;

            Assert.AreEqual(eShowDataAs.PercentOfParentColumn, df.ShowDataAs.Value);

            tbl.Calculate();
        }
        [TestMethod]
        public void ShowAsPercentOfRunningTotal()
        {
            var ws = _pck.Workbook.Worksheets.Add("ShowDataAsPercentOfRunningTotal");

            LoadTestdata(ws);
            var tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTablePercentOfParentRow");
            var rf = tbl.RowFields.Add(tbl.Fields[1]);
            var df = tbl.DataFields.Add(tbl.Fields[3]);
            df.ShowDataAs.SetPercentOfRunningTotal(rf);
            tbl.DataOnRows = false;
            tbl.GridDropZones = false;

            Assert.AreEqual(eShowDataAs.PercentOfRunningTotal, df.ShowDataAs.Value);

            tbl.Calculate();
        }
        [TestMethod]
        public void ShowAsRankAscending()
        {
            var ws = _pck.Workbook.Worksheets.Add("ShowDataAsRankAscending");

            LoadTestdata(ws);
            var tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTableRankAscending");
            var rf = tbl.RowFields.Add(tbl.Fields[1]);
            var df = tbl.DataFields.Add(tbl.Fields[3]);
            df.ShowDataAs.SetRankAscending(rf);
            tbl.DataOnRows = false;
            tbl.GridDropZones = false;

            Assert.AreEqual(eShowDataAs.RankAscending, df.ShowDataAs.Value);

            tbl.Calculate();
        }
        [TestMethod]
        public void ShowAsRankDescending()
        {
            var ws = _pck.Workbook.Worksheets.Add("ShowDataAsRankDescending");

            LoadTestdata(ws);
            var tbl = ws.PivotTables.Add(ws.Cells["F1"], ws.Cells["A1:D100"], "PivotTableRankDescending");
            var rf = tbl.RowFields.Add(tbl.Fields[1]);
            var df = tbl.DataFields.Add(tbl.Fields[3]);
            df.ShowDataAs.SetRankDescending(rf);
            tbl.DataOnRows = false;
            tbl.GridDropZones = false;

            Assert.AreEqual(eShowDataAs.RankDescending, df.ShowDataAs.Value);

            tbl.Calculate();
        }
    }
}

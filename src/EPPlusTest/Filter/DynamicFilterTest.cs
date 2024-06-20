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
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *******************************************************************************/
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Filter;

namespace EPPlusTest.Filter
{
    [TestClass]
    public class DynamicFilterTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("ValueFilter.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void AboveAverage()
        {
            var ws = _pck.Workbook.Worksheets.Add("AboveAverage");
            LoadTestdata(ws);
            SetDateValues(ws);

            ws.AutoFilterAddress = ws.Cells["A1:D100"];
            var col=ws.AutoFilter.Columns.AddDynamicFilterColumn(1);
            col.Type = eDynamicFilterType.AboveAverage;
            ws.AutoFilter.ApplyFilter();    
            Assert.AreEqual(true, ws.Row(48).Hidden);
            Assert.AreEqual(false, ws.Row(50).Hidden);
            Assert.AreEqual(false, ws.Row(51).Hidden);
            Assert.AreEqual(false, ws.Row(52).Hidden);
            Assert.AreEqual(true, ws.Row(53).Hidden);
        }
        [TestMethod]
        public void BelowAverage()
        {
            var ws = _pck.Workbook.Worksheets.Add("BelowAverage");
            LoadTestdata(ws);
            SetDateValues(ws);

            ws.AutoFilterAddress = ws.Cells["A1:D100"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(1);
            col.Type = eDynamicFilterType.BelowAverage;
            ws.AutoFilter.ApplyFilter();
            Assert.AreEqual(false, ws.Row(48).Hidden);
            Assert.AreEqual(true, ws.Row(50).Hidden);
            Assert.AreEqual(true, ws.Row(51).Hidden);
            Assert.AreEqual(true, ws.Row(52).Hidden);
            Assert.AreEqual(false, ws.Row(53).Hidden);
        }
        #region Day
        [TestMethod]
        public void Yesterday()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("Yesterday");
			var date = new DateTime(DateTime.Today.Year - 1, 11, 1);
			LoadTestdata(ws, 500, 1, 1, false, false, date);

			//Act
			ws.AutoFilterAddress = ws.Cells["A1:D500"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.Yesterday;
            ws.AutoFilter.ApplyFilter();
            
            //Assert
            var dt = DateTime.Today.AddDays(-1);
            var row = GetRowFromDate(dt, date);
            Assert.AreEqual(true, ws.Row(row - 1).Hidden);
            Assert.AreEqual(false, ws.Row(row).Hidden);
            Assert.AreEqual(true, ws.Row(row + 1).Hidden);
        }
        [TestMethod]
        public void Today()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("Today");
			var date = new DateTime(DateTime.Today.Year - 1, 11, 1);
			LoadTestdata(ws, 500, 1, 1, false, false, date);

			//Act
			ws.AutoFilterAddress = ws.Cells["A1:D500"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.Today;
            ws.AutoFilter.ApplyFilter();

            //Assert
            var dt = DateTime.Today;
            var row = GetRowFromDate(dt, date);
            Assert.AreEqual(true, ws.Row(row - 1).Hidden);
            Assert.AreEqual(false, ws.Row(row).Hidden);
            Assert.AreEqual(true, ws.Row(row + 1).Hidden);
        }
        [TestMethod]
        public void Tomorrow()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("Tomorrow");
			var date = new DateTime(DateTime.Today.Year - 1, 11, 1);
			LoadTestdata(ws, 500, 1, 1, false, false, date);

			//Act
			ws.AutoFilterAddress = ws.Cells["A1:D500"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.Tomorrow;
            ws.AutoFilter.ApplyFilter();

            //Assert
            var dt = DateTime.Today.AddDays(1);
            var row = GetRowFromDate(dt, date);
            Assert.AreEqual(true, ws.Row(row - 1).Hidden);
            Assert.AreEqual(false, ws.Row(row).Hidden);
            Assert.AreEqual(true, ws.Row(row + 1).Hidden);
        }

        #endregion
        #region Week
        [TestMethod]
        public void LastWeek()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("LastWeek");
			var date = new DateTime(DateTime.Today.Year - 1, 11, 1);
			LoadTestdata(ws, 500, 1, 1, false, false, date);

			//Act
			ws.AutoFilterAddress = ws.Cells["A1:D500"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.LastWeek;
            ws.AutoFilter.ApplyFilter();


            //Assert
            var dt = GetPrevSunday(DateTime.Today.AddDays(-7));
            var startRow = GetRowFromDate(dt, date);
            var endRow = startRow + 6;
            Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
            Assert.AreEqual(false, ws.Row(startRow).Hidden);
            Assert.AreEqual(false, ws.Row(endRow).Hidden);
            Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
        }
        [TestMethod]
        public void ThisWeek()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("ThisWeek");
			var date = new DateTime(DateTime.Today.Year - 1, 11, 1);
			LoadTestdata(ws, 500, 1, 1, false, false, date);

			//Act
			ws.AutoFilterAddress = ws.Cells["A1:D500"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.ThisWeek;
            ws.AutoFilter.ApplyFilter();

            //Assert
            var dt = GetPrevSunday(DateTime.Today);
            var startRow = GetRowFromDate(dt, date);
            var endRow = startRow + 6;
            Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
            Assert.AreEqual(false, ws.Row(startRow).Hidden);
            Assert.AreEqual(false, ws.Row(endRow).Hidden);
            Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
        }
        [TestMethod]
        public void NextWeek()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("NextWeek");
			var date = new DateTime(DateTime.Today.Year - 1, 11, 1);
			LoadTestdata(ws, 500, 1, 1, false, false, date);

			//Act
			ws.AutoFilterAddress = ws.Cells["A1:D500"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.NextWeek;
            ws.AutoFilter.ApplyFilter();

            //Assert
            var dt = GetPrevSunday(DateTime.Today.AddDays(7));
            var startRow = GetRowFromDate(dt, date);
            var endRow = startRow + 6;
            Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
            Assert.AreEqual(false, ws.Row(startRow).Hidden);
            Assert.AreEqual(false, ws.Row(endRow).Hidden);
            Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
        }
        #endregion
        #region Month
        [TestMethod]
        public void LastMonth()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("LastMonth");
			var date = new DateTime(DateTime.Today.Year - 1, 11, 1);
			LoadTestdata(ws, 500, 1, 1, false, false, date);

			//Act
			ws.AutoFilterAddress = ws.Cells["A1:D500"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.LastMonth;
            ws.AutoFilter.ApplyFilter();


            //Assert
            var dt = DateTime.Today.AddMonths(-1);
            var startRow = GetRowFromDate(new DateTime(dt.Year, dt.Month, 1), date);
            var endRow = GetRowFromDate(new DateTime(dt.Year, dt.Month, 1).AddMonths(1).AddDays(-1), date);
            Assert.AreEqual(true, ws.Row(startRow-1).Hidden);
            Assert.AreEqual(false, ws.Row(startRow).Hidden);
            Assert.AreEqual(false, ws.Row(endRow).Hidden);
            Assert.AreEqual(true, ws.Row(endRow+1).Hidden);
        }
        [TestMethod]
        public void ThisMonth()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("ThisMonth");
			var date = new DateTime(DateTime.Today.Year - 1, 11, 1);
			LoadTestdata(ws, 500, 1, 1, false, false, date);

            //Act
            ws.AutoFilterAddress = ws.Cells["A1:D500"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.ThisMonth;
            ws.AutoFilter.ApplyFilter();


            //Assert
            var dt = DateTime.Today;
            var startRow = GetRowFromDate(new DateTime(dt.Year, dt.Month, 1), date);
            var endRow = GetRowFromDate(new DateTime(dt.Year, dt.Month, 1).AddMonths(1).AddDays(-1), date);
            Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
            Assert.AreEqual(false, ws.Row(startRow).Hidden);
            Assert.AreEqual(false, ws.Row(endRow).Hidden);
            Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
        }
        [TestMethod]
        public void NextMonth()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("NextMonth");
            var startDate = DateTime.Today.AddMonths(-5);
            LoadTestdata(ws, 500, 1, 1, false, false, startDate);

            //Act
            ws.AutoFilterAddress = ws.Cells["A1:D500"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.NextMonth;
            ws.AutoFilter.ApplyFilter();


            //Assert
            var dt = DateTime.Today.AddMonths(1);
            var startRow = GetRowFromDate(new DateTime(dt.Year, dt.Month, 1), startDate);
            var endRow = GetRowFromDate(new DateTime(dt.Year, dt.Month, 1).AddMonths(1).AddDays(-1), startDate);
            Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
            Assert.AreEqual(false, ws.Row(startRow).Hidden);
            Assert.AreEqual(false, ws.Row(endRow).Hidden);
            Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
        }
        [TestMethod]
        public void M1()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("M1");
			var date = new DateTime(DateTime.Today.Year - 1, 11, 1);
			LoadTestdata(ws, 500, 1, 1, false, false, date);

			//Act
			ws.AutoFilterAddress = ws.Cells["A1:D500"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.M1;
            ws.AutoFilter.ApplyFilter();


            //Assert
            var dt = DateTime.Today.AddMonths(1);
            var startRow = GetRowFromDate(new DateTime(dt.Year, 1, 1), date);
            var endRow = GetRowFromDate(new DateTime(dt.Year, 1, 1).AddMonths(1).AddDays(-1), date);
            //Will only verify this year
            Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
            Assert.AreEqual(false, ws.Row(startRow).Hidden);
            Assert.AreEqual(false, ws.Row(endRow).Hidden);
            Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
        }
        [TestMethod]
        public void M2()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("M2");
			var date = new DateTime(DateTime.Today.Year - 1, 11, 1);
			LoadTestdata(ws, 500, 1, 1, false, false, date);

			//Act
			ws.AutoFilterAddress = ws.Cells["A1:D500"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.M2;
            ws.AutoFilter.ApplyFilter();


            //Assert
            var dt = DateTime.Today.AddMonths(1);
            var startRow = GetRowFromDate(new DateTime(dt.Year, 2, 1), date);
            var endRow = GetRowFromDate(new DateTime(dt.Year, 2, 1).AddMonths(1).AddDays(-1), date);
            //Will only verify this year
            Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
            Assert.AreEqual(false, ws.Row(startRow).Hidden);
            Assert.AreEqual(false, ws.Row(endRow).Hidden);
            Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);

        }
        [TestMethod]
        public void M3()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("M3");
			var date = new DateTime(DateTime.Today.Year - 1, 11, 1);
			LoadTestdata(ws, 500, 1, 1, false, false, date);

			//Act
			ws.AutoFilterAddress = ws.Cells["A1:D600"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.M3;
            ws.AutoFilter.ApplyFilter();


            //Assert
            var dt = DateTime.Today.AddMonths(1);
            var startRow = GetRowFromDate(new DateTime(dt.Year, 3, 1), date);
            var endRow = GetRowFromDate(new DateTime(dt.Year, 3, 1).AddMonths(1).AddDays(-1), date);
            //Will only verify this year
            Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
            Assert.AreEqual(false, ws.Row(startRow).Hidden);
            Assert.AreEqual(false, ws.Row(endRow).Hidden);
            Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);

        }
        [TestMethod]
        public void M4()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("M4");
			var date = new DateTime(DateTime.Today.Year - 1, 11, 1);
			LoadTestdata(ws, 500, 1, 1, false, false, date);

			//Act
			ws.AutoFilterAddress = ws.Cells["A1:D600"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.M4;
            ws.AutoFilter.ApplyFilter();


            //Assert
            var dt = DateTime.Today.AddMonths(1);
            var startRow = GetRowFromDate(new DateTime(dt.Year, 4, 1), date);
            var endRow = GetRowFromDate(new DateTime(dt.Year, 4, 1).AddMonths(1).AddDays(-1), date);
            //Will only verify this year
            Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
            Assert.AreEqual(false, ws.Row(startRow).Hidden);
            Assert.AreEqual(false, ws.Row(endRow).Hidden);
            Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);

        }
        [TestMethod]
        public void M5()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("M5");
			var date = new DateTime(DateTime.Today.Year - 1, 11, 1);
			LoadTestdata(ws, 500, 1, 1, false, false, date);

			//Act
			ws.AutoFilterAddress = ws.Cells["A1:D700"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.M5;
            ws.AutoFilter.ApplyFilter();


            //Assert
            var dt = DateTime.Today.AddMonths(1);
            var startRow = GetRowFromDate(new DateTime(dt.Year, 5, 1), date);
            var endRow = GetRowFromDate(new DateTime(dt.Year, 5, 1).AddMonths(1).AddDays(-1), date);
            //Will only verify this year
            Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
            Assert.AreEqual(false, ws.Row(startRow).Hidden);
            Assert.AreEqual(false, ws.Row(endRow).Hidden);
            Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);

        }
        [TestMethod]
        public void M6()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("M6");
			var date = new DateTime(DateTime.Today.Year - 1, 11, 1);
			LoadTestdata(ws, 500, 1, 1, false, false, date);

			//Act
			ws.AutoFilterAddress = ws.Cells["A1:D700"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.M6;
            ws.AutoFilter.ApplyFilter();


            //Assert
            var dt = DateTime.Today.AddMonths(1);
            var startRow = GetRowFromDate(new DateTime(dt.Year, 6, 1), date);
            var endRow = GetRowFromDate(new DateTime(dt.Year, 6, 1).AddMonths(1).AddDays(-1), date);
            //Will only verify this year
            Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
            Assert.AreEqual(false, ws.Row(startRow).Hidden);
            Assert.AreEqual(false, ws.Row(endRow).Hidden);
            Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
        }
        [TestMethod]
        public void M7()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("M7");
			var date = new DateTime(DateTime.Today.Year - 1, 11, 1);
			LoadTestdata(ws, 500, 1, 1, false, false, date);

			//Act
			ws.AutoFilterAddress = ws.Cells["A1:D700"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.M7;
            ws.AutoFilter.ApplyFilter();


            //Assert
            var dt = DateTime.Today.AddMonths(1);
            var startRow = GetRowFromDate(new DateTime(dt.Year, 7, 1), date);
            var endRow = GetRowFromDate(new DateTime(dt.Year, 7, 1).AddMonths(1).AddDays(-1), date);
            //Will only verify this year
            Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
            Assert.AreEqual(false, ws.Row(startRow).Hidden);
            Assert.AreEqual(false, ws.Row(endRow).Hidden);
            Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
        }
        [TestMethod]
        public void M8()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("M8");
			var date = new DateTime(DateTime.Today.Year - 1, 11, 1);
			LoadTestdata(ws, 500, 1, 1, false, false, date);

			//Act
			ws.AutoFilterAddress = ws.Cells["A1:D800"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.M8;
            ws.AutoFilter.ApplyFilter();


            //Assert
            var dt = DateTime.Today.AddMonths(1);
            var startRow = GetRowFromDate(new DateTime(dt.Year, 8, 1), date);
            var endRow = GetRowFromDate(new DateTime(dt.Year, 8, 1).AddMonths(1).AddDays(-1), date);
            //Will only verify this year
            Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
            Assert.AreEqual(false, ws.Row(startRow).Hidden);
            Assert.AreEqual(false, ws.Row(endRow).Hidden);
            Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
        }
        [TestMethod]
        public void M9()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("M9");
			var date = new DateTime(DateTime.Today.Year - 1, 11, 1);
			LoadTestdata(ws, 500, 1, 1, false, false, date);

			//Act
			ws.AutoFilterAddress = ws.Cells["A1:D800"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.M9;
            ws.AutoFilter.ApplyFilter();


            //Assert
            var dt = DateTime.Today.AddMonths(1);
            var startRow = GetRowFromDate(new DateTime(dt.Year, 9, 1), date);
            var endRow = GetRowFromDate(new DateTime(dt.Year, 9, 1).AddMonths(1).AddDays(-1), date);
            //Will only verify this year
            Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
            Assert.AreEqual(false, ws.Row(startRow).Hidden);
            Assert.AreEqual(false, ws.Row(endRow).Hidden);
            Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
        }
        [TestMethod]
        public void M10()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("M10");
			var date = new DateTime(DateTime.Today.Year - 1, 11, 1);
			LoadTestdata(ws, 500, 1, 1, false, false, date);

			//Act
			ws.AutoFilterAddress = ws.Cells["A1:D800"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.M10;
            ws.AutoFilter.ApplyFilter();


            //Assert
            var dt = DateTime.Today.AddMonths(1);
            var startRow = GetRowFromDate(new DateTime(dt.Year, 10, 1), date);
            var endRow = GetRowFromDate(new DateTime(dt.Year, 10, 1).AddMonths(1).AddDays(-1), date);
            //Will only verify this year
            Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
            Assert.AreEqual(false, ws.Row(startRow).Hidden);
            Assert.AreEqual(false, ws.Row(endRow).Hidden);
            Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
        }
        [TestMethod]
        public void M11()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("M11");
			var date = new DateTime(DateTime.Today.Year - 1, 11, 1);
			LoadTestdata(ws, 500, 1, 1, false, false, date);

			//Act
			ws.AutoFilterAddress = ws.Cells["A1:D800"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.M11;
            ws.AutoFilter.ApplyFilter();


            //Assert
            var dt = DateTime.Today.AddMonths(1);
            var startRow = GetRowFromDate(new DateTime(dt.Year, 11, 1), date);
            var endRow = GetRowFromDate(new DateTime(dt.Year, 11, 1).AddMonths(1).AddDays(-1), date);
            //Will only verify this year
            Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
            Assert.AreEqual(false, ws.Row(startRow).Hidden);
            Assert.AreEqual(false, ws.Row(endRow).Hidden);
            Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
        }
        [TestMethod]
        public void M12()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("M12");
			var date = new DateTime(DateTime.Today.Year - 1, 11, 1);
			LoadTestdata(ws, 500, 1, 1, false, false, date);

			//Act
			ws.AutoFilterAddress = ws.Cells["A1:D900"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.M12;
            ws.AutoFilter.ApplyFilter();


            //Assert
            var dt = DateTime.Today.AddMonths(1);
            var startRow = GetRowFromDate(new DateTime(dt.Year, 12, 1), date);
            var endRow = GetRowFromDate(new DateTime(dt.Year, 12, 1).AddMonths(1).AddDays(-1), date);
            //Will only verify this year
            Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
            Assert.AreEqual(false, ws.Row(startRow).Hidden);
            Assert.AreEqual(false, ws.Row(endRow).Hidden);
            Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
        }

        #endregion
        #region Quarter
        [TestMethod]
        public void LastQuarter()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("LastQuarter");
			var date = new DateTime(DateTime.Today.Year - 1, 11, 1);
			LoadTestdata(ws, 500, 1, 1, false, false, date);

			//Act
			ws.AutoFilterAddress = ws.Cells["A1:D500"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.LastQuarter;
            ws.AutoFilter.ApplyFilter();

            //Assert
            var startDate = GetStartOfQuarter(DateTime.Today.AddMonths(-3));
            var endDate = startDate.AddMonths(3).AddDays(-1);
            var startRow = GetRowFromDate(startDate, date);
            var endRow = GetRowFromDate(endDate, date);
            if (startRow > 2)
            {
                Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
            }
            Assert.AreEqual(false, ws.Row(startRow).Hidden);
            Assert.AreEqual(false, ws.Row(endRow).Hidden);
            Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
        }
        [TestMethod]
        public void ThisQuarter()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("ThisQuarter");
			var date = new DateTime(DateTime.Today.Year - 1, 11, 1);
			LoadTestdata(ws, 500, 1, 1, false, false, date);

			//Act   
			ws.AutoFilterAddress = ws.Cells["A1:D500"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.ThisQuarter;
            ws.AutoFilter.ApplyFilter();

            //Assert
            var startDate = GetStartOfQuarter(DateTime.Today);
            var endDate = startDate.AddMonths(3).AddDays(-1);
            var startRow = GetRowFromDate(startDate, date);
            var endRow = GetRowFromDate(endDate, date);
            Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
            Assert.AreEqual(false, ws.Row(startRow).Hidden);
            Assert.AreEqual(false, ws.Row(endRow).Hidden);
            Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
        }
        [TestMethod]
        public void NextQuarter()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("NextQuarter");
            var date = new DateTime(DateTime.Today.Year - 1, 11, 1);
			LoadTestdata(ws, 600, 1, 1, false, false, date);

            //Act   
            ws.AutoFilterAddress = ws.Cells["A1:D600"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.NextQuarter;
            ws.AutoFilter.ApplyFilter();

            //Assert
            var startDate = GetStartOfQuarter(DateTime.Today.AddMonths(3));
            var endDate = startDate.AddMonths(3).AddDays(-1);
            var startRow = GetRowFromDate(startDate, date);
            var endRow = GetRowFromDate(endDate, date );
            Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
            Assert.AreEqual(false, ws.Row(startRow).Hidden);
            Assert.AreEqual(false, ws.Row(endRow).Hidden);
            Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
        }
        [TestMethod]
        public void Q1()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("Q1");
			var date = new DateTime(DateTime.Today.Year - 1, 11, 1);
			LoadTestdata(ws, 600, 1, 1, false, false, date);

			//Act
			ws.AutoFilterAddress = ws.Cells["A1:D500"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.Q1;
            ws.AutoFilter.ApplyFilter();


			//Assert
			var dt = DateTime.Today;
			var startRow = GetRowFromDate(new DateTime(dt.Year, 1, 1), date);
            var endRow = GetRowFromDate(new DateTime(dt.Year, 3, 31), date);
            //Will only verify this year
            Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
            Assert.AreEqual(false, ws.Row(startRow).Hidden);
            Assert.AreEqual(false, ws.Row(endRow).Hidden);
            Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
        }
        [TestMethod]
        public void Q2()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("Q2");
			var date = new DateTime(DateTime.Today.Year - 1, 11, 1);
			LoadTestdata(ws, 600, 1, 1, false, false, date);

			//Act
			ws.AutoFilterAddress = ws.Cells["A1:D500"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.Q2;
            ws.AutoFilter.ApplyFilter();


            //Assert
            var dt = DateTime.Today;
            var startRow = GetRowFromDate(new DateTime(dt.Year, 4, 1), date);
            var endRow = GetRowFromDate(new DateTime(dt.Year, 6, 30), date);
            //Will only verify this year
            Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
            Assert.AreEqual(false, ws.Row(startRow).Hidden);
            Assert.AreEqual(false, ws.Row(endRow).Hidden);
            Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
        }
        [TestMethod]
        public void Q3()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("Q3");
			var date = new DateTime(DateTime.Today.Year - 1, 11, 1);
			LoadTestdata(ws, 600, 1, 1, false, false, date);

			//Act
			ws.AutoFilterAddress = ws.Cells["A1:D500"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.Q3;
            ws.AutoFilter.ApplyFilter();


			//Assert
			var dt = DateTime.Today;
			var startRow = GetRowFromDate(new DateTime(dt.Year, 7, 1), date);
            var endRow = GetRowFromDate(new DateTime(dt.Year, 9, 30), date);
            //Will only verify this year
            Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
            Assert.AreEqual(false, ws.Row(startRow).Hidden);
            Assert.AreEqual(false, ws.Row(endRow).Hidden);
            Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
        }
        [TestMethod]
        public void Q4()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("Q4");
			var date = new DateTime(DateTime.Today.Year - 1, 11, 1);
			LoadTestdata(ws, 600, 1, 1, false, false, date);

			//Act
			ws.AutoFilterAddress = ws.Cells["A1:D500"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.Q4;
            ws.AutoFilter.ApplyFilter();


			//Assert
			var dt = DateTime.Today;
			var startRow = GetRowFromDate(new DateTime(dt.Year, 10, 1), date);
            var endRow = GetRowFromDate(new DateTime(dt.Year, 12, 31), date);
            //Will only verify this year
            Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
            Assert.AreEqual(false, ws.Row(startRow).Hidden);
            Assert.AreEqual(false, ws.Row(endRow).Hidden);
            Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
        }

        #endregion
        #region Year
        [TestMethod]
        public void LastYear()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("LastYear");
			var date = new DateTime(DateTime.Today.Year - 1, 11, 1);
			LoadTestdata(ws, 500, 1, 1, false, false, date);

			//Act
			ws.AutoFilterAddress = ws.Cells["A1:D500"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.LastYear;
            ws.AutoFilter.ApplyFilter();

            //Assert
            var startDate = new DateTime(DateTime.Today.Year-1, 1, 1);
            var endDate = new DateTime(DateTime.Today.Year - 1, 12, 31);
            var startRow = GetRowFromDate(startDate, date);
            var endRow = GetRowFromDate(endDate, date);
            Assert.AreEqual(false, ws.Row(startRow).Hidden);
            Assert.AreEqual(false, ws.Row(endRow).Hidden);
            Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
        }
        [TestMethod]
        public void ThisYear()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("ThisYear");
			var date = new DateTime(DateTime.Today.Year - 1, 11, 1);
			LoadTestdata(ws, 500, 1, 1, false, false, date);

			//Act
			ws.AutoFilterAddress = ws.Cells["A1:D500"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.ThisYear;
            ws.AutoFilter.ApplyFilter();

            //Assert
            var startDate = new DateTime(DateTime.Today.Year , 1, 1);
            var endDate = new DateTime(DateTime.Today.Year, 12, 31);
            var startRow = GetRowFromDate(startDate, date);
            var endRow = GetRowFromDate(endDate, date);
            Assert.AreEqual(true, ws.Row(startRow-1).Hidden);
            Assert.AreEqual(false, ws.Row(startRow).Hidden);
            Assert.AreEqual(false, ws.Row(endRow).Hidden);
            Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
        }

        [TestMethod]
        public void NextYear()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("NextYear");
			var date = new DateTime(DateTime.Today.Year - 1, 11, 1);
			LoadTestdata(ws, 500, 1, 1, false, false, date);

			//Act
			ws.AutoFilterAddress = ws.Cells["A1:D500"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.NextYear;
            ws.AutoFilter.ApplyFilter();

            //Assert
            var startDate = new DateTime(DateTime.Today.Year + 1, 1, 1);
            var endDate = new DateTime(DateTime.Today.Year + 1, 12, 31);
            var startRow = GetRowFromDate(startDate, date);
            var endRow = GetRowFromDate(endDate, date);
            Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
            Assert.AreEqual(false, ws.Row(startRow).Hidden);
            Assert.AreEqual(false, ws.Row(500).Hidden);
            Assert.AreEqual(false, ws.Row(501).Hidden);
        }

        [TestMethod]
        public void YearToDate()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("YearToDate");
			var date = new DateTime(DateTime.Today.Year - 1, 11, 1);
			LoadTestdata(ws, 500, 1, 1, false, false, date);

			//Act
			ws.AutoFilterAddress = ws.Cells["A1:D500"];
            var col = ws.AutoFilter.Columns.AddDynamicFilterColumn(0);
            col.Type = eDynamicFilterType.YearToDate;
            ws.AutoFilter.ApplyFilter();

            //Assert
            var startDate = new DateTime(DateTime.Today.Year, 1, 1);
            var endDate = DateTime.Today;
            var startRow = GetRowFromDate(startDate, date);
            var endRow = GetRowFromDate(endDate, date   );
            Assert.AreEqual(true, ws.Row(startRow - 1).Hidden);
            Assert.AreEqual(false, ws.Row(startRow).Hidden);
            Assert.AreEqual(false, ws.Row(endRow).Hidden);
            Assert.AreEqual(true, ws.Row(endRow + 1).Hidden);
        }
        #endregion

        #region Private methods
        private DateTime GetStartOfQuarter(DateTime dt)
        {
            var quarter = ((dt.Month - (dt.Month - 1) % 3) + 1) / 3;
                      
            return new DateTime(dt.Year, (quarter * 3) + 1, 1);
        }
        private DateTime GetPrevSunday(DateTime dt)
        {
            while (dt.DayOfWeek != DayOfWeek.Sunday)
            {
                dt = dt.AddDays(-1);
            }
            return dt;
        }
        #endregion
    }
}

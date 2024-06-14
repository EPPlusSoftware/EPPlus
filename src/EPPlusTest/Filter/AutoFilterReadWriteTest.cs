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
    public class AutoFilterReadWriteTest : TestBase
    {        
        [TestMethod]
        public void ValuesFilter()
        {
            var pck=OpenPackage("AutoFilterValues.xlsx", true);
            var ws = pck.Workbook.Worksheets.Add("Values");
            LoadTestdata(ws);

            ws.AutoFilterAddress = ws.Cells["A1:D100"];
            var col=ws.AutoFilter.Columns.AddValueFilterColumn(1);
            col.Filters.Add("3");
            col.Filters.Add("6");
            col.Filters.Add("19");
            col.Filters.Blank = true;
            col.Filters.Add(new ExcelFilterDateGroupItem(2018, 12));
            
            var col2 = ws.AutoFilter.Columns.AddValueFilterColumn(2);
            col2.Filters.Add("Value 6");
            ws.AutoFilter.ApplyFilter();

            Assert.AreEqual(true, ws.Row(2).Hidden);
            Assert.AreEqual(false, ws.Row(6).Hidden);
            Assert.AreEqual(true, ws.Row(4).Hidden);
            Assert.AreEqual(true, ws.Row(100).Hidden);
            Assert.AreEqual(false, ws.Row(101).Hidden);
            ws.AutoFilter.Save();
            pck.Save();
            pck.Dispose();

            /*** Reopen and validate ***/
            pck = OpenPackage("AutoFilterValues.xlsx", false);
            ws = pck.Workbook.Worksheets["Values"];
            Assert.AreEqual(2, ws.AutoFilter.Columns.Count);
            Assert.AreEqual(4, ((ExcelValueFilterColumn)ws.AutoFilter.Columns[1]).Filters.Count);
            Assert.AreEqual(1, ((ExcelValueFilterColumn)ws.AutoFilter.Columns[2]).Filters.Count);
            pck.Dispose();
        }
        [TestMethod]
        public void TableValuesFilter()
        {
            var pck = OpenPackage("TableFilterValues.xlsx", true);
            var ws = pck.Workbook.Worksheets.Add("TableValues");
            LoadTestdata(ws);

            var tbl =ws.Tables.Add(ws.Cells["A1:D100"], "Table1");

            tbl.ShowFilter = true;
            var col = tbl.AutoFilter.Columns.AddValueFilterColumn(1);
            col.Filters.Add("3");
            col.Filters.Add("6");
            col.Filters.Add("19");
            col.Filters.Blank = true;
            col.Filters.Add(new ExcelFilterDateGroupItem(2018, 12));

            var col2 = tbl.AutoFilter.Columns.AddValueFilterColumn(2);
            col2.Filters.Add("Value 6");
            tbl.AutoFilter.ApplyFilter();

            Assert.AreEqual(true, ws.Row(2).Hidden);
            Assert.AreEqual(false, ws.Row(6).Hidden);
            Assert.AreEqual(true, ws.Row(4).Hidden);
            Assert.AreEqual(true, ws.Row(100).Hidden);
            Assert.AreEqual(false, ws.Row(101).Hidden);
            pck.Save();
            pck.Dispose();

            /*** Reopen and validate ***/
            pck=OpenPackage("TableFilterValues.xlsx", false);
            tbl = pck.Workbook.Worksheets["TableValues"].Tables[0];
            Assert.AreEqual(2, tbl.AutoFilter.Columns.Count);
            Assert.AreEqual(4, ((ExcelValueFilterColumn)tbl.AutoFilter.Columns[1]).Filters.Count);
            Assert.AreEqual(1, ((ExcelValueFilterColumn)tbl.AutoFilter.Columns[2]).Filters.Count);
            pck.Dispose();
        }
        [TestMethod]
        public void CustomFilter()
        {
            var pck = OpenPackage("AutoFilterCustom.xlsx", true);
            var ws = pck.Workbook.Worksheets.Add("Autofilter");
            LoadTestdata(ws);

            ws.AutoFilterAddress = ws.Cells["A1:D100"];
            var col = ws.AutoFilter.Columns.AddCustomFilterColumn(2);            
            col.And = true;
            col.Filters.Add(new ExcelFilterCustomItem("Val*"));
            col.Filters.Add(new ExcelFilterCustomItem("*3"));
            ws.AutoFilter.ApplyFilter();

            ws.AutoFilter.Save();
            Assert.AreEqual(true, ws.Row(2).Hidden);
            Assert.AreEqual(false, ws.Row(3).Hidden);
            Assert.AreEqual(true, ws.Row(52).Hidden);
            Assert.AreEqual(false, ws.Row(53).Hidden);
            Assert.AreEqual(false, ws.Row(101).Hidden);
            pck.Save();
            pck.Dispose();

            /*** Reopen and validate ***/
            pck = OpenPackage("AutoFilterCustom.xlsx", false);
            ws = pck.Workbook.Worksheets["Autofilter"];
            Assert.AreEqual(1, ws.AutoFilter.Columns.Count);
            Assert.AreEqual(2, ((ExcelCustomFilterColumn)ws.AutoFilter.Columns[2]).Filters.Count);
            ws.AutoFilter.Save();
            pck.Save();
        }
        [TestMethod]
        public void Top10Filter_100()
        {
            var pck = OpenPackage("AutoFilterTop10_100.xlsx", true);

            /*** Bottom 12 ***/
            var ws = pck.Workbook.Worksheets.Add("Bottom12");
            LoadTestdata(ws);
            ws.AutoFilterAddress = ws.Cells["A1:D100"];
            var col = ws.AutoFilter.Columns.AddTop10FilterColumn(1);
            col.Top = false;
            col.Value = 12;
            ws.AutoFilter.ApplyFilter();

            /*** Top 10 ***/
            ws = pck.Workbook.Worksheets.Add("Top10");
            LoadTestdata(ws);
            SetDateValues(ws);
            ws.AutoFilterAddress = ws.Cells["A1:D100"];
            col = ws.AutoFilter.Columns.AddTop10FilterColumn(1);
            col.Value = 10;
            col.Top = true;
            ws.AutoFilter.ApplyFilter();

            /*** Bottom 12 Percent ***/
            ws = pck.Workbook.Worksheets.Add("Bottom12Percent");
            LoadTestdata(ws);
            ws.AutoFilterAddress = ws.Cells["A1:D100"];
            col = ws.AutoFilter.Columns.AddTop10FilterColumn(1);
            col.Value = 12;
            col.Percent = true;
            col.Top = false;
            ws.AutoFilter.ApplyFilter();

            /*** Top 10 ***/
            ws = pck.Workbook.Worksheets.Add("Top10Percent");
            LoadTestdata(ws);
            SetDateValues(ws);
            ws.AutoFilterAddress = ws.Cells["A1:D100"];
            col = ws.AutoFilter.Columns.AddTop10FilterColumn(1);
            col.Value = 10;
            col.Top = true;
            col.Percent = true;
            ws.AutoFilter.ApplyFilter();

            pck.Save();
            pck.Dispose();

            /*** Reopen and validate ***/
            pck=OpenPackage("AutoFilterTop10_100.xlsx", false);
            ws = pck.Workbook.Worksheets["Bottom12"];
            var top10Col = (ExcelTop10FilterColumn)ws.AutoFilter.Columns[1];
            Assert.AreEqual(13, top10Col.FilterValue);

            ws = pck.Workbook.Worksheets["Top10"];
            top10Col = (ExcelTop10FilterColumn)ws.AutoFilter.Columns[1];
            Assert.AreEqual(94, top10Col.FilterValue);

            ws = pck.Workbook.Worksheets["Bottom12Percent"];
            top10Col = (ExcelTop10FilterColumn)ws.AutoFilter.Columns[1];
            Assert.AreEqual(12, top10Col.FilterValue);

            ws = pck.Workbook.Worksheets["Top10Percent"];
            top10Col = (ExcelTop10FilterColumn)ws.AutoFilter.Columns[1];
            Assert.AreEqual(95, top10Col.FilterValue);

            pck.Dispose();
        }
        [TestMethod]
        public void Top10Filter_500()
        {
            var pck = OpenPackage("AutoFilterTop10_500.xlsx", true);

            /*** Bottom 12 ***/
            var ws = pck.Workbook.Worksheets.Add("Bottom12");
            LoadTestdata(ws, 500);
            SetDateValues(ws);
            ws.AutoFilterAddress = ws.Cells["A1:D500"];
            var col = ws.AutoFilter.Columns.AddTop10FilterColumn(1);
            col.Top = false;
            col.Value = 12;
            ws.AutoFilter.ApplyFilter();

            /*** Top 10 ***/
            ws = pck.Workbook.Worksheets.Add("Top10");
            LoadTestdata(ws, 500);
            SetDateValues(ws);
            ws.AutoFilterAddress = ws.Cells["A1:D500"];
            col = ws.AutoFilter.Columns.AddTop10FilterColumn(1);
            col.Value = 10;
            col.Top = true;
            ws.AutoFilter.ApplyFilter();

            /*** Bottom 12 Percent ***/
            ws = pck.Workbook.Worksheets.Add("Bottom12Percent");
            LoadTestdata(ws, 500);
            SetDateValues(ws);
            ws.AutoFilterAddress = ws.Cells["A1:D500"];
            col = ws.AutoFilter.Columns.AddTop10FilterColumn(1);
            col.Value = 12;
            col.Top = false;
            col.Percent = true;
            ws.AutoFilter.ApplyFilter();

            /*** Top 10 ***/
            ws = pck.Workbook.Worksheets.Add("Top10Percent");
            LoadTestdata(ws, 500);
            SetDateValues(ws);
            ws.AutoFilterAddress = ws.Cells["A1:D500"];
            col = ws.AutoFilter.Columns.AddTop10FilterColumn(1);
            col.Value = 10;
            col.Top = true;
            col.Percent = true;
            ws.AutoFilter.ApplyFilter();

            pck.Save();
            pck.Dispose();

            /*** Reopen and validate ***/
            pck=OpenPackage("AutoFilterTop10_500.xlsx", false);
            ws = pck.Workbook.Worksheets["Bottom12"];
            var top10Col = (ExcelTop10FilterColumn)ws.AutoFilter.Columns[1];
            Assert.AreEqual(13, top10Col.FilterValue);

            ws = pck.Workbook.Worksheets["Top10"];
            top10Col = (ExcelTop10FilterColumn)ws.AutoFilter.Columns[1];
            Assert.AreEqual(494, top10Col.FilterValue);

            ws = pck.Workbook.Worksheets["Bottom12Percent"];
            top10Col = (ExcelTop10FilterColumn)ws.AutoFilter.Columns[1];
            Assert.AreEqual(63, top10Col.FilterValue);

            ws = pck.Workbook.Worksheets["Top10Percent"];
            top10Col = (ExcelTop10FilterColumn)ws.AutoFilter.Columns[1];
            Assert.AreEqual(455, top10Col.FilterValue);

            pck.Dispose();
        }

        [TestMethod]
        public void Top10Filter_733()
        {
            var pck = OpenPackage("AutoFilterTop10_733.xlsx", true);

            /*** Bottom 12 ***/
            var ws = pck.Workbook.Worksheets.Add("Bottom12");
            LoadTestdata(ws, 733);
            SetDateValues(ws);
            ws.AutoFilterAddress = ws.Cells["A1:D733"];
            var col = ws.AutoFilter.Columns.AddTop10FilterColumn(1);
            col.Value = 12;
            col.Top = false;
            ws.AutoFilter.ApplyFilter();

            /*** Top 10 ***/
            ws = pck.Workbook.Worksheets.Add("Top10");
            LoadTestdata(ws, 733);
            SetDateValues(ws);
            ws.AutoFilterAddress = ws.Cells["A1:D733"];
            col = ws.AutoFilter.Columns.AddTop10FilterColumn(1);
            col.Value = 10;
            ws.AutoFilter.ApplyFilter();

            /*** Bottom 12 Percent ***/
            ws = pck.Workbook.Worksheets.Add("Bottom12Percent");
            LoadTestdata(ws, 733);
            SetDateValues(ws);
            ws.AutoFilterAddress = ws.Cells["A1:D733"];
            col = ws.AutoFilter.Columns.AddTop10FilterColumn(1);
            col.Value = 12;
            col.Percent = true;
            col.Top = false;
            ws.AutoFilter.ApplyFilter();

            /*** Top 10 ***/
            ws = pck.Workbook.Worksheets.Add("Top10Percent");
            LoadTestdata(ws, 733);
            SetDateValues(ws);
            ws.AutoFilterAddress = ws.Cells["A1:D733"];
            col = ws.AutoFilter.Columns.AddTop10FilterColumn(1);
            col.Value = 10;
            col.Percent = true;
            ws.AutoFilter.ApplyFilter();

            pck.Save();
            pck.Dispose();

            /*** Reopen and validate ***/
            pck = OpenPackage("AutoFilterTop10_733.xlsx", false);
            ws = pck.Workbook.Worksheets["Bottom12"];
            var top10Col = (ExcelTop10FilterColumn)ws.AutoFilter.Columns[1];
            Assert.AreEqual(13, top10Col.FilterValue);

            ws = pck.Workbook.Worksheets["Top10"];
            top10Col = (ExcelTop10FilterColumn)ws.AutoFilter.Columns[1];
            Assert.AreEqual(727, top10Col.FilterValue);

            ws = pck.Workbook.Worksheets["Bottom12Percent"];
            top10Col = (ExcelTop10FilterColumn)ws.AutoFilter.Columns[1];
            Assert.AreEqual(91, top10Col.FilterValue);

            ws = pck.Workbook.Worksheets["Top10Percent"];
            top10Col = (ExcelTop10FilterColumn)ws.AutoFilter.Columns[1];
            Assert.AreEqual(664, top10Col.FilterValue);

            pck.Dispose();
        }
        [TestMethod]
        public void ColorFilter()
        {
            var pck = OpenPackage("ColorFilter.xlsx", true);

            /*** Bottom 12 ***/
            var ws = pck.Workbook.Worksheets.Add("ColorFilter");
            LoadTestdata(ws, 100);
            ws.AutoFilterAddress = ws.Cells["A1:D733"];
            var col = ws.AutoFilter.Columns.AddColorFilterColumn(1);            
        }
        [TestMethod]
        public void IconFilter()
        {
            var pck = OpenPackage("Iconfilter.xlsx", true);

            /*** Bottom 12 ***/
            var ws = pck.Workbook.Worksheets.Add("IconFilter");
            LoadTestdata(ws, 100);
            ws.AutoFilterAddress = ws.Cells["A1:D733"];
            var col = ws.AutoFilter.Columns.AddIconFilterColumn(1);
            col.IconSet = OfficeOpenXml.ConditionalFormatting.eExcelconditionalFormattingIconsSetType.ThreeTrafficLights1;
            col.IconId = 1;
        }

        [TestMethod]
        public void ValuesFilterBlankOnlyTest()
        {
            var pck = OpenPackage("AutofilterValuesWithBlanks.xlsx", true);
            var ws = pck.Workbook.Worksheets.Add("Values");
            LoadTestdata(ws);

            ws.Cells["B3"].Value = null;
            ws.Cells["B10"].Value = "";

            ws.Cells["C3"].Value = "";

            ws.AutoFilterAddress = ws.Cells["A1:D100"];
            var col = ws.AutoFilter.Columns.AddValueFilterColumn(1);
            col.Filters.Blank = true;

            var col2 = ws.AutoFilter.Columns.AddValueFilterColumn(2);
            col2.Filters.Blank = true;

            ws.AutoFilter.ApplyFilter();

            Assert.IsFalse(ws.Row(3).Hidden);

            for (int i = 2; i < 100; i++)
            {
                if (i != 3)
                {
                    Assert.IsTrue(ws.Row(i).Hidden);
                }
            }

            SaveAndCleanup(pck);
        }
    }
}

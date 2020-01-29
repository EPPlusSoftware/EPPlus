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
using System.Globalization;
using System.Threading;

namespace EPPlusTest.Filter
{
    [TestClass]
    public class CustomFilter : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("CustomFilter.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            _pck.Save();
            _pck.Dispose();
        }

        [TestMethod]
        public void CustomEndWith()
        {
            var ws = _pck.Workbook.Worksheets.Add("CustomEndWith");
            LoadTestdata(ws);
            
            ws.AutoFilterAddress = ws.Cells["A1:D100"];
            var col=ws.AutoFilter.Columns.AddCustomFilterColumn(2);
            col.Filters.Add(new ExcelFilterCustomItem("*3"));
            ws.AutoFilter.ApplyFilter();

            Assert.AreEqual(true, ws.Row(2).Hidden);
            Assert.AreEqual(false, ws.Row(3).Hidden);
        }
        [TestMethod]
        public void CustomStartsWithOrContainsText()
        {
            var ws = _pck.Workbook.Worksheets.Add("StartOrContains");
            LoadTestdata(ws);

            ws.AutoFilterAddress = ws.Cells["A1:D100"];
            var col = ws.AutoFilter.Columns.AddCustomFilterColumn(2);
            col.Filters.Add(new ExcelFilterCustomItem("*3"));
            col.Filters.Add(new ExcelFilterCustomItem("*ue ?2"));
            col.And = false;
            ws.AutoFilter.ApplyFilter();

            Assert.AreEqual(true, ws.Row(2).Hidden);
            Assert.AreEqual(false, ws.Row(3).Hidden);
            Assert.AreEqual(false, ws.Row(22).Hidden);
            Assert.AreEqual(false, ws.Row(33).Hidden);
        }
        [TestMethod]
        public void CustomNumericEqualOrGreaterThanOrEqual()
        {
            var ws = _pck.Workbook.Worksheets.Add("NumberEqOrGrEq");
            LoadTestdata(ws);

            ws.AutoFilterAddress = ws.Cells["A1:D100"];
            var col = ws.AutoFilter.Columns.AddCustomFilterColumn(1);
            col.Filters.Add(new ExcelFilterCustomItem("14"));
            col.Filters.Add(new ExcelFilterCustomItem("95", eFilterOperator.GreaterThanOrEqual));
            col.And = false;
            ws.AutoFilter.ApplyFilter();

            Assert.AreEqual(true, ws.Row(13).Hidden);
            Assert.AreEqual(false, ws.Row(14).Hidden);
            Assert.AreEqual(true, ws.Row(94).Hidden);
            Assert.AreEqual(false, ws.Row(95).Hidden);
            Assert.AreEqual(false, ws.Row(96).Hidden);
        }
        [TestMethod]
        public void CustomNumericEqualOrLessThanOrEqual()
        {
            var ws = _pck.Workbook.Worksheets.Add("NumberEqOrLessEq");
            LoadTestdata(ws);

            ws.AutoFilterAddress = ws.Cells["A1:D100"];
            var col = ws.AutoFilter.Columns.AddCustomFilterColumn(1);
            col.Filters.Add(new ExcelFilterCustomItem("14"));
            col.Filters.Add(new ExcelFilterCustomItem("12.3", eFilterOperator.LessThanOrEqual));
            col.And = false;
            ws.AutoFilter.ApplyFilter();

            Assert.AreEqual(false, ws.Row(2).Hidden);
            Assert.AreEqual(false, ws.Row(12).Hidden);
            Assert.AreEqual(true, ws.Row(13).Hidden);
            Assert.AreEqual(false, ws.Row(14).Hidden);
        }
        [TestMethod]
        public void CustomNumericEqualAndLessThanOrEqual()
        {
            var ws = _pck.Workbook.Worksheets.Add("NumberEqAndLess");
            LoadTestdata(ws);

            ws.AutoFilterAddress = ws.Cells["A1:D100"];
            var col = ws.AutoFilter.Columns.AddCustomFilterColumn(1);
            col.Filters.Add(new ExcelFilterCustomItem("13"));
            col.Filters.Add(new ExcelFilterCustomItem("12", eFilterOperator.LessThan));
            col.And = true;
            ws.AutoFilter.ApplyFilter();

            Assert.AreEqual(true, ws.Row(2).Hidden);
            Assert.AreEqual(true, ws.Row(12).Hidden);
            Assert.AreEqual(true, ws.Row(13).Hidden);
            Assert.AreEqual(true, ws.Row(14).Hidden);
        }
        [TestMethod]
        public void CustomNumericEqualAndNotEqual()
        {
            var ws = _pck.Workbook.Worksheets.Add("NumberGtAndNotEq");
            LoadTestdata(ws);

            ws.AutoFilterAddress = ws.Cells["A1:D100"];
            var col = ws.AutoFilter.Columns.AddCustomFilterColumn(1);
            col.Filters.Add(new ExcelFilterCustomItem("94", eFilterOperator.GreaterThan));
            col.Filters.Add(new ExcelFilterCustomItem("98", eFilterOperator.NotEqual));
            col.And = true;
            ws.AutoFilter.ApplyFilter();

            Assert.AreEqual(true, ws.Row(93).Hidden);
            Assert.AreEqual(true, ws.Row(94).Hidden);
            Assert.AreEqual(false, ws.Row(95).Hidden);
            Assert.AreEqual(false, ws.Row(97).Hidden);
            Assert.AreEqual(true, ws.Row(98).Hidden);
            Assert.AreEqual(false, ws.Row(99).Hidden);
        }
    }
}

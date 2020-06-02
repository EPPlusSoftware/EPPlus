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
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace EPPlusTest.Drawing.Chart
{
    [TestClass]
    public class ExcelStockChartTest : StockChartTestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("Stock.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }        
        [TestMethod]
        public void ReadStockVHLC()
        {
            using(var p=OpenTemplatePackage("StockVHLC.xlsx"))
            {
                var c = p.Workbook.Worksheets[0].Drawings[0];
                SaveWorkbook("StockVHLCSaved.xlsx", p);
            }
        }
        [TestMethod]
        public void AddStockHLCText()
        {
            var ws = _pck.Workbook.Worksheets.Add("StockTextHLC");
            LoadStockChartDataText(ws);
            
            var chart = ws.Drawings.AddStockChart("StockPeriodHLC", ws.Cells["A2:A7"], ws.Cells["D2:D7"], ws.Cells["E2:E7"], ws.Cells["F2:F7"]);
            chart.Series[0].HeaderAddress = ws.Cells["D1"];
            chart.Series[1].HeaderAddress = ws.Cells["E1"];
            chart.Series[2].HeaderAddress = ws.Cells["F1"];
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);
            chart.YAxis.AddGridlines();
            Assert.AreEqual(eChartType.StockHLC, chart.ChartType);
        }
        [TestMethod]
        public void AddStockHLCPeriod()
        {
            var ws = _pck.Workbook.Worksheets.Add("StockPeriodHLC");
            LoadStockChartDataPeriod(ws);

            var chart = ws.Drawings.AddStockChart("StockPeriodHLC",  ws.Cells["A1:A7"], ws.Cells["D1:D7"], ws.Cells["E1:E7"], ws.Cells["F1:F7"]);
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);
            Assert.AreEqual(eChartType.StockHLC, chart.ChartType);
        }
        [TestMethod]
        public void AddStockOHLCText()
        {
            var ws = _pck.Workbook.Worksheets.Add("StockTextOHLC");
            LoadStockChartDataText(ws);

            var chart = ws.Drawings.AddStockChart("StockTextOHLC", ws.Cells["A1:A7"], ws.Cells["D1:D7"], ws.Cells["E1:E7"], ws.Cells["F1:F7"], ws.Cells["C1:C7"]);
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);
            Assert.AreEqual(eChartType.StockOHLC, chart.ChartType);
        }
        [TestMethod]
        public void AddStockOHLCPeriod()
        {
            var ws = _pck.Workbook.Worksheets.Add("StockPeriodOHLC");
            LoadStockChartDataPeriod(ws);

            var chart = ws.Drawings.AddStockChart("StockPeriodOHLC", ws.Cells["A1:A7"], ws.Cells["D1:D7"], ws.Cells["E1:E7"], ws.Cells["F1:F7"], ws.Cells["C1:C7"]);
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);
            Assert.AreEqual(eChartType.StockOHLC, chart.ChartType);
        }
        [TestMethod]
        public void AddStockVHLCPeriod()
        {
            var ws = _pck.Workbook.Worksheets.Add("StockPeriodVHLC");
            LoadStockChartDataPeriod(ws);

            var chart = ws.Drawings.AddStockChart("StockPeriodVHLC", ws.Cells["A1:A7"], ws.Cells["D1:D7"], ws.Cells["E1:E7"], ws.Cells["F1:F7"], null, ws.Cells["B1:B7"]);
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);
            Assert.AreEqual(eChartType.StockVHLC, chart.ChartType);
        }
        [TestMethod]
        public void AddStockVHLCText()
        {
            var ws = _pck.Workbook.Worksheets.Add("StockTextVHLC");
            LoadStockChartDataText(ws);

            var chart = ws.Drawings.AddStockChart("StockTextVHLC", ws.Cells["A1:A7"], ws.Cells["D1:D7"], ws.Cells["E1:E7"], ws.Cells["F1:F7"], null, ws.Cells["B1:B7"]);
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);
            Assert.AreEqual(eChartType.StockVHLC, chart.ChartType);
        }
        [TestMethod]
        public void AddStockVOHLCPeriod()
        {
            var ws = _pck.Workbook.Worksheets.Add("StockPeriodVOHLC");
            LoadStockChartDataPeriod(ws);

            var chart = ws.Drawings.AddStockChart("StockPeriodVOHLC", ws.Cells["A1:A7"], ws.Cells["D1:D7"], ws.Cells["E1:E7"], ws.Cells["F1:F7"], ws.Cells["C1:C7"], ws.Cells["B1:B7"]);
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);
            Assert.AreEqual(eChartType.StockVOHLC, chart.ChartType);
        }

        [TestMethod]
        public void AddStockVOHLCText()
        {
            var ws = _pck.Workbook.Worksheets.Add("StockTextVOHLC");
            LoadStockChartDataText(ws);

            var chart = ws.Drawings.AddStockChart("StockTextVOHLC", ws.Cells["A1:A7"], ws.Cells["D1:D7"], ws.Cells["E1:E7"], ws.Cells["F1:F7"], ws.Cells["C1:C7"], ws.Cells["B1:B7"]);
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);
            Assert.AreEqual(eChartType.StockVOHLC, chart.ChartType);
        }

        [TestMethod]
        public void AddStockWithDataTable()
        {
            var ws = _pck.Workbook.Worksheets.Add("StockTextVOHLCDTable");
            LoadStockChartDataText(ws);

            var chart = ws.Drawings.AddStockChart("StockTextVOHLCDTable", ws.Cells["A1:A7"], ws.Cells["D1:D7"], ws.Cells["E1:E7"], ws.Cells["F1:F7"], ws.Cells["C1:C7"], ws.Cells["B1:B7"]);
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);
            chart.PlotArea.CreateDataTable();
            Assert.AreEqual(eChartType.StockVOHLC, chart.ChartType);
            Assert.IsNotNull(chart.PlotArea.DataTable);
        }
        [TestMethod]
        public void AddStockWithTrendLines()
        {
            var ws = _pck.Workbook.Worksheets.Add("StockTextVOHLCTrendLines");
            LoadStockChartDataText(ws);

            var chart = ws.Drawings.AddStockChart("StockTextVOHLCTrendLines", ws.Cells["A1:A7"], ws.Cells["D1:D7"], ws.Cells["E1:E7"], ws.Cells["F1:F7"], ws.Cells["C1:C7"], ws.Cells["B1:B7"]);
            chart.StyleManager.SetChartStyle(OfficeOpenXml.Drawing.Chart.Style.ePresetChartStyle.StockChartStyle9);
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);
            chart.Series[1].TrendLines.Add(eTrendLine.Linear);
            Assert.AreEqual(eChartType.StockVOHLC, chart.ChartType);
            Assert.AreEqual(1, chart.Series[1].TrendLines.Count);
        }
        [TestMethod]
        public void AddStockWithGridLines()
        {
            var ws = _pck.Workbook.Worksheets.Add("StockTextVOHLCGridLines");
            LoadStockChartDataText(ws);

            var chart = ws.Drawings.AddStockChart("StockTextVOHLCGridLines", ws.Cells["A1:A7"], ws.Cells["D1:D7"], ws.Cells["E1:E7"], ws.Cells["F1:F7"], ws.Cells["C1:C7"], ws.Cells["B1:B7"]);
            chart.StyleManager.SetChartStyle(OfficeOpenXml.Drawing.Chart.Style.ePresetChartStyle.StockChartStyle9);
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);
            chart.XAxis.AddGridlines(true, true);
            chart.YAxis.AddGridlines(true, true);
            chart.Axis[2].AddGridlines(true, true);
            Assert.AreEqual(eChartType.StockVOHLC, chart.ChartType);
        }
        [TestMethod]
        public void AddStockWithDataLabels()
        {
            var ws = _pck.Workbook.Worksheets.Add("StockTextVOHLCDatalabels");
            LoadStockChartDataText(ws);

            var chart = ws.Drawings.AddStockChart("StockTextVOHLCDatalabels", ws.Cells["A1:A7"], ws.Cells["D1:D7"], ws.Cells["E1:E7"], ws.Cells["F1:F7"], ws.Cells["C1:C7"], ws.Cells["B1:B7"]);
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);
            chart.DataLabel.ShowValue = true;
            var dl = chart.Series[0].DataLabel.DataLabels.Add(0);
            dl.ShowSeriesName = true;
            dl.ShowCategory = true;
            dl.Effect.SetPresetShadow(OfficeOpenXml.Drawing.ePresetExcelShadowType.OuterCenter);
            Assert.AreEqual(eChartType.StockVOHLC, chart.ChartType);
        }
    }
}

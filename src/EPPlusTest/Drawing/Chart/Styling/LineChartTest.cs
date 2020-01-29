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
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Xml;

namespace EPPlusTest.Drawing.Chart
{
    [TestClass]
    public class LineChartTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("LineChart.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            _pck.Save();
            _pck.Dispose();
        }
        [TestMethod]
        public void DropLines()
        {
            var ws = _pck.Workbook.Worksheets.Add("DropLines");
            LoadTestdata(ws);

            var chart = AddLine(ws, eLineChartType.Line, "line1", 0, 0);
            var dl=chart.AddDropLines();
            dl.Border.Fill.Style = eFillStyle.SolidFill;
            dl.Border.Fill.SolidFill.Color.SetSchemeColor(eSchemeColor.Accent2);

            Assert.IsNotNull(chart.DropLine);
            Assert.AreEqual(eFillStyle.SolidFill, dl.Border.Fill.Style);
            Assert.AreEqual(eSchemeColor.Accent2, dl.Border.Fill.SolidFill.Color.SchemeColor.Color);
        }
        [TestMethod]
        public void UpDownBars()
        {
            var ws = _pck.Workbook.Worksheets.Add("UpDownBars");
            LoadTestdata(ws);

            var chart = AddLine(ws, eLineChartType.Line, "line1", 0, 0);
            chart.Series.Add("B2:B50", "D2:D50");
            chart.Series.Add("C2:C50", "D2:D50");
            chart.AddUpDownBars();
            chart.UpDownBarGapWidth = 4;
            chart.DownBar.Border.Fill.Style = eFillStyle.SolidFill;
            chart.DownBar.Border.Fill.SolidFill.Color.SetPresetColor(ePresetColor.DarkRed);
            chart.UpBar.Border.Fill.Style = eFillStyle.SolidFill;
            chart.UpBar.Border.Fill.SolidFill.Color.SetSystemColor(eSystemColor.CaptionText);

            Assert.AreEqual(eFillStyle.SolidFill, chart.DownBar.Border.Fill.Style);
            Assert.AreEqual(ePresetColor.DarkRed, chart.DownBar.Border.Fill.SolidFill.Color.PresetColor.Color);

            Assert.AreEqual(eFillStyle.SolidFill, chart.UpBar.Border.Fill.Style);
            Assert.AreEqual(eSystemColor.CaptionText, chart.UpBar.Border.Fill.SolidFill.Color.SystemColor.Color);
        }
        [TestMethod]
        public void HighLowLines()
        {
            var ws = _pck.Workbook.Worksheets.Add("HighLowLines");
            LoadTestdata(ws);

            var chart = AddLine(ws, eLineChartType.Line, "line1", 0, 0);
            chart.Series.Add("B2:B50", "D2:D50");
            chart.Series.Add("C2:C50", "D2:D50");
            chart.AddHighLowLines();
            chart.HighLowLine.Border.Fill.Style = eFillStyle.SolidFill;
            chart.HighLowLine.Border.Fill.SolidFill.Color.SetPresetColor(ePresetColor.Red);
        }
        private static ExcelLineChart AddLine(ExcelWorksheet ws, eLineChartType type, string name, int row, int col)    
        {
            var chart = ws.Drawings.AddLineChart(name, type);
            chart.SetPosition(row, 0, col, 0);
            chart.To.Column = col+12;
            chart.To.ColumnOff = 0;
            chart.To.Row = row + 18;
            chart.To.RowOff = 0;
            chart.Series.Add("D2:D50", "A2:A50");
            return chart;
        }
    }
}

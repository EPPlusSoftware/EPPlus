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
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.Drawing.Chart.Style;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Xml;

namespace EPPlusTest.Drawing.Chart
{
    [TestClass]
    public class HistogramChartStylingTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("HistogramChartStyling.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void HistogramChart_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("HistogramChart");
            LoadTestdata(ws);
            HistogramChartStyle(ws, eChartExType.Histogram);
        }
        [TestMethod]
        public void ParetoChart_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("ParetoChart");
            LoadTestdata(ws);
            HistogramChartStyle(ws, eChartExType.Pareto);
        }
        private static void HistogramChartStyle(ExcelWorksheet ws, eChartExType type)
        {
            //Histogram Chart styles

            //Histogram chart Style 1
            AddChartEx(ws, ePresetChartStyle.HistogramChartStyle1, "HistogramChartStyle1", 0, 5, type,
                c =>
                {
                    c.Title.Text = "sunburst" +
                    " 1";
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Histogram chart Style 2
            AddChartEx(ws, ePresetChartStyle.HistogramChartStyle2, "HistogramChartStyle2", 0, 18, type,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Histogram chart Style 3
            AddChartEx(ws, ePresetChartStyle.HistogramChartStyle3, "HistogramChartStyle3", 0, 31, type,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Histogram chart Style 4
            AddChartEx(ws, ePresetChartStyle.HistogramChartStyle4, "HistogramChartStyle4", 20, 5, type,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Histogram chart Style 5
            AddChartEx(ws, ePresetChartStyle.HistogramChartStyle5, "HistogramChartStyle5", 20, 18, type,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Histogram chart Style 6
            AddChartEx(ws, ePresetChartStyle.HistogramChartStyle6, "HistogramChartStyle6", 20, 31, type,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });
        }
        private static ExcelHistogramChart AddChartEx(ExcelWorksheet ws, ePresetChartStyle style, string name, int row, int col,eChartExType type, Action<ExcelHistogramChart> SetProperties)
        {
            var chart = ws.Drawings.AddHistogramChart(name, type==eChartExType.Pareto);
            chart.SetPosition(row, 0, col, 0);
            chart.To.Column = col+12;
            chart.To.ColumnOff = 0;
            chart.To.Row = row + 18;
            chart.To.RowOff = 0;
            chart.Series.Add("A2:A8", "D2:D8");

            SetProperties(chart);

            chart.StyleManager.SetChartStyle(style);
            return chart;
        }
    }
}

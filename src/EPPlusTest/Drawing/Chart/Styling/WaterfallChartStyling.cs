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
    public class WaterfallChartStylingTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("WaterfallChartStyling.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void WaterfallChart_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("WaterfallChart");
            LoadTestdata(ws);
            WaterfallChartStyle(ws);
        }
        private static void WaterfallChartStyle(ExcelWorksheet ws)
        {
            //Waterfall Chart styles

            //Waterfall chart Style 1
            AddChartEx(ws, ePresetChartStyle.WaterfallChartStyle1, "WaterfallChartStyle1", 0, 5, 
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                    c.Legend.PositionAlignment = ePositionAlign.Min;
                    c.Series[0].DataPoints.Add(0).SubTotal = true;
                    c.Series[0].DataPoints.Add(6).SubTotal = true;
                });

            //Waterfall chart Style 2
            AddChartEx(ws, ePresetChartStyle.WaterfallChartStyle2, "WaterfallChartStyle2", 0, 18, 
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                    var dl = c.Series[0].DataLabel.DataLabels.Add(1);
                    dl.Format = "#,##0.00";
                    dl =c.Series[0].DataLabel.DataLabels.Add(0);
                    dl.Border.Width = 1;
                    dl.Border.Fill.Style = OfficeOpenXml.Drawing.eFillStyle.SolidFill;
                    dl.Border.Fill.SolidFill.Color.SetPresetColor(ePresetColor.DarkCyan);
                    dl.Position = eLabelPosition.Top;
                    dl.ShowSeriesName = true;
                    dl.ShowValue = true;
                    dl.ShowCategory = true;
                });

            //Waterfall chart Style 3
            AddChartEx(ws, ePresetChartStyle.WaterfallChartStyle3, "WaterfallChartStyle3", 0, 31, 
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Waterfall chart Style 4
            AddChartEx(ws, ePresetChartStyle.WaterfallChartStyle4, "WaterfallChartStyle4", 20, 5, 
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Waterfall chart Style 5
            AddChartEx(ws, ePresetChartStyle.WaterfallChartStyle5, "WaterfallChartStyle5", 20, 18, 
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Waterfall chart Style 6
            AddChartEx(ws, ePresetChartStyle.WaterfallChartStyle6, "WaterfallChartStyle6", 20, 31, 
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Waterfall chart Style 7
            AddChartEx(ws, ePresetChartStyle.WaterfallChartStyle7, "WaterfallChartStyle7", 40, 5,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Waterfall chart Style 8
            AddChartEx(ws, ePresetChartStyle.WaterfallChartStyle8, "WaterfallChartStyle8", 40, 18,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

        }
        private static ExcelWaterfallChart AddChartEx(ExcelWorksheet ws, ePresetChartStyle style, string name, int row, int col,Action<ExcelWaterfallChart> SetProperties)
        {
            var chart = ws.Drawings.AddWaterfallChart(name);
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

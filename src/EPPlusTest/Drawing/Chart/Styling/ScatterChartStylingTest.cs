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
using OfficeOpenXml.Drawing.Chart.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;

namespace EPPlusTest.Drawing.Chart
{
    [TestClass]
    public class ScatterChartStylingTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("ScatterChartStyling.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            _pck.Save();
            _pck.Dispose();
        }
        [TestMethod]
        public void ScatterLinesSmooth_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("ScatterSmooth");
            LoadTestdata(ws);

            ScatterLineStyle(ws, eScatterChartType.XYScatterSmooth);
        }
        [TestMethod]
        public void ScatterLines_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("ScatterLines");
            LoadTestdata(ws);

            ScatterLineStyle(ws, eScatterChartType.XYScatterLines);
        }
        [TestMethod]
        public void Scatter_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("Scatter");
            LoadTestdata(ws);

            ScatterLineStyle(ws, eScatterChartType.XYScatter);
        }
        [TestMethod]
        public void ScatterLinesNoMarkers_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("ScatterLinesNoMarkers");
            LoadTestdata(ws);

            ScatterLineStyle(ws, eScatterChartType.XYScatterLinesNoMarkers);
        }
        [TestMethod]
        public void ScatterSmoothNoMarkers_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("ScatterSmoothNoMarkers");
            LoadTestdata(ws);

            ScatterLineStyle(ws, eScatterChartType.XYScatterSmoothNoMarkers);
        }        
        private static void ScatterLineStyle(ExcelWorksheet ws, eScatterChartType chartType)
        {
            //Style 1
            AddScatter(ws, chartType, "ScatterChartStyle1", 0, 5, ePresetChartStyle.ScatterChartStyle1,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 2
            AddScatter(ws, chartType, "ScatterChartStyle2", 0, 18, ePresetChartStyle.ScatterChartStyle2,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                });

            //Style 3
            AddScatter(ws, chartType, "ScatterChartStyle3", 0, 31, ePresetChartStyle.ScatterChartStyle3,
            c =>
            {
                c.DataLabel.ShowPercent = true;
            });

            //Style 4
            AddScatter(ws, chartType, "ScatterChartStyle4", 22, 5, ePresetChartStyle.ScatterChartStyle4,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 5
            AddScatter(ws, chartType, "ScatterChartStyle5", 22, 18, ePresetChartStyle.ScatterChartStyle5,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 6
            AddScatter(ws, chartType, "ScatterChartStyle6", 22, 31, ePresetChartStyle.ScatterChartStyle6,
                c =>
                {
                });

            //Style 7
            AddScatter(ws, chartType, "ScatterChartStyle7", 44, 5, ePresetChartStyle.ScatterChartStyle7,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 8
            AddScatter(ws, chartType, "ScatterChartStyle8", 44, 18, ePresetChartStyle.ScatterChartStyle8,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                    c.DataLabel.ShowPercent = true;
                });

            //Style 9
            AddScatter(ws, chartType, "ScatterChartStyle9", 44, 31, ePresetChartStyle.ScatterChartStyle9,
                c =>
                {
                    c.Legend.Remove();
                    c.DataLabel.ShowValue = true;
                    c.DataLabel.ShowPercent = true;
                    c.DataLabel.ShowCategory = true;
                });

            //Style 10
            AddScatter(ws, chartType, "ScatterChartStyle10", 66, 5, ePresetChartStyle.ScatterChartStyle10,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                    c.DataLabel.ShowPercent = true;
                });

            //Style 11
            AddScatter(ws, chartType, "ScatterChartStyle11", 66, 18, ePresetChartStyle.ScatterChartStyle11,
                c =>
                {
                });

            //Style 12
            AddScatter(ws, chartType, "ScatterChartStyle12", 66, 31, ePresetChartStyle.ScatterChartStyle12,
                c =>
                {
                });
        }
        
        private static ExcelScatterChart AddScatter(ExcelWorksheet ws, eScatterChartType type, string name, int row, int col, ePresetChartStyle style, Action<ExcelScatterChart> SetProperties)    
        {
            var chart = ws.Drawings.AddScatterChart(name, type);
            chart.SetPosition(row, 0, col, 0);
            chart.To.Column = col+12;
            chart.To.ColumnOff = 0;
            chart.To.Row = row + 18;
            chart.To.RowOff = 0;
            var serie = chart.Series.Add("D2:D8", "A2:A8");

            SetProperties(chart);

            chart.StyleManager.SetChartStyle(style);
            return chart;
        }
    }
}

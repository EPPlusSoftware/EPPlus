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
    public class RadarChartStylingTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("RadarChartStyling.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void Radar_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("Radar");
            LoadTestdata(ws);

            RadarLineStyle(ws, eRadarChartType.Radar);
        }
        [TestMethod]
        public void RadarFilled_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("RadarFilled");
            LoadTestdata(ws);

            RadarLineStyle(ws, eRadarChartType.RadarFilled);
        }
        [TestMethod]
        public void RadarMarkers_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("RadarMarkers");
            LoadTestdata(ws);

            RadarLineStyle(ws, eRadarChartType.RadarMarkers);
        }
        private static void RadarLineStyle(ExcelWorksheet ws, eRadarChartType chartType)
        {
            //Style 1
            AddRadar(ws, chartType, "RadarChartStyle1", 0, 5, ePresetChartStyle.RadarChartStyle1,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 2
            AddRadar(ws, chartType, "RadarChartStyle2", 0, 18, ePresetChartStyle.RadarChartStyle2,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                });

            //Style 3
            AddRadar(ws, chartType, "RadarChartStyle3", 0, 31, ePresetChartStyle.RadarChartStyle3,
            c =>
            {
                c.DataLabel.ShowPercent = true;
            });

            //Style 4
            AddRadar(ws, chartType, "RadarChartStyle4", 22, 5, ePresetChartStyle.RadarChartStyle4,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 5
            AddRadar(ws, chartType, "RadarChartStyle5", 22, 18, ePresetChartStyle.RadarChartStyle5,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 6
            AddRadar(ws, chartType, "RadarChartStyle6", 22, 31, ePresetChartStyle.RadarChartStyle6,
                c =>
                {
                });

            //Style 7
            AddRadar(ws, chartType, "RadarChartStyle7", 44, 5, ePresetChartStyle.RadarChartStyle7,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 8
            AddRadar(ws, chartType, "RadarChartStyle8", 44, 18, ePresetChartStyle.RadarChartStyle8,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                    c.DataLabel.ShowPercent = true;
                });
        }
        private static ExcelRadarChart AddRadar(ExcelWorksheet ws, eRadarChartType type, string name, int row, int col, ePresetChartStyle style, Action<ExcelRadarChart> SetProperties)    
        {
            var chart = ws.Drawings.AddRadarChart(name, type);
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

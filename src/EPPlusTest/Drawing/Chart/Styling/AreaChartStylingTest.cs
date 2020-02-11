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
    public class AreaChartStylingTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("AreaChartStyling.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void Area_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("AreaChartStyling");
            LoadTestdata(ws);

            AreaStyle(ws, eAreaChartType.Area);
        }
        [TestMethod]
        public void Area3D_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("Area3DChartStyles");
            LoadTestdata(ws);

            Area3DStyle(ws, eAreaChartType.Area3D);
        }

        private static void AreaStyle(ExcelWorksheet ws, eAreaChartType chartType)
        {
            //Style 1
            AddArea(ws, chartType, "AreaChartStyle1", 0, 5, ePresetChartStyle.AreaChartStyle1,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 2
            AddArea(ws, chartType, "AreaChartStyle2", 0, 18, ePresetChartStyle.AreaChartStyle2,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                });

            //Style 3
            AddArea(ws, chartType, "AreaChartStyle3", 0, 31, ePresetChartStyle.AreaChartStyle3,
            c =>
            { 
              
            });

            //Style 4
            AddArea(ws, chartType, "AreaChartStyle4", 22, 5, ePresetChartStyle.AreaChartStyle4,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 5
            AddArea(ws, chartType, "AreaChartStyle5", 22, 18, ePresetChartStyle.AreaChartStyle5,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 6
            AddArea(ws, chartType, "AreaChartStyle6", 22, 31, ePresetChartStyle.AreaChartStyle6,
                c =>
                {
                });

            //Style 7
            AddArea(ws, chartType, "AreaChartStyle7", 44, 5, ePresetChartStyle.AreaChartStyle7,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 8
            AddArea(ws, chartType, "AreaChartStyle8", 44, 18, ePresetChartStyle.AreaChartStyle8,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                });

            //Style 9
            AddArea(ws, chartType, "AreaChartStyle9", 44, 31, ePresetChartStyle.AreaChartStyle9,
                c =>
                {
                    c.Legend.Remove();
                });

            //Style 10
            AddArea(ws, chartType, "AreaChartStyle10", 66, 5, ePresetChartStyle.AreaChartStyle10,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 11
            AddArea(ws, chartType, "AreaChartStyle11", 66, 18, ePresetChartStyle.AreaChartStyle11,
                c =>
                {
                });
        }
        private static void Area3DStyle(ExcelWorksheet ws, eAreaChartType chartType)
        {
            //Style 1
            AddArea(ws, chartType, "AreaChartStyle1", 0, 5, ePresetChartStyle.Area3dChartStyle1,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 2
            AddArea(ws, chartType, "AreaChartStyle2", 0, 18, ePresetChartStyle.Area3dChartStyle2,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                });

            //Style 3
            AddArea(ws, chartType, "AreaChartStyle3", 0, 31, ePresetChartStyle.Area3dChartStyle3,
            c =>
            {
            });

            //Style 4
            AddArea(ws, chartType, "AreaChartStyle4", 22, 5, ePresetChartStyle.Area3dChartStyle4,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 5
            AddArea(ws, chartType, "AreaChartStyle5", 22, 18, ePresetChartStyle.Area3dChartStyle5,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 6
            AddArea(ws, chartType, "AreaChartStyle6", 22, 31, ePresetChartStyle.Area3dChartStyle6,
                c =>
                {
                });

            //Style 7
            AddArea(ws, chartType, "AreaChartStyle7", 44, 5, ePresetChartStyle.Area3dChartStyle7,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 8
            AddArea(ws, chartType, "AreaChartStyle8", 44, 18, ePresetChartStyle.Area3dChartStyle8,
                c =>
                {
                });

            //Style 9
            AddArea(ws, chartType, "AreaChartStyle9", 44, 31, ePresetChartStyle.Area3dChartStyle9,
                c =>
                {
                });

            //Style 10
            AddArea(ws, chartType, "AreaChartStyle10", 66, 5, ePresetChartStyle.Area3dChartStyle10,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });
        }
        private static ExcelAreaChart AddArea(ExcelWorksheet ws, eAreaChartType type, string name, int row, int col, ePresetChartStyle style, Action<ExcelAreaChart> SetProperties)    
        {
            var chart = ws.Drawings.AddAreaChart(name, type);
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


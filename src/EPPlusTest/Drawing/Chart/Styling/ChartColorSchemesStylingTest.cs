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
    public class ChartColorSchemesStylingTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("ColorSchemesChartStyling.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void ColorSchemesArea3D()
        {
            var ws = _pck.Workbook.Worksheets.Add("Area3DChartStyles");
            LoadTestdata(ws);

            Area3DStyle(ws, eAreaChartType.Area3D);
        }
        [TestMethod]
        public void ColorSchemesPie3D()
        {
            var ws = _pck.Workbook.Worksheets.Add("Pie3DChartStyles");
            LoadTestdata(ws);

            Pie3DStyle(ws, ePieChartType.Pie3D);
        }

        private static void Pie3DStyle(ExcelWorksheet ws, ePieChartType chartType)
        {
            //Style 1
            AddPieWithColor(ws, chartType, "ColorStyleColorfulPalette1", 0, 5, ePresetChartStyle.Pie3dChartStyle1, ePresetChartColors.ColorfulPalette1,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 2
            AddPieWithColor(ws, chartType, "ColorStyleColorfulPalette2", 0, 18, ePresetChartStyle.Pie3dChartStyle2, ePresetChartColors.ColorfulPalette2,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                });

            //Style 3
            AddPieWithColor(ws, chartType, "ColorStyleColorfulPalette3", 0, 31, ePresetChartStyle.Pie3dChartStyle3, ePresetChartColors.ColorfulPalette3,
            c =>
            {
            });

            //Style 4
            AddPieWithColor(ws, chartType, "ColorStyleColorfulPalette4", 22, 5, ePresetChartStyle.Pie3dChartStyle4, ePresetChartColors.ColorfulPalette4,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 5
            AddPieWithColor(ws, chartType, "ColorStyleMonochromaticPalette1", 22, 18, ePresetChartStyle.Pie3dChartStyle5, ePresetChartColors.MonochromaticPalette1,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 6
            AddPieWithColor(ws, chartType, "ColorStyleMonochromaticPalette2", 22, 31, ePresetChartStyle.Pie3dChartStyle6, ePresetChartColors.MonochromaticPalette2,
                c =>
                {
                });

            //Style 7
            AddPieWithColor(ws, chartType, "ColorStyleMonochromaticPalette3", 44, 5, ePresetChartStyle.Pie3dChartStyle7, ePresetChartColors.MonochromaticPalette3,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 8
            AddPieWithColor(ws, chartType, "ColorStyleMonochromaticPalette4", 44, 18, ePresetChartStyle.Pie3dChartStyle8, ePresetChartColors.MonochromaticPalette4,
                c =>
                {
                });

            //Style 9
            AddPieWithColor(ws, chartType, "ColorStyleMonochromaticPalette5", 44, 31, ePresetChartStyle.Pie3dChartStyle9, ePresetChartColors.MonochromaticPalette5,
                c =>
                {
                });

            //Style 10
            AddPieWithColor(ws, chartType, "ColorStyleMonochromaticPalette6", 66, 5, ePresetChartStyle.Pie3dChartStyle10, ePresetChartColors.MonochromaticPalette6,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 10
            AddPieWithColor(ws, chartType, "ColorStyleMonochromaticPalette7", 66, 18, ePresetChartStyle.Pie3dChartStyle9, ePresetChartColors.MonochromaticPalette7,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 10
            AddPieWithColor(ws, chartType, "ColorStyleMonochromaticPalette8", 66, 31, ePresetChartStyle.Pie3dChartStyle8, ePresetChartColors.MonochromaticPalette8,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 10
            AddPieWithColor(ws, chartType, "ColorStyleMonochromaticPalette9", 88, 5, ePresetChartStyle.Pie3dChartStyle7, ePresetChartColors.MonochromaticPalette9,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 10
            AddPieWithColor(ws, chartType, "ColorStyleMonochromaticPalette10", 88, 18, ePresetChartStyle.Pie3dChartStyle6, ePresetChartColors.MonochromaticPalette10,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 10
            AddPieWithColor(ws, chartType, "ColorStyleMonochromaticPalette11", 88, 31, ePresetChartStyle.Pie3dChartStyle5, ePresetChartColors.MonochromaticPalette11,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 10
            AddPieWithColor(ws, chartType, "ColorStyleMonochromaticPalette12", 110, 5, ePresetChartStyle.Pie3dChartStyle4, ePresetChartColors.MonochromaticPalette12,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 10
            AddPieWithColor(ws, chartType, "ColorStyleMonochromaticPalette13", 110, 18, ePresetChartStyle.Pie3dChartStyle3, ePresetChartColors.MonochromaticPalette13,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });
        }
        private static void Area3DStyle(ExcelWorksheet ws, eAreaChartType chartType)
        {
            //Style 1
            AddAreaWithColor(ws, chartType, "ColorStyleColorfulPalette1", 0, 5, ePresetChartStyle.Area3dChartStyle1, ePresetChartColors.ColorfulPalette1,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 2
            AddAreaWithColor(ws, chartType, "ColorStyleColorfulPalette2", 0, 18, ePresetChartStyle.Area3dChartStyle2, ePresetChartColors.ColorfulPalette2,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                });

            //Style 3
            AddAreaWithColor(ws, chartType, "ColorStyleColorfulPalette3", 0, 31, ePresetChartStyle.Area3dChartStyle3, ePresetChartColors.ColorfulPalette3,
            c =>
            {
            });

            //Style 4
            AddAreaWithColor(ws, chartType, "ColorStyleColorfulPalette4", 22, 5, ePresetChartStyle.Area3dChartStyle4, ePresetChartColors.ColorfulPalette4,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 5
            AddAreaWithColor(ws, chartType, "ColorStyleMonochromaticPalette1", 22, 18, ePresetChartStyle.Area3dChartStyle5, ePresetChartColors.MonochromaticPalette1,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 6
            AddAreaWithColor(ws, chartType, "ColorStyleMonochromaticPalette2", 22, 31, ePresetChartStyle.Area3dChartStyle6, ePresetChartColors.MonochromaticPalette2,
                c =>
                {
                });

            //Style 7
            AddAreaWithColor(ws, chartType, "ColorStyleMonochromaticPalette3", 44, 5, ePresetChartStyle.Area3dChartStyle7, ePresetChartColors.MonochromaticPalette3,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 8
            AddAreaWithColor(ws, chartType, "ColorStyleMonochromaticPalette4", 44, 18, ePresetChartStyle.Area3dChartStyle8, ePresetChartColors.MonochromaticPalette4,
                c =>
                {
                });

            //Style 9
            AddAreaWithColor(ws, chartType, "ColorStyleMonochromaticPalette5", 44, 31, ePresetChartStyle.Area3dChartStyle9, ePresetChartColors.MonochromaticPalette5,
                c =>
                {
                });

            //Style 10
            AddAreaWithColor(ws, chartType, "ColorStyleMonochromaticPalette6", 66, 5, ePresetChartStyle.Area3dChartStyle10, ePresetChartColors.MonochromaticPalette6,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 10
            AddAreaWithColor(ws, chartType, "ColorStyleMonochromaticPalette7", 66, 18, ePresetChartStyle.Area3dChartStyle10, ePresetChartColors.MonochromaticPalette7,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 10
            AddAreaWithColor(ws, chartType, "ColorStyleMonochromaticPalette8", 66, 31, ePresetChartStyle.Area3dChartStyle10, ePresetChartColors.MonochromaticPalette8,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 10
            AddAreaWithColor(ws, chartType, "ColorStyleMonochromaticPalette9", 88, 5, ePresetChartStyle.Area3dChartStyle10, ePresetChartColors.MonochromaticPalette9,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 10
            AddAreaWithColor(ws, chartType, "ColorStyleMonochromaticPalette10", 88, 18, ePresetChartStyle.Area3dChartStyle10, ePresetChartColors.MonochromaticPalette10,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 10
            AddAreaWithColor(ws, chartType, "ColorStyleMonochromaticPalette11", 88, 31, ePresetChartStyle.Area3dChartStyle10, ePresetChartColors.MonochromaticPalette11,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 10
            AddAreaWithColor(ws, chartType, "ColorStyleMonochromaticPalette12", 110, 5, ePresetChartStyle.Area3dChartStyle10, ePresetChartColors.MonochromaticPalette12,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 10
            AddAreaWithColor(ws, chartType, "ColorStyleMonochromaticPalette13", 110, 18, ePresetChartStyle.Area3dChartStyle10, ePresetChartColors.MonochromaticPalette13,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });
        }
        private static ExcelAreaChart AddAreaWithColor(ExcelWorksheet ws, eAreaChartType type, string name, int row, int col, ePresetChartStyle style, ePresetChartColors colors, Action<ExcelAreaChart> SetProperties)    
        {
            var chart = ws.Drawings.AddAreaChart(name, type);
            chart.SetPosition(row, 0, col, 0);
            chart.To.Column = col+12;
            chart.To.ColumnOff = 0;
            chart.To.Row = row + 18;
            chart.To.RowOff = 0;
            var serie = chart.Series.Add("D2:D8", "A2:A8");

            SetProperties(chart);

            chart.StyleManager.SetChartStyle(style, colors);
            return chart;
        }
        private static ExcelPieChart AddPieWithColor(ExcelWorksheet ws, ePieChartType type, string name, int row, int col, ePresetChartStyle style, ePresetChartColors colors, Action<ExcelPieChart> SetProperties)
        {
            var chart = ws.Drawings.AddPieChart(name, type);
            chart.SetPosition(row, 0, col, 0);
            chart.To.Column = col + 12;
            chart.To.ColumnOff = 0;
            chart.To.Row = row + 18;
            chart.To.RowOff = 0;
            var serie = chart.Series.Add("D2:D8", "A2:A8");

            SetProperties(chart);

            chart.StyleManager.SetChartStyle(style, colors);
            return chart;
        }
    }
}


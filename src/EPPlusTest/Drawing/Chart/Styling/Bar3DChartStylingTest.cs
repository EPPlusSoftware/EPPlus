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
    public class Bar3DChartStylingTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("BarChart3DStyling.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void BarChart3D_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("BarClustChart3DStyles");
            LoadTestdata(ws);

            StyleBar3DChart(ws, eBarChartType.BarClustered3D);
        }
        [TestMethod]
        public void BarStackedChart3D_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("BarStackedChart3D");
            LoadTestdata(ws);

            StyleStackedBar3DChart(ws, eBarChartType.BarStacked3D);
        }
        [TestMethod]
        public void BarStacked100Chart3D_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("BarStacked100Chart3D");
            LoadTestdata(ws);

            StyleStackedBar3DChart(ws, eBarChartType.BarStacked1003D);
        }
        [TestMethod]
        public void PyramidBarChart3D_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("PyramidBarChart3D");
            LoadTestdata(ws);

            StyleBar3DChart(ws, eBarChartType.PyramidBarClustered);
        }
        [TestMethod]
        public void PyramidBarStackedChart3D_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("PyramidBarStackedChart3D");
            LoadTestdata(ws);

            StyleStackedBar3DChart(ws, eBarChartType.PyramidBarStacked100);
        }
        [TestMethod]
        public void PyramidBarStacked100Chart3D_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("PyramidBarStacked100Chart3D");
            LoadTestdata(ws);

            StyleStackedBar3DChart(ws, eBarChartType.PyramidBarStacked100);
        }
        private static void StyleColumn3DChart(ExcelWorksheet ws, eBarChartType chartType)
        {
            //Style 1
            AddBar(ws, chartType, "Column3DChartStyle1", 0, 5, ePresetChartStyle.Column3dChartStyle1,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                    c.Axis[0].MajorTickMark = eAxisTickMark.None;
                    c.Axis[0].MinorTickMark = eAxisTickMark.None;
                    c.Axis[1].MajorTickMark = eAxisTickMark.None;
                    c.Axis[1].MinorTickMark = eAxisTickMark.None;
                });

            //Style 2
            var chart2 = AddBar(ws, chartType, "Column3DChartStyle2", 0, 18, ePresetChartStyle.Column3dChartStyle2,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                });

            //Style 3
            var chart3 = AddBar(ws, chartType, "Column3DChartStyle3", 0, 31, ePresetChartStyle.Column3dChartStyle3,
            c =>
            {
                c.DataLabel.ShowPercent = true;
            });

            //Style 4
            AddBar(ws, chartType, "Column3DChartStyle4", 22, 5, ePresetChartStyle.Column3dChartStyle4,
                c =>
                {
                });

            //Style 5
            AddBar(ws, chartType, "Column3DChartStyle5", 22, 18, ePresetChartStyle.Column3dChartStyle5,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                });

            //Style 6
            AddBar(ws, chartType, "Column3DChartStyle6", 22, 31, ePresetChartStyle.Column3dChartStyle6,
            c =>
            {
                c.DataLabel.ShowPercent = true;
            });

            //Style 7
            AddBar(ws, chartType, "Column3DChartStyle7", 44, 5, ePresetChartStyle.Column3dChartStyle7,
                c =>
                {
                });

            //Style 8
            AddBar(ws, chartType, "Column3DChartStyle8", 44, 18, ePresetChartStyle.Column3dChartStyle8,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                });
        }
        private static void StyleColumn3DChart_MultiSeries(ExcelWorksheet ws, eBarChartType chartType)
        {
            //Style 1
            AddBarMultiSeries(ws, chartType, "Column3DChartStyle1", 0, 5, ePresetMultiSeriesChartStyle.Column3dChartStyle1,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                    c.Axis[0].MajorTickMark = eAxisTickMark.None;
                    c.Axis[0].MinorTickMark = eAxisTickMark.None;
                    c.Axis[1].MajorTickMark = eAxisTickMark.None;
                    c.Axis[1].MinorTickMark = eAxisTickMark.None;
                });

            //Style 2
            var chart2 = AddBarMultiSeries(ws, chartType, "Column3DChartStyle2", 0, 18, ePresetMultiSeriesChartStyle.Column3dChartStyle2,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                });

            //Style 3
            var chart3 = AddBarMultiSeries(ws, chartType, "Column3DChartStyle3", 0, 31, ePresetMultiSeriesChartStyle.Column3dChartStyle3,
            c =>
            {
                c.DataLabel.ShowPercent = true;
            });

            //Style 4
            AddBarMultiSeries(ws, chartType, "Column3DChartStyle4", 22, 5, ePresetMultiSeriesChartStyle.Column3dChartStyle4,
                c =>
                {
                });

            //Style 5
            AddBarMultiSeries(ws, chartType, "Column3DChartStyle5", 22, 18, ePresetMultiSeriesChartStyle.Column3dChartStyle5,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                });

            //Style 6
            AddBarMultiSeries(ws, chartType, "Column3DChartStyle6", 22, 31, ePresetMultiSeriesChartStyle.Column3dChartStyle6,
            c =>
            {
                c.DataLabel.ShowPercent = true;
            });

            //Style 7
            AddBarMultiSeries(ws, chartType, "Column3DChartStyle7", 44, 5, ePresetMultiSeriesChartStyle.Column3dChartStyle7,
                c =>
                {
                });

            //Style 8
            AddBarMultiSeries(ws, chartType, "Column3DChartStyle8", 44, 18, ePresetMultiSeriesChartStyle.Column3dChartStyle8,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                });

            //Style 9
            AddBarMultiSeries(ws, chartType, "Column3DChartStyle9", 44, 31, ePresetMultiSeriesChartStyle.Column3dChartStyle9,
            c =>
            {
                c.DataLabel.ShowPercent = true;
            });

            //Style 10
            AddBarMultiSeries(ws, chartType, "Column3DChartStyle10", 66, 5, ePresetMultiSeriesChartStyle.Column3dChartStyle10,
                c =>
                {
                });

            //Style 11
            AddBarMultiSeries(ws, chartType, "Column3DChartStyle11", 66, 18, ePresetMultiSeriesChartStyle.Column3dChartStyle11,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                });
        }
        private static void StyleBar3DChart(ExcelWorksheet ws, eBarChartType chartType)
        {
            //Style 1
            AddBar(ws, chartType, "Column3DChartStyle1", 0, 5, ePresetChartStyle.Bar3dChartStyle1,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                    c.Axis[0].MajorTickMark = eAxisTickMark.None;
                    c.Axis[0].MinorTickMark = eAxisTickMark.None;
                    c.Axis[1].MajorTickMark = eAxisTickMark.None;
                    c.Axis[1].MinorTickMark = eAxisTickMark.None;
                });

            //Style 2
            var chart2 = AddBar(ws, chartType, "Column3DChartStyle2", 0, 18, ePresetChartStyle.Bar3dChartStyle2,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                });

            //Style 3
            var chart3 = AddBar(ws, chartType, "Column3DChartStyle3", 0, 31, ePresetChartStyle.Bar3dChartStyle3,
            c =>
            {
                c.DataLabel.ShowPercent = true;
            });

            //Style 4
            AddBar(ws, chartType, "Column3DChartStyle4", 22, 5, ePresetChartStyle.Bar3dChartStyle4,
                c =>
                {
                });

            //Style 5
            AddBar(ws, chartType, "Column3DChartStyle5", 22, 18, ePresetChartStyle.Bar3dChartStyle5,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                });

            //Style 6
            AddBar(ws, chartType, "Column3DChartStyle6", 22, 31, ePresetChartStyle.Bar3dChartStyle6,
            c =>
            {
                c.DataLabel.ShowPercent = true;
            });

            //Style 7
            AddBar(ws, chartType, "Column3DChartStyle7", 44, 5, ePresetChartStyle.Bar3dChartStyle7,
                c =>
                {
                });

            //Style 8
            AddBar(ws, chartType, "Column3DChartStyle8", 44, 18, ePresetChartStyle.Bar3dChartStyle8,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                });

            //Style 9
            AddBar(ws, chartType, "Column3DChartStyle9", 44, 31, ePresetChartStyle.Bar3dChartStyle9,
            c =>
            {
                c.DataLabel.ShowPercent = true;
            });

            //Style 10
            AddBar(ws, chartType, "Column3DChartStyle10", 66, 5, ePresetChartStyle.Bar3dChartStyle10,
                c =>
                {
                });

            //Style 11
            AddBar(ws, chartType, "Column3DChartStyle11", 66, 18, ePresetChartStyle.Bar3dChartStyle11,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                });
        }
        private static void StyleStackedBar3DChart(ExcelWorksheet ws, eBarChartType chartType)
        {
            //Style 1
            AddBar(ws, chartType, "Column3DChartStyle1", 0, 5, ePresetChartStyle.StackedColumn3dChartStyle1,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                    c.Axis[0].MajorTickMark = eAxisTickMark.None;
                    c.Axis[0].MinorTickMark = eAxisTickMark.None;
                    c.Axis[1].MajorTickMark = eAxisTickMark.None;
                    c.Axis[1].MinorTickMark = eAxisTickMark.None;
                });

            //Style 2
            var chart2 = AddBar(ws, chartType, "Column3DChartStyle2", 0, 18, ePresetChartStyle.StackedColumn3dChartStyle2,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                });

            //Style 3
            var chart3 = AddBar(ws, chartType, "Column3DChartStyle3", 0, 31, ePresetChartStyle.StackedColumn3dChartStyle3,
            c =>
            {
                c.DataLabel.ShowPercent = true;
            });

            //Style 4
            AddBar(ws, chartType, "Column3DChartStyle4", 22, 5, ePresetChartStyle.StackedColumn3dChartStyle4,
                c =>
                {
                });

            //Style 5
            AddBar(ws, chartType, "Column3DChartStyle5", 22, 18, ePresetChartStyle.StackedColumn3dChartStyle5,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                });

            //Style 6
            AddBar(ws, chartType, "Column3DChartStyle6", 22, 31, ePresetChartStyle.StackedColumn3dChartStyle6,
            c =>
            {
                c.DataLabel.ShowPercent = true;
            });

            //Style 7
            AddBar(ws, chartType, "Column3DChartStyle7", 44, 5, ePresetChartStyle.StackedColumn3dChartStyle7,
                c =>
                {
                });

            //Style 8
            AddBar(ws, chartType, "Column3DChartStyle8", 44, 18, ePresetChartStyle.StackedColumn3dChartStyle8,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                });
        }
        private static ExcelBarChart AddBar(ExcelWorksheet ws, eBarChartType type, string name, int row, int col, ePresetChartStyle style, Action<ExcelBarChart> SetProperties)
        {
            var chart = ws.Drawings.AddBarChart(name, type);
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
        private static ExcelBarChart AddBarMultiSeries(ExcelWorksheet ws, eBarChartType type, string name, int row, int col, ePresetMultiSeriesChartStyle style, Action<ExcelBarChart> SetProperties)
        {
            var chart = ws.Drawings.AddBarChart(name, type);
            chart.SetPosition(row, 0, col, 0);
            chart.To.Column = col + 12;
            chart.To.ColumnOff = 0;
            chart.To.Row = row + 18;
            chart.To.RowOff = 0;
            var serie = chart.Series.Add("D2:D8", "A2:A8");
            var serie2 = chart.Series.Add("B2:B8", "A2:A8");
            SetProperties(chart);

            chart.StyleManager.SetChartStyle(style);
            return chart;
        }
    }
}

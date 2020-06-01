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
using System.IO;
using System.Xml;

namespace EPPlusTest.Drawing.Chart.Styling
{
    [TestClass]
    public class BarChartStylingTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("BarChartStyling.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void BarChart_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("BarClusteredChartStyles");
            LoadTestdata(ws);

            StyleBarChart(ws, eBarChartType.BarClustered);
        }
        [TestMethod]
        public void BarChartStacked_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("BarStackedChartStyles");
            LoadTestdata(ws);

            StyleStackedBarChart(ws, eBarChartType.BarStacked);
        }
        [TestMethod]
        public void BarChartStacked100_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("BarStacked100ChartStyles");
            LoadTestdata(ws);

            StyleStackedBarChart(ws, eBarChartType.BarStacked100);
        }

        private static void StyleBarChart(ExcelWorksheet ws, eBarChartType chartType)
        {
            //Style 1
            AddBar(ws, chartType, "ColumnChartStyle1", 0, 5, ePresetChartStyle.BarChartStyle1,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                    c.Axis[0].MajorTickMark = eAxisTickMark.None;
                    c.Axis[0].MinorTickMark = eAxisTickMark.None;
                    c.Axis[1].MajorTickMark = eAxisTickMark.None;
                    c.Axis[1].MinorTickMark = eAxisTickMark.None;
                    c.GapWidth = 219;
                    c.Overlap = -27;
                });

            //Style 2
            AddBar(ws, chartType, "ColumnChartStyle2", 0, 18, ePresetChartStyle.BarChartStyle2,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                });

            //Style 3
            var chart3 = AddBar(ws, chartType, "ColumnChartStyle3", 0, 31, ePresetChartStyle.BarChartStyle3,
            c =>
            {
                c.DataLabel.ShowPercent = true;
                c.DataLabel.Position = eLabelPosition.Center;
            });

            //Style 4
            AddBar(ws, chartType, "ColumnChartStyle4", 22, 5, ePresetChartStyle.BarChartStyle4,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 5
            AddBar(ws, chartType, "ColumnChartStyle5", 22, 18, ePresetChartStyle.BarChartStyle5,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 6
            AddBar(ws, chartType, "ColumnChartStyle6", 22, 31, ePresetChartStyle.BarChartStyle6,
                c =>
                {
                });


            //Style 7
            AddBar(ws, chartType, "ColumnChartStyle7", 44, 5, ePresetChartStyle.BarChartStyle7,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 8
            AddBar(ws, chartType, "ColumnChartStyle8", 44, 18, ePresetChartStyle.BarChartStyle8,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 9
            AddBar(ws, chartType, "ColumnChartStyle9", 44, 31, ePresetChartStyle.BarChartStyle9,
                c =>
                {
                });

            //Style 10
            AddBar(ws, chartType, "ColumnChartStyle10", 66, 5, ePresetChartStyle.BarChartStyle10,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 11
            AddBar(ws, chartType, "ColumnChartStyle11", 66, 18, ePresetChartStyle.BarChartStyle11,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 12
            AddBar(ws, chartType, "ColumnChartStyle12", 66, 31, ePresetChartStyle.BarChartStyle12,
                c =>
                {
                });

            //Style 13
            AddBar(ws, chartType, "ColumnChartStyle13", 88, 5, ePresetChartStyle.BarChartStyle13,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });
        }
        private static void StyleStackedBarChart(ExcelWorksheet ws, eBarChartType chartType)
        {
            //Style 1
            AddBar(ws, chartType, "ColumnChartStyle1", 0, 5, ePresetChartStyle.StackedBarChartStyle1,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                    c.Axis[0].MajorTickMark = eAxisTickMark.None;
                    c.Axis[0].MinorTickMark = eAxisTickMark.None;
                    c.Axis[1].MajorTickMark = eAxisTickMark.None;
                    c.Axis[1].MinorTickMark = eAxisTickMark.None;
                    c.GapWidth = 219;
                    c.Overlap = -27;
                });

            //Style 2
            AddBar(ws, chartType, "ColumnChartStyle2", 0, 18, ePresetChartStyle.StackedBarChartStyle2,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                });

            //Style 3
            var chart3 = AddBar(ws, chartType, "ColumnChartStyle3", 0, 31, ePresetChartStyle.StackedBarChartStyle3,
            c =>
            {
                c.DataLabel.ShowPercent = true;
                c.DataLabel.Position = eLabelPosition.Center;
            });

            //Style 4
            AddBar(ws, chartType, "ColumnChartStyle4", 22, 5, ePresetChartStyle.StackedBarChartStyle4,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 5
            AddBar(ws, chartType, "ColumnChartStyle5", 22, 18, ePresetChartStyle.StackedBarChartStyle5,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 6
            AddBar(ws, chartType, "ColumnChartStyle6", 22, 31, ePresetChartStyle.StackedBarChartStyle6,
                c =>
                {
                });


            //Style 7
            AddBar(ws, chartType, "ColumnChartStyle7", 44, 5, ePresetChartStyle.StackedBarChartStyle7,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 8
            AddBar(ws, chartType, "ColumnChartStyle8", 44, 18, ePresetChartStyle.StackedBarChartStyle8,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 9
            AddBar(ws, chartType, "ColumnChartStyle9", 44, 31, ePresetChartStyle.StackedBarChartStyle9,
                c =>
                {
                });

            //Style 10
            AddBar(ws, chartType, "ColumnChartStyle10", 66, 5, ePresetChartStyle.StackedBarChartStyle10,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 11
            AddBar(ws, chartType, "ColumnChartStyle11", 66, 18, ePresetChartStyle.StackedBarChartStyle11,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
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
    }
}

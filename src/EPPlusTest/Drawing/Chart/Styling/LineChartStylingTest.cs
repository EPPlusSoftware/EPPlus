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
    public class LineChartStylingTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("LineChartStyling.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void LineChart_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("LineChartStyles");
            LoadTestdata(ws);

            StyleLineChart(ws, eLineChartType.Line);
        }
        [TestMethod]
        public void LineChartStacked_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("LineStackedChartStyles");
            LoadTestdata(ws);

            StyleLineChart(ws, eLineChartType.LineStacked);
        }

        [TestMethod]
        public void LineChart3D_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("Line3DChartStyles");
            LoadTestdata(ws);

            StyleLine3dChart(ws, eLineChartType.Line3D);
        }
        [TestMethod]
        public void LineMarkersChartStacked_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("LineMarkersChartStyles");
            LoadTestdata(ws);

            StyleLineChart(ws, eLineChartType.LineMarkers);
        }

        private static void StyleLineChart(ExcelWorksheet ws, eLineChartType chartType)
        {
            //Style 1
            AddLine(ws, chartType, "ColumnChartStyle1", 0, 5, ePresetChartStyle.LineChartStyle1,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                    c.Axis[0].MajorTickMark = eAxisTickMark.None;
                    c.Axis[0].MinorTickMark = eAxisTickMark.None;
                    c.Axis[1].MajorTickMark = eAxisTickMark.None;
                    c.Axis[1].MinorTickMark = eAxisTickMark.None;
                });

            //Style 2
            AddLine(ws, chartType, "ColumnChartStyle2", 0, 18, ePresetChartStyle.LineChartStyle2,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                    c.YAxis.Deleted = true;                    
                    c.Axis[0].MajorTickMark = eAxisTickMark.None;
                    c.Axis[0].MinorTickMark = eAxisTickMark.None;
                    c.Axis[1].MajorTickMark = eAxisTickMark.None;
                    c.Axis[1].MinorTickMark = eAxisTickMark.None;
                    if (chartType != eLineChartType.Line3D)
                    {
                        c.Marker = true;
                    }
                    c.DataLabel.ShowValue = true;
                    foreach(var serie in c.Series)
                    {
                        if (chartType != eLineChartType.Line3D)
                        {
                            serie.Marker.Style = eMarkerStyle.Circle;
                            serie.Marker.Size = 17;
                        }
                        serie.DataLabel.Position=eLabelPosition.Center;
                        serie.DataLabel.ShowValue = true;
                    }
                });

            //Style 3
            AddLine(ws, chartType, "ColumnChartStyle3", 0, 31, ePresetChartStyle.LineChartStyle3,
            c =>
            {
                c.DataLabel.ShowPercent = true;
                c.DataLabel.Position = eLabelPosition.Center;
                c.AddDropLines();
            });

            //Style 4
            AddLine(ws, chartType, "ColumnChartStyle4", 22, 5, ePresetChartStyle.LineChartStyle4,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 5
            AddLine(ws, chartType, "ColumnChartStyle5", 22, 18, ePresetChartStyle.LineChartStyle5,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 6
            AddLine(ws, chartType, "ColumnChartStyle6", 22, 31, ePresetChartStyle.LineChartStyle6,
                c =>
                {
                });


            //Style 7
            AddLine(ws, chartType, "ColumnChartStyle7", 44, 5, ePresetChartStyle.LineChartStyle7,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 8
            AddLine(ws, chartType, "ColumnChartStyle8", 44, 18, ePresetChartStyle.LineChartStyle8,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 9
            AddLine(ws, chartType, "ColumnChartStyle9", 44, 31, ePresetChartStyle.LineChartStyle9,
                c =>
                {
                });

            //Style 10
            AddLine(ws, chartType, "ColumnChartStyle10", 66, 5, ePresetChartStyle.LineChartStyle10,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 11
            AddLine(ws, chartType, "ColumnChartStyle11", 66, 18, ePresetChartStyle.LineChartStyle11,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 12
            AddLine(ws, chartType, "ColumnChartStyle12", 66, 31, ePresetChartStyle.LineChartStyle12,
                c =>
                {
                });

            //Style 13
            AddLine(ws, chartType, "ColumnChartStyle13", 88, 5, ePresetChartStyle.LineChartStyle13,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 14
            AddLine(ws, chartType, "ColumnChartStyle14", 88, 18, ePresetChartStyle.LineChartStyle14,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 15
            AddLine(ws, chartType, "ColumnChartStyle15", 88, 31, ePresetChartStyle.LineChartStyle15,
                c =>
                {
                });
        }
        private static void StyleLine3dChart(ExcelWorksheet ws, eLineChartType chartType)
        {
            //Style 1
            AddLine3D(ws, chartType, "Line3dChartStyle1", 0, 5, ePresetChartStyle.Line3dChartStyle1,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                    c.Axis[0].MajorTickMark = eAxisTickMark.None;
                    c.Axis[0].MinorTickMark = eAxisTickMark.None;
                    c.Axis[1].MajorTickMark = eAxisTickMark.None;
                    c.Axis[1].MinorTickMark = eAxisTickMark.None;
                });

            //Style 2
            AddLine(ws, chartType, "Line3dChartStyle2", 0, 18, ePresetChartStyle.Line3dChartStyle2,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                    c.YAxis.Deleted = true;
                    c.Axis[0].MajorTickMark = eAxisTickMark.None;
                    c.Axis[0].MinorTickMark = eAxisTickMark.None;
                    c.Axis[1].MajorTickMark = eAxisTickMark.None;
                    c.Axis[1].MinorTickMark = eAxisTickMark.None;
                    if (chartType != eLineChartType.Line3D)
                    {
                        c.Marker = true;
                    }
                    c.DataLabel.ShowValue = true;
                    //foreach (var serie in c.Series)
                    //{
                    //    if (chartType != eLineChartType.Line3D)
                    //    {
                    //        serie.Marker.Style = eMarkerStyle.Circle;
                    //        serie.Marker.Size = 17;
                    //    }
                    //    serie.DataLabel.Position = eLabelPosition.Center;
                    //    serie.DataLabel.ShowValue = true;
                    //}
                });

            //Style 3
            AddLine(ws, chartType, "ColumnChartStyle3", 0, 31, ePresetChartStyle.Line3dChartStyle3,
            c =>
            {
            });

            //Style 4
            AddLine(ws, chartType, "ColumnChartStyle4", 22, 5, ePresetChartStyle.Line3dChartStyle4,
                c =>
                {
                });

        }
        private static ExcelLineChart AddLine(ExcelWorksheet ws, eLineChartType type, string name, int row, int col, ePresetChartStyle style, Action<ExcelLineChart> SetProperties)    
        {
            var chart = ws.Drawings.AddLineChart(name, type);
            chart.SetPosition(row, 0, col, 0);
            chart.To.Column = col+12;
            chart.To.ColumnOff = 0;
            chart.To.Row = row + 18;
            chart.To.RowOff = 0;
            var serie = chart.Series.Add("D2:D50", "A2:A50");
            SetProperties(chart);
            chart.StyleManager.SetChartStyle(style);

            return chart;
        }
        private static ExcelLineChart AddLine3D(ExcelWorksheet ws, eLineChartType type, string name, int row, int col, ePresetChartStyle style, Action<ExcelLineChart> SetProperties)
        {
            var chart = ws.Drawings.AddLineChart(name, type);
            chart.SetPosition(row, 0, col, 0);
            chart.To.Column = col + 12;
            chart.To.ColumnOff = 0;
            chart.To.Row = row + 18;
            chart.To.RowOff = 0;
            var serie = chart.Series.Add("D2:D50", "A2:A50");
            SetProperties(chart);
            chart.StyleManager.SetChartStyle(style);
            return chart;
        }
    }
}

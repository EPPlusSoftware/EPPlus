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
using System.Drawing;
using System.IO;
using System.Text;
using System.Xml;

namespace EPPlusTest.Drawing.Chart
{
    [TestClass]
    public class PieChartStylingTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("PieChartStyling.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            _pck.Save();
            _pck.Dispose();
        }
        [TestMethod]
        public void PieChart_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("PieChartStyles");
            LoadTestdata(ws);

            PieStyles(ws, ePieChartType.Pie);
        }
        [TestMethod]
        public void PieExplodedChart_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("PieExlodedChartStyles");
            LoadTestdata(ws);

            PieStyles(ws, ePieChartType.PieExploded);
        }

        private static void PieStyles(ExcelWorksheet ws, ePieChartType chartType)
        {
            //Style 1
            AddPie(ws, chartType, "PieChartStyle1", 0, 5, ePresetChartStyle.PieChartStyle1,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 2
            var chart2 = AddPie(ws, chartType, "PieChartStyle2", 0, 18, ePresetChartStyle.PieChartStyle2,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                });

            //Style 3
            var chart3 = AddPie(ws, chartType, "PieChartStyle3", 0, 31, ePresetChartStyle.PieChartStyle3,
            c =>
            {
                c.DataLabel.ShowPercent = true;
                c.DataLabel.Position = eLabelPosition.Center;
            });

            //Style 4
            AddPie(ws, chartType, "PieChartStyle4", 22, 5, ePresetChartStyle.PieChartStyle4,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 5
            AddPie(ws, chartType, "PieChartStyle5", 22, 18, ePresetChartStyle.PieChartStyle5,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 6
            AddPie(ws, chartType, "PieChartStyle6", 22, 31, ePresetChartStyle.PieChartStyle6,
                c =>
                {
                });

            //Style 7
            AddPie(ws, chartType, "PieChartStyle7", 44, 5, ePresetChartStyle.PieChartStyle7,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 8
            AddPie(ws, chartType, "PieChartStyle8", 44, 18, ePresetChartStyle.PieChartStyle8,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                    c.DataLabel.ShowPercent = true;
                    c.DataLabel.Position = eLabelPosition.InEnd;
                });

            //Style 9
            AddPie(ws, chartType, "PieChartStyle9", 44, 31, ePresetChartStyle.PieChartStyle9,
                c =>
                {
                    c.Legend.Remove();
                    c.DataLabel.ShowPercent = true;
                    c.DataLabel.ShowCategory = true;
                    c.DataLabel.ShowValue = true;
                    c.DataLabel.Position = eLabelPosition.OutEnd;
                });

            //Style 10
            AddPie(ws, chartType, "PieChartStyle10", 66, 5, ePresetChartStyle.PieChartStyle10,
                c =>
                {
                    c.Legend.Remove();
                    c.DataLabel.ShowValue = true;
                    c.DataLabel.ShowPercent = true;
                    c.DataLabel.ShowCategory = true;
                    c.DataLabel.Position = eLabelPosition.InEnd;
                });

            //Style 11
            AddPie(ws, chartType, "PieChartStyle11", 66, 18, ePresetChartStyle.PieChartStyle11,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                    c.DataLabel.ShowPercent = true;
                });

            //Style 12
            AddPie(ws, chartType, "PieChartStyle12", 66, 31, ePresetChartStyle.PieChartStyle12,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                    c.DataLabel.Position = eLabelPosition.InEnd;
                });
        }

        [TestMethod]
        public void PieChart3D_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("Pie3DChartStyles");
            LoadTestdata(ws);

            Pie3DStyles(ws, ePieChartType.Pie3D);
        }
        [TestMethod]
        public void PieExplodedChart3D_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("PieExploded3DChartStyles");
            LoadTestdata(ws);

            Pie3DStyles(ws, ePieChartType.PieExploded3D);
        }

        private static void Pie3DStyles(ExcelWorksheet ws, ePieChartType ePieChartType)
        {
            //Style 1
            AddPie(ws, ePieChartType, "PieChartStyle1", 0, 5, ePresetChartStyle.Pie3dChartStyle1,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 2
            var chart2 = AddPie(ws, ePieChartType, "PieChartStyle2", 0, 18, ePresetChartStyle.Pie3dChartStyle2,
                c =>
                {
                    c.Legend.Remove();
                    c.DataLabel.ShowCategory = true;
                    c.DataLabel.ShowValue = true;
                    c.View3D.RotY = 50;
                    c.View3D.RotX = 50;
                    c.View3D.DepthPercent = 100;
                    c.View3D.RightAngleAxes = false;
                });

            //Style 3
            AddPie(ws, ePieChartType, "PieChartStyle3", 0, 31, ePresetChartStyle.Pie3dChartStyle3,
            c =>
            {
                c.DataLabel.ShowPercent = true;
                c.DataLabel.Position = eLabelPosition.Center;
            });

            //Style 4
            AddPie(ws, ePieChartType, "PieChartStyle4", 22, 5, ePresetChartStyle.Pie3dChartStyle4,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 5
            AddPie(ws, ePieChartType, "PieChartStyle5", 22, 18, ePresetChartStyle.Pie3dChartStyle5,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 6
            AddPie(ws, ePieChartType, "PieChartStyle6", 22, 31, ePresetChartStyle.Pie3dChartStyle6,
                c =>
                {
                });

            //Style 7
            AddPie(ws, ePieChartType, "PieChartStyle7", 44, 5, ePresetChartStyle.Pie3dChartStyle7,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 8
            AddPie(ws, ePieChartType, "PieChartStyle8", 44, 18, ePresetChartStyle.Pie3dChartStyle8,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                    c.DataLabel.ShowPercent = true;
                    c.DataLabel.Position = eLabelPosition.InEnd;
                });

            //Style 9
            AddPie(ws, ePieChartType, "PieChartStyle9", 44, 31, ePresetChartStyle.Pie3dChartStyle9,
                c =>
                {
                    c.Legend.Remove();
                    c.DataLabel.ShowPercent = true;
                    c.DataLabel.ShowCategory = true;
                    c.DataLabel.ShowValue = true;
                    c.DataLabel.Position = eLabelPosition.OutEnd;
                });

            //Style 10
            AddPie(ws, ePieChartType, "PieChartStyle10", 66, 5, ePresetChartStyle.Pie3dChartStyle10,
                c =>
                {
                    c.Legend.Remove();
                    c.DataLabel.ShowValue = true;
                    c.DataLabel.ShowPercent = true;
                    c.DataLabel.ShowCategory = true;
                    c.DataLabel.Position = eLabelPosition.InEnd;
                });
        }

        private static ExcelPieChart AddPie(ExcelWorksheet ws, ePieChartType type, string name, int row, int col, ePresetChartStyle style, Action<ExcelPieChart> SetProperties)    
        {
            var chart = ws.Drawings.AddPieChart(name, type);
            chart.SetPosition(row, 0, col, 0);
            chart.To.Column = col+12;
            chart.To.ColumnOff = 0;
            chart.To.Row = row + 18;
            chart.To.RowOff = 0;
            var serie = chart.Series.Add("D2:D8", "A2:A8");
            var dp=serie.DataPoints.Add(3);
            dp.Border.Fill.Style = OfficeOpenXml.Drawing.eFillStyle.SolidFill;
            dp.Border.Fill.SolidFill.Color.SetRgbColor(Color.Black);
            SetProperties(chart);

            chart.StyleManager.SetChartStyle(style);
            return chart;
        }
    }
}


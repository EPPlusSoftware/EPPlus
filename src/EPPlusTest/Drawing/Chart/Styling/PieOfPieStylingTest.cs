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
namespace EPPlusTest.Drawing.Chart.Styling
{
    [TestClass]
    public class OfPieChartStylingTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("OfPieChartStyling.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void PieOfPie_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("PieOfPieChartStyles");
            LoadTestdata(ws);

            OfPieStyles(ws, eOfPieChartType.PieOfPie);
        }
        [TestMethod]
        public void BarOfPie_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("BarOfPieChartStyles");
            LoadTestdata(ws);

            OfPieStyles(ws, eOfPieChartType.BarOfPie);
        }

        private static void OfPieStyles(ExcelWorksheet ws, eOfPieChartType chartType)
        {
            //Style 1
            AddOfPie(ws, chartType, "OfPieChartStyle1", 0, 5, ePresetChartStyle.OfPieChartStyle1,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 2
            AddOfPie(ws, chartType, "OfPieChartStyle2", 0, 18, ePresetChartStyle.OfPieChartStyle2,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                });

            //Style 3
            AddOfPie(ws, chartType, "OfPieChartStyle3", 0, 31, ePresetChartStyle.OfPieChartStyle3,
            c =>
            {
                c.DataLabel.ShowPercent = true;
            });

            //Style 4
            AddOfPie(ws, chartType, "OfPieChartStyle4", 22, 5, ePresetChartStyle.OfPieChartStyle4,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 5
            AddOfPie(ws, chartType, "OfPieChartStyle5", 22, 18, ePresetChartStyle.OfPieChartStyle5,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 6
            AddOfPie(ws, chartType, "OfPieChartStyle6", 22, 31, ePresetChartStyle.OfPieChartStyle6,
                c =>
                {
                });

            //Style 7
            AddOfPie(ws, chartType, "OfPieChartStyle7", 44, 5, ePresetChartStyle.OfPieChartStyle7,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 8
            AddOfPie(ws, chartType, "OfPieChartStyle8", 44, 18, ePresetChartStyle.OfPieChartStyle8,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                    c.DataLabel.ShowPercent = true;
                });

            //Style 9
            AddOfPie(ws, chartType, "OfPieChartStyle9", 44, 31, ePresetChartStyle.OfPieChartStyle9,
                c =>
                {
                    c.Legend.Remove();
                    c.DataLabel.ShowValue = true;
                    c.DataLabel.ShowPercent = true;
                    c.DataLabel.ShowCategory = true;
                });

            //Style 10
            AddOfPie(ws, chartType, "OfPieChartStyle10", 66, 5, ePresetChartStyle.OfPieChartStyle10,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                    c.DataLabel.ShowPercent = true;
                });

            //Style 11
            AddOfPie(ws, chartType, "OfPieChartStyle11", 66, 18, ePresetChartStyle.OfPieChartStyle11,
                c =>
                {
                });

            //Style 12
            AddOfPie(ws, chartType, "OfPieChartStyle12", 66, 31, ePresetChartStyle.OfPieChartStyle12,
                c =>
                {
                });
        }


        private static ExcelOfPieChart AddOfPie(ExcelWorksheet ws, eOfPieChartType type, string name, int row, int col, ePresetChartStyle style, Action<ExcelOfPieChart> SetProperties)    
        {
            var chart = ws.Drawings.AddOfPieChart(name, type);
            chart.SetPosition(row, 0, col, 0);
            chart.To.Column = col+12;
            chart.To.ColumnOff = 0;
            chart.To.Row = row + 18;
            chart.To.RowOff = 0;
            var serie = chart.Series.Add("D2:D8", "A2:A8");
            serie.DataPoints.Add(0);
            serie.DataPoints.Add(1);
            serie.DataPoints.Add(2);
            serie.DataPoints.Add(3);
            serie.DataPoints.Add(4);
            serie.DataPoints.Add(5);
            serie.DataPoints.Add(6);
            serie.DataPoints.Add(7);

            SetProperties(chart);

            chart.StyleManager.SetChartStyle(style);
            return chart;
        }
    }
}

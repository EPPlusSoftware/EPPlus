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
    public class DoughnutChartStylingTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("DoughnutChartStyling.xlsx", true);
            ExcelChartStyleManager.LoadStyles(new DirectoryInfo(@"c:\temp\"));
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void DoughnutChart_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("DoughnutChartStyles");
            LoadTestdata(ws);

            DoughnutStyles(ws, eDoughnutChartType.Doughnut);
        }
        [TestMethod]
        public void DoughnutExplodedChart_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("PieExlodedChartStyles");
            LoadTestdata(ws);

            DoughnutStyles(ws, eDoughnutChartType.DoughnutExploded);
        }

        private static void DoughnutStyles(ExcelWorksheet ws, eDoughnutChartType chartType)
        {
            //Style 1
            AddDoughnut(ws, chartType, "DoughnutChartStyle1", 0, 5, ePresetChartStyle.DoughnutChartStyle1,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 2
            AddDoughnut(ws, chartType, "DoughnutChartStyle2", 0, 18, ePresetChartStyle.DoughnutChartStyle2,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                });

            //Style 3
            AddDoughnut(ws, chartType, "DoughnutChartStyle3", 0, 31, ePresetChartStyle.DoughnutChartStyle3,
            c =>
            {
                c.DataLabel.ShowPercent = true;
            });

            //Style 4
            AddDoughnut(ws, chartType, "DoughnutChartStyle4", 22, 5, ePresetChartStyle.DoughnutChartStyle4,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 5
            AddDoughnut(ws, chartType, "DoughnutChartStyle5", 22, 18, ePresetChartStyle.DoughnutChartStyle5,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 6
            AddDoughnut(ws, chartType, "DoughnutChartStyle6", 22, 31, ePresetChartStyle.DoughnutChartStyle6,
                c =>
                {
                });

            //Style 7
            AddDoughnut(ws, chartType, "DoughnutChartStyle7", 44, 5, ePresetChartStyle.DoughnutChartStyle7,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 8
            AddDoughnut(ws, chartType, "DoughnutChartStyle8", 44, 18, ePresetChartStyle.DoughnutChartStyle8,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                    c.DataLabel.ShowPercent = true;
                });

            //Style 9
            AddDoughnut(ws, chartType, "DoughnutChartStyle9", 44, 31, ePresetChartStyle.DoughnutChartStyle9,
                c =>
                {
                    c.Legend.Remove();
                    c.DataLabel.ShowValue = true;
                    c.DataLabel.ShowPercent = true;
                    c.DataLabel.ShowCategory = true;
                });

            //Style 10
            AddDoughnut(ws, chartType, "DoughnutChartStyle10", 66, 5, ePresetChartStyle.DoughnutChartStyle10,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                    c.DataLabel.ShowPercent = true;
                });
        }


        private static ExcelDoughnutChart AddDoughnut(ExcelWorksheet ws, eDoughnutChartType type, string name, int row, int col, ePresetChartStyle style, Action<ExcelDoughnutChart> SetProperties)    
        {
            var chart = ws.Drawings.AddDoughnutChart(name, type);
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

            SetProperties(chart);

            chart.StyleManager.SetChartStyle(style);
            return chart;
        }
    }
}

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
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.Drawing.Chart.Style;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Xml;

namespace EPPlusTest.Drawing.Chart.Styling
{
    [TestClass]
    public class FunnelChartStylingTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("FunnelChartStyling.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void FunnelChart_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("FunnelChart");
            LoadTestdata(ws);
            FunnelChartStyle(ws);
        }
        private static void FunnelChartStyle(ExcelWorksheet ws)
        {
            //Funnel Chart styles

            //Funnel chart Style 1
            AddChart(ws, ePresetChartStyle.FunnelChartStyle1, "FunnelChartStyle1", 0, 5,
                c =>
                {
                    c.Title.Text = "Funnel 1";
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Funnel chart Style 2
            AddChart(ws, ePresetChartStyle.FunnelChartStyle2, "FunnelChartStyle2", 0, 18,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Funnel chart Style 3
            AddChart(ws, ePresetChartStyle.FunnelChartStyle3, "FunnelChartStyle3", 0, 31,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Funnel chart Style 4
            AddChart(ws, ePresetChartStyle.FunnelChartStyle4, "FunnelChartStyle4", 20, 5,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Funnel chart Style 5
            AddChart(ws, ePresetChartStyle.FunnelChartStyle5, "FunnelChartStyle5", 20, 18,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Funnel chart Style 6
            AddChart(ws, ePresetChartStyle.FunnelChartStyle6, "FunnelChartStyle6", 20, 31,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Funnel chart Style 7
            AddChart(ws, ePresetChartStyle.FunnelChartStyle7, "FunnelChartStyle7", 40, 5,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Funnel chart Style 8
            AddChart(ws, ePresetChartStyle.FunnelChartStyle8, "FunnelChartStyle8", 40, 18,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Funnel chart Style 9
            AddChart(ws, ePresetChartStyle.FunnelChartStyle9, "FunnelChartStyle9", 40, 31,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });
        }


        private static ExcelFunnelChart AddChart(ExcelWorksheet ws, ePresetChartStyle style, string name, int row, int col, Action<ExcelFunnelChart> SetProperties)
        {
            var chart = ws.Drawings.AddFunnelChart(name);
            chart.SetPosition(row, 0, col, 0);
            chart.To.Column = col+12;
            chart.To.ColumnOff = 0;
            chart.To.Row = row + 18;
            chart.To.RowOff = 0;
            chart.Series.Add("D2:D8", "A2:A8");

            SetProperties(chart);

            chart.StyleManager.SetChartStyle(style);
            return chart;
        }
    }
}

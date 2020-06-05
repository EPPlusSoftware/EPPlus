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
    public class BoxWhiskerChartStylingTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("BoxWhiskerChartStyling.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void BoxWhiskerChart_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("BoxWhiskerChart");
            LoadTestdata(ws);
            BoxWhiskerChartStyle(ws);
        }
        private static void BoxWhiskerChartStyle(ExcelWorksheet ws)
        {
            //Box & Whisker Chart styles

            //Box & Whisker chart Style 1
            AddChartEx(ws, ePresetChartStyle.BoxWhiskerChartStyle1, "BoxWhiskerChartStyle1", 0, 5, 
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                    c.Legend.PositionAlignment = ePositionAlign.Min;
                });

            //Box & Whisker chart Style 2
            AddChartEx(ws, ePresetChartStyle.BoxWhiskerChartStyle2, "BoxWhiskerStyle2", 0, 18, 
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                    c.Series[0].ShowOutliers = false;
                    c.Series[0].ShowMeanLine = false;
                    c.Series[0].ShowMeanMarker = true;
                    c.Series[0].ShowNonOutliers = true;
                });

            //Box & Whisker chart Style 3
            AddChartEx(ws, ePresetChartStyle.BoxWhiskerChartStyle3, "BoxWhiskerChartStyle3", 0, 31, 
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                    c.Series[0].ShowMeanMarker = false;
                    c.Series[0].ShowNonOutliers = false;
                    c.Series[0].ShowOutliers = true;
                    c.Series[0].ShowMeanLine = true;
                });

            //Box & Whisker chart Style 4
            AddChartEx(ws, ePresetChartStyle.BoxWhiskerChartStyle4, "BoxWhiskerChartStyle4", 20, 5, 
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Box & Whisker chart Style 5
            AddChartEx(ws, ePresetChartStyle.BoxWhiskerChartStyle5, "BoxWhiskerChartStyle5", 20, 18, 
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Box & Whisker chart Style 6
            AddChartEx(ws, ePresetChartStyle.BoxWhiskerChartStyle6, "BoxWhiskerChartStyle6", 20, 31, 
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

        }
        private static ExcelBoxWhiskerChart AddChartEx(ExcelWorksheet ws, ePresetChartStyle style, string name, int row, int col,Action<ExcelBoxWhiskerChart> SetProperties)
        {
            var chart = ws.Drawings.AddBoxWhiskerChart(name);
            chart.SetPosition(row, 0, col, 0);
            chart.To.Column = col+12;
            chart.To.ColumnOff = 0;
            chart.To.Row = row + 18;
            chart.To.RowOff = 0;
            chart.Series.Add("A2:A8", "D2:D8");

            SetProperties(chart);

            chart.StyleManager.SetChartStyle(style);
            return chart;
        }
    }
}

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
    public class BubbleChartStylingTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("BubbleChartStyling.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            _pck.Save();
            _pck.Dispose();
        }
        [TestMethod]
        public void Bubble_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("BubbleChartStyling");
            LoadTestdata(ws);

            BubbleStyle(ws, eBubbleChartType.Bubble);
        }
        [TestMethod]
        public void Bubble3D_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("BarOfPieChartStyles");
            LoadTestdata(ws);

            Bubble3DStyle(ws, eBubbleChartType.Bubble3DEffect);
        }

        private static void BubbleStyle(ExcelWorksheet ws, eBubbleChartType chartType)
        {
            //Style 1
            AddBubble(ws, chartType, "BubbleChartStyle1", 0, 5, ePresetChartStyle.BubbleChartStyle1,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 2
            AddBubble(ws, chartType, "BubbleChartStyle2", 0, 18, ePresetChartStyle.BubbleChartStyle2,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                });

            //Style 3
            AddBubble(ws, chartType, "BubbleChartStyle3", 0, 31, ePresetChartStyle.BubbleChartStyle3,
            c =>
            {
                c.DataLabel.ShowPercent = true;
            });

            //Style 4
            AddBubble(ws, chartType, "BubbleChartStyle4", 22, 5, ePresetChartStyle.BubbleChartStyle4,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 5
            AddBubble(ws, chartType, "BubbleChartStyle5", 22, 18, ePresetChartStyle.BubbleChartStyle5,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 6
            AddBubble(ws, chartType, "BubbleChartStyle6", 22, 31, ePresetChartStyle.BubbleChartStyle6,
                c =>
                {
                });

            //Style 7
            AddBubble(ws, chartType, "BubbleChartStyle7", 44, 5, ePresetChartStyle.BubbleChartStyle7,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 8
            AddBubble(ws, chartType, "BubbleChartStyle8", 44, 18, ePresetChartStyle.BubbleChartStyle8,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                    c.DataLabel.ShowPercent = true;
                });

            //Style 9
            AddBubble(ws, chartType, "BubbleChartStyle9", 44, 31, ePresetChartStyle.BubbleChartStyle9,
                c =>
                {
                    c.Legend.Remove();
                    c.DataLabel.ShowValue = true;
                    c.DataLabel.ShowPercent = true;
                    c.DataLabel.ShowCategory = true;
                });

            //Style 10
            AddBubble(ws, chartType, "BubbleChartStyle10", 66, 5, ePresetChartStyle.BubbleChartStyle10,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                    c.DataLabel.ShowPercent = true;
                });

            //Style 11
            AddBubble(ws, chartType, "BubbleChartStyle11", 66, 18, ePresetChartStyle.BubbleChartStyle11,
                c =>
                {
                });
        }
        private static void Bubble3DStyle(ExcelWorksheet ws, eBubbleChartType chartType)
        {
            //Style 1
            AddBubble(ws, chartType, "BubbleChartStyle1", 0, 5, ePresetChartStyle.Bubble3dChartStyle1,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 2
            AddBubble(ws, chartType, "BubbleChartStyle2", 0, 18, ePresetChartStyle.Bubble3dChartStyle2,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                });

            //Style 3
            AddBubble(ws, chartType, "BubbleChartStyle3", 0, 31, ePresetChartStyle.Bubble3dChartStyle3,
            c =>
            {
                c.DataLabel.ShowPercent = true;
            });

            //Style 4
            AddBubble(ws, chartType, "BubbleChartStyle4", 22, 5, ePresetChartStyle.Bubble3dChartStyle4,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 5
            AddBubble(ws, chartType, "BubbleChartStyle5", 22, 18, ePresetChartStyle.Bubble3dChartStyle5,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 6
            AddBubble(ws, chartType, "BubbleChartStyle6", 22, 31, ePresetChartStyle.Bubble3dChartStyle6,
                c =>
                {
                });

            //Style 7
            AddBubble(ws, chartType, "BubbleChartStyle7", 44, 5, ePresetChartStyle.Bubble3dChartStyle7,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Style 8
            AddBubble(ws, chartType, "BubbleChartStyle8", 44, 18, ePresetChartStyle.Bubble3dChartStyle8,
                c =>
                {
                    c.Legend.Position = eLegendPosition.Top;
                    c.DataLabel.ShowPercent = true;
                });

            //Style 9
            AddBubble(ws, chartType, "BubbleChartStyle9", 44, 31, ePresetChartStyle.Bubble3dChartStyle9,
                c =>
                {
                    c.Legend.Remove();
                    c.DataLabel.ShowValue = true;
                    c.DataLabel.ShowPercent = true;
                    c.DataLabel.ShowCategory = true;
                });
        }


        private static ExcelBubbleChart AddBubble(ExcelWorksheet ws, eBubbleChartType type, string name, int row, int col, ePresetChartStyle style, Action<ExcelBubbleChart> SetProperties)    
        {
            var chart = ws.Drawings.AddBubbleChart(name, type);
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

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
using System.CodeDom;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Xml;

namespace EPPlusTest.Drawing.Chart
{
    [TestClass]
    public class StockChartStylingTest : StockChartTestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("StockChartStyling.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void StockChartHLC_Styles()
        {
            var ws = _pck.Workbook.Worksheets.Add("StockChartStyling");
            var members = new MemberInfo[] 
            {
                typeof(PeriodData).GetProperty("Date"),
                typeof(PeriodData).GetProperty("HighPrice"),
                typeof(PeriodData).GetProperty("LowPrice"),
                typeof(PeriodData).GetProperty("ClosePrice"),
            };

            LoadStockChartDataPeriod(ws, members);

            StockChartStyle(ws, eStockChartType.StockHLC);
        }
        private static void StockChartStyle(ExcelWorksheet ws, eStockChartType chartType)
        {
            //Surface charts don't use chart styling in Excel, but styles can be applied anyway. 

            //Stock chart Style 1
            AddStockChart(ws, chartType, ePresetChartStyle.StockChartStyle1, "StockChartStyle1", 0, 5, ws.Cells["A1:D7"],
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                }) ;

            //Stock chart Style 2
            AddStockChart(ws, chartType, ePresetChartStyle.StockChartStyle2, "StockChartStyle2", 0, 18, ws.Cells["A1:D7"],
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Stock chart Style 3
            AddStockChart(ws, chartType, ePresetChartStyle.StockChartStyle3, "StockChartStyle3", 0, 31, ws.Cells["A1:D7"],
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Stock chart Style 4
            AddStockChart(ws, chartType, ePresetChartStyle.StockChartStyle4, "StockChartStyle4", 20, 5, ws.Cells["A1:D7"],
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Stock chart Style 5
            AddStockChart(ws, chartType, ePresetChartStyle.StockChartStyle5, "StockChartStyle5", 20, 18, ws.Cells["A1:D7"],
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Stock chart Style 6
            AddStockChart(ws, chartType, ePresetChartStyle.StockChartStyle6, "StockChartStyle6", 20, 31, ws.Cells["A1:D7"],
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Stock chart Style 7
            AddStockChart(ws, chartType, ePresetChartStyle.StockChartStyle7, "StockChartStyle7", 40, 5, ws.Cells["A1:D7"],
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Stock chart Style 8
            AddStockChart(ws, chartType, ePresetChartStyle.StockChartStyle8, "StockChartStyle8", 40, 18, ws.Cells["A1:D7"],
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Stock chart Style 9
            AddStockChart(ws, chartType, ePresetChartStyle.StockChartStyle9, "StockChartStyle9", 40, 31, ws.Cells["A1:D7"],
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });


            //Stock chart Style 10
            AddStockChart(ws, chartType, ePresetChartStyle.StockChartStyle10, "StockChartStyle10", 60, 18, ws.Cells["A1:D7"],
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });

            //Stock chart Style 11
            AddStockChart(ws, chartType, ePresetChartStyle.StockChartStyle11, "StockChartStyle11", 60, 31, ws.Cells["A1:D7"],
                c =>
                {
                    c.Legend.Position = eLegendPosition.Bottom;
                });
        }


        private static ExcelStockChart AddStockChart(ExcelWorksheet ws, eStockChartType type, ePresetChartStyle style, string name, int row, int col,ExcelRange range, Action<ExcelStockChart> SetProperties)    
        {
            var chart = ws.Drawings.AddStockChart(name, type, range);
            chart.SetPosition(row, 0, col, 0);
            chart.To.Column = col+12;
            chart.To.ColumnOff = 0;
            chart.To.Row = row + 18;
            chart.To.RowOff = 0;

            SetProperties(chart);

            chart.StyleManager.SetChartStyle(style);
            return chart;
        }
    }
}

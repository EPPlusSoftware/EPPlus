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
using EPPlusTest.Properties;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Style;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Drawing.Theme;
using System.Drawing;

namespace EPPlusTest.Drawing.Chart
{
    [TestClass]
    public class ErrorBarsTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("ErrorBars.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void ErrorBars_StdDev()
        {
            var ws = _pck.Workbook.Worksheets.Add("ErrorBar_StDev");
            LoadTestdata(ws);

            var chart = ws.Drawings.AddLineChart("LineChart1", eLineChartType.Line);
            Assert.AreEqual(eDrawingType.Chart, chart.DrawingType);
            var serie = chart.Series.Add("D2:D100", "A2:A100");
            serie.AddErrorBars(eErrorBarType.Both, eErrorValueType.StandardDeviation);
            serie.ErrorBars.Direction = eErrorBarDirection.Y;
            serie.ErrorBars.Value = 14;
            chart.SetPosition(1, 0, 5, 0);

            Assert.AreEqual(eErrorBarType.Both, serie.ErrorBars.BarType);
            Assert.AreEqual(eErrorValueType.StandardDeviation, serie.ErrorBars.ValueType);
            Assert.AreEqual(eErrorBarDirection.Y, serie.ErrorBars.Direction);
            Assert.AreEqual(false, serie.ErrorBars.NoEndCap);
        }
        [TestMethod]
        public void ErrorBars_StdErr()
        {
            var ws = _pck.Workbook.Worksheets.Add("ErrorBar_StErr");
            LoadTestdata(ws);

            var chart = ws.Drawings.AddLineChart("LineChart1", eLineChartType.Line);
            var serie = chart.Series.Add("D2:D100", "A2:A100");
            serie.AddErrorBars(eErrorBarType.Both, eErrorValueType.StandardError);
            serie.ErrorBars.Direction = eErrorBarDirection.X;
            serie.ErrorBars.NoEndCap = true;
            chart.SetPosition(1, 0, 5, 0);

            Assert.AreEqual(eErrorBarType.Both, serie.ErrorBars.BarType);
            Assert.AreEqual(eErrorValueType.StandardError, serie.ErrorBars.ValueType);
            Assert.AreEqual(eErrorBarDirection.X, serie.ErrorBars.Direction);
            Assert.AreEqual(true, serie.ErrorBars.NoEndCap);
        }
        [TestMethod]
        public void ErrorBars_Percentage()
        {
            var ws = _pck.Workbook.Worksheets.Add("ErrorBar_Percentage");
            LoadTestdata(ws);

            var chart = ws.Drawings.AddLineChart("LineChart1", eLineChartType.Line);
            var serie = chart.Series.Add("D2:D100", "A2:A100");
            serie.AddErrorBars(eErrorBarType.Plus, eErrorValueType.Percentage);
            serie.ErrorBars.Value = 5;
            chart.SetPosition(1, 0, 5, 0);

            Assert.AreEqual(eErrorBarType.Plus, serie.ErrorBars.BarType);
            Assert.AreEqual(eErrorValueType.Percentage, serie.ErrorBars.ValueType);
            Assert.AreEqual(eErrorBarDirection.Y, serie.ErrorBars.Direction);
            Assert.AreEqual(false, serie.ErrorBars.NoEndCap);
        }
        [TestMethod]
        public void ErrorBars_Fixed()
        {
            var ws = _pck.Workbook.Worksheets.Add("ErrorBar_Fixed");
            LoadTestdata(ws);

            var chart = ws.Drawings.AddLineChart("LineChart1", eLineChartType.Line);
            var serie = chart.Series.Add("D2:D100", "A2:A100");
            serie.AddErrorBars(eErrorBarType.Plus, eErrorValueType.FixedValue);
            serie.ErrorBars.Value = 5.2;
            chart.SetPosition(1, 0, 5, 0);

            Assert.AreEqual(eErrorBarType.Plus, serie.ErrorBars.BarType);
            Assert.AreEqual(eErrorValueType.FixedValue, serie.ErrorBars.ValueType);
            Assert.AreEqual(false, serie.ErrorBars.NoEndCap);
        }
        [TestMethod]
        public void ErrorBars_Custom()
        {
            var ws = _pck.Workbook.Worksheets.Add("ErrorBar_Custom");
            LoadTestdata(ws);

            var chart = ws.Drawings.AddLineChart("LineChart1", eLineChartType.Line);
            var serie = chart.Series.Add("D2:D100", "A2:A100");
            serie.AddErrorBars(eErrorBarType.Plus, eErrorValueType.Custom);
            serie.ErrorBars.Plus.ValuesSource = "{1}";
            serie.ErrorBars.Minus.FormatCode = "General";
            serie.ErrorBars.Minus.ValuesSource = "=A2:A15";
            serie.ErrorBars.Minus.FormatCode = "0";
            chart.SetPosition(1, 0, 5, 0);

            Assert.AreEqual(eErrorBarType.Plus, serie.ErrorBars.BarType);
            Assert.AreEqual(eErrorValueType.Custom, serie.ErrorBars.ValueType);
            Assert.AreEqual(false, serie.ErrorBars.NoEndCap);

            Assert.AreEqual("{1}", serie.ErrorBars.Plus.ValuesSource);
            Assert.AreEqual("A2:A15", serie.ErrorBars.Minus.ValuesSource);
        }

    }
}

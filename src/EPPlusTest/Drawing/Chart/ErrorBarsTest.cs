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
        [TestMethod]
        public void ErrorBars_Scatter()
        {
            var ws = _pck.Workbook.Worksheets.Add("ErrorBarScatter");
            LoadTestdata(ws);

            var chart = ws.Drawings.AddScatterChart("ScatterChart1", eScatterChartType.XYScatter);
            var serie = chart.Series.Add("D2:D100", "A2:A100");
            serie.AddErrorBars(eErrorBarType.Both, eErrorValueType.Custom);
            serie.ErrorBars.Plus.ValuesSource = "{1}";
            serie.ErrorBars.Minus.FormatCode = "General";
            serie.ErrorBars.Minus.ValuesSource = "=ErrorBarScatter!$A$2:$A$15";
            serie.ErrorBars.Minus.FormatCode = "0";

            serie.ErrorBarsX.Plus.ValuesSource = "{2}";
            serie.ErrorBarsX.Minus.FormatCode = "General";
            serie.ErrorBarsX.Minus.ValuesSource = "=ErrorBarScatter!$A$2:$A$16";
            serie.ErrorBarsX.Minus.FormatCode = "0";

            chart.SetPosition(1, 0, 5, 0);

            Assert.IsNotNull(serie.ErrorBars);
            Assert.IsNotNull(serie.ErrorBarsX);
            Assert.AreEqual(eErrorBarType.Both, serie.ErrorBars.BarType);
            Assert.AreEqual(eErrorValueType.Custom, serie.ErrorBars.ValueType);
            Assert.AreEqual(false, serie.ErrorBars.NoEndCap);

            Assert.AreEqual("{1}", serie.ErrorBars.Plus.ValuesSource);
            Assert.AreEqual("ErrorBarScatter!$A$2:$A$15", serie.ErrorBars.Minus.ValuesSource);

            Assert.AreEqual(eErrorBarType.Both, serie.ErrorBarsX.BarType);
            Assert.AreEqual(eErrorValueType.Custom, serie.ErrorBarsX.ValueType);
            Assert.AreEqual(false, serie.ErrorBarsX.NoEndCap);

            Assert.AreEqual("{2}", serie.ErrorBarsX.Plus.ValuesSource);
            Assert.AreEqual("ErrorBarScatter!$A$2:$A$16", serie.ErrorBarsX.Minus.ValuesSource);
        }
        [TestMethod]
        public void ErrorBars_ReadScatter()
        {
            using (var p1 = new ExcelPackage())
            {
                var ws = p1.Workbook.Worksheets.Add("ErrorBarsScatter");
                LoadTestdata(ws);

                var chart = ws.Drawings.AddScatterChart("ScatterChart1", eScatterChartType.XYScatter);
                var serie = chart.Series.Add("D2:D100", "A2:A100");
                serie.AddErrorBars(eErrorBarType.Both, eErrorValueType.Custom);
                serie.ErrorBars.Plus.ValuesSource = "{1}";
                serie.ErrorBars.Minus.FormatCode = "General";
                serie.ErrorBars.Minus.ValuesSource = "=ErrorBarsScatter!$A$2:$A$15";
                serie.ErrorBars.Minus.FormatCode = "0";

                serie.ErrorBarsX.Plus.ValuesSource = "{2}";
                serie.ErrorBarsX.Minus.FormatCode = "General";
                serie.ErrorBarsX.Minus.ValuesSource = "ErrorBarsScatter!$A$2:$A$16";
                serie.ErrorBarsX.Minus.FormatCode = "0";

                chart.SetPosition(1, 0, 5, 0);
                p1.Save();
                using(var p2=new ExcelPackage(p1.Stream))
                {
                    ws = p2.Workbook.Worksheets["ErrorBarsScatter"];

                    chart = ws.Drawings[0].As.Chart.ScatterChart;
                    serie = chart.Series[0];

                    Assert.IsNotNull(serie.ErrorBars);
                    Assert.IsNotNull(serie.ErrorBarsX);
                    Assert.AreEqual(eErrorBarType.Both, serie.ErrorBars.BarType);
                    Assert.AreEqual(eErrorValueType.Custom, serie.ErrorBars.ValueType);
                    Assert.AreEqual(false, serie.ErrorBars.NoEndCap);

                    Assert.AreEqual("{1}", serie.ErrorBars.Plus.ValuesSource);
                    Assert.AreEqual("ErrorBarsScatter!$A$2:$A$15", serie.ErrorBars.Minus.ValuesSource);

                    Assert.AreEqual(eErrorBarType.Both, serie.ErrorBarsX.BarType);
                    Assert.AreEqual(eErrorValueType.Custom, serie.ErrorBarsX.ValueType);
                    Assert.AreEqual(false, serie.ErrorBarsX.NoEndCap);

                    Assert.AreEqual("{2}", serie.ErrorBarsX.Plus.ValuesSource);
                    Assert.AreEqual("ErrorBarsScatter!$A$2:$A$16", serie.ErrorBarsX.Minus.ValuesSource);
                }
            }
        }
        [TestMethod]
        public void ErrorBars_Bubble()
        {
            var ws = _pck.Workbook.Worksheets.Add("ErrorBar_Bubble");
            LoadTestdata(ws);

            var chart = ws.Drawings.AddBubbleChart("BubbleChart1", eBubbleChartType.Bubble);
            var serie = chart.Series.Add("D2:D100", "A2:A100");
            serie.AddErrorBars(eErrorBarType.Both, eErrorValueType.Custom);
            serie.ErrorBars.Plus.ValuesSource = "{1}";
            serie.ErrorBars.Minus.FormatCode = "General";
            serie.ErrorBars.Minus.ValuesSource = "=ErrorBar_Bubble!$A$2:$A$15";
            serie.ErrorBars.Minus.FormatCode = "0";

            serie.ErrorBarsX.Plus.ValuesSource = "{2}";
            serie.ErrorBarsX.Minus.FormatCode = "General";
            serie.ErrorBarsX.Minus.ValuesSource = "=ErrorBar_Bubble!$A$2:$A$16";
            serie.ErrorBarsX.Minus.FormatCode = "0";

            chart.SetPosition(1, 0, 5, 0);

            Assert.IsNotNull(serie.ErrorBars);
            Assert.IsNotNull(serie.ErrorBarsX);
            Assert.AreEqual(eErrorBarType.Both, serie.ErrorBars.BarType);
            Assert.AreEqual(eErrorValueType.Custom, serie.ErrorBars.ValueType);
            Assert.AreEqual(false, serie.ErrorBars.NoEndCap);

            Assert.AreEqual("{1}", serie.ErrorBars.Plus.ValuesSource);
            Assert.AreEqual("ErrorBar_Bubble!$A$2:$A$15", serie.ErrorBars.Minus.ValuesSource);

            Assert.AreEqual(eErrorBarType.Both, serie.ErrorBarsX.BarType);
            Assert.AreEqual(eErrorValueType.Custom, serie.ErrorBarsX.ValueType);
            Assert.AreEqual(false, serie.ErrorBarsX.NoEndCap);

            Assert.AreEqual("{2}", serie.ErrorBarsX.Plus.ValuesSource);
            Assert.AreEqual("ErrorBar_Bubble!$A$2:$A$16", serie.ErrorBarsX.Minus.ValuesSource);
        }
        [TestMethod]
        public void ErrorBars_ReadBubble()
        {
            using (var p1 = new ExcelPackage())
            {
                var ws = p1.Workbook.Worksheets.Add("ErrorBars");
                LoadTestdata(ws);

                var chart = ws.Drawings.AddBubbleChart("BubbleChart1", eBubbleChartType.Bubble);
                var serie = chart.Series.Add("D2:D100", "A2:A100");
                serie.AddErrorBars(eErrorBarType.Both, eErrorValueType.Custom);
                serie.ErrorBars.Plus.ValuesSource = "{1}";
                serie.ErrorBars.Minus.FormatCode = "General";
                serie.ErrorBars.Minus.ValuesSource = "=ErrorBars!$A$2:$A$15";
                serie.ErrorBars.Minus.FormatCode = "0";

                serie.ErrorBarsX.Plus.ValuesSource = "{2}";
                serie.ErrorBarsX.Minus.FormatCode = "General";
                serie.ErrorBarsX.Minus.ValuesSource = "=ErrorBars!$A$2:$A$16";
                serie.ErrorBarsX.Minus.FormatCode = "0";

                chart.SetPosition(1, 0, 5, 0);
                p1.Save();
                using (var p2 = new ExcelPackage(p1.Stream))
                {
                    ws = p2.Workbook.Worksheets["ErrorBars"];

                    chart = ws.Drawings[0].As.Chart.BubbleChart;
                    serie = chart.Series[0];

                    Assert.IsNotNull(serie.ErrorBars);
                    Assert.IsNotNull(serie.ErrorBarsX);
                    Assert.AreEqual(eErrorBarType.Both, serie.ErrorBars.BarType);
                    Assert.AreEqual(eErrorValueType.Custom, serie.ErrorBars.ValueType);
                    Assert.AreEqual(false, serie.ErrorBars.NoEndCap);

                    Assert.AreEqual("{1}", serie.ErrorBars.Plus.ValuesSource);
                    Assert.AreEqual("ErrorBars!$A$2:$A$15", serie.ErrorBars.Minus.ValuesSource);

                    Assert.AreEqual(eErrorBarType.Both, serie.ErrorBarsX.BarType);
                    Assert.AreEqual(eErrorValueType.Custom, serie.ErrorBarsX.ValueType);
                    Assert.AreEqual(false, serie.ErrorBarsX.NoEndCap);

                    Assert.AreEqual("{2}", serie.ErrorBarsX.Plus.ValuesSource);
                    Assert.AreEqual("ErrorBars!$A$2:$A$16", serie.ErrorBarsX.Minus.ValuesSource);
                }
            }
        }
        [TestMethod]
        public void ErrorBars_Area()
        {
            var ws = _pck.Workbook.Worksheets.Add("ErrorBar_Area");
            LoadTestdata(ws);

            var chart = ws.Drawings.AddAreaChart("AreaChart1", eAreaChartType.Area);
            var serie = chart.Series.Add("D2:D100", "A2:A100");
            serie.AddErrorBars(eErrorBarType.Both, eErrorValueType.Custom);
            serie.ErrorBars.Plus.ValuesSource = "{1}";
            serie.ErrorBars.Minus.FormatCode = "General";
            serie.ErrorBars.Minus.ValuesSource = "=ErrorBar_Area!$A$2:$A$15";
            serie.ErrorBars.Minus.FormatCode = "0";

            serie.ErrorBarsX.Plus.ValuesSource = "{2}";
            serie.ErrorBarsX.Minus.FormatCode = "General";
            serie.ErrorBarsX.Minus.ValuesSource = "=ErrorBar_Area!$A$2:$A$16";
            serie.ErrorBarsX.Minus.FormatCode = "0";

            chart.SetPosition(1, 0, 5, 0);

            Assert.IsNotNull(serie.ErrorBars);
            Assert.IsNotNull(serie.ErrorBarsX);
            Assert.AreEqual(eErrorBarType.Both, serie.ErrorBars.BarType);
            Assert.AreEqual(eErrorValueType.Custom, serie.ErrorBars.ValueType);
            Assert.AreEqual(false, serie.ErrorBars.NoEndCap);

            Assert.AreEqual("{1}", serie.ErrorBars.Plus.ValuesSource);
            Assert.AreEqual("ErrorBar_Area!$A$2:$A$15", serie.ErrorBars.Minus.ValuesSource);

            Assert.AreEqual(eErrorBarType.Both, serie.ErrorBarsX.BarType);
            Assert.AreEqual(eErrorValueType.Custom, serie.ErrorBarsX.ValueType);
            Assert.AreEqual(false, serie.ErrorBarsX.NoEndCap);

            Assert.AreEqual("{2}", serie.ErrorBarsX.Plus.ValuesSource);
            Assert.AreEqual("ErrorBar_Area!$A$2:$A$16", serie.ErrorBarsX.Minus.ValuesSource);
        }
        [TestMethod]
        public void ErrorBars_ReadArea()
        {
            using (var p1 = new ExcelPackage())
            {
                var ws = p1.Workbook.Worksheets.Add("ErrorBars");
                LoadTestdata(ws);

                var chart = ws.Drawings.AddAreaChart("ScatterChart1", eAreaChartType.Area);
                var serie = chart.Series.Add("D2:D100", "A2:A100");
                serie.AddErrorBars(eErrorBarType.Both, eErrorValueType.Custom);
                serie.ErrorBars.Plus.ValuesSource = "{1}";
                serie.ErrorBars.Minus.FormatCode = "General";
                serie.ErrorBars.Minus.ValuesSource = "=ErrorBars!$A$2:$A$15";
                serie.ErrorBars.Minus.FormatCode = "0";

                serie.ErrorBarsX.Plus.ValuesSource = "{2}";
                serie.ErrorBarsX.Minus.FormatCode = "General";
                serie.ErrorBarsX.Minus.ValuesSource = "=ErrorBars!$A$2:$A$16";
                serie.ErrorBarsX.Minus.FormatCode = "0";

                chart.SetPosition(1, 0, 5, 0);
                p1.Save();
                using (var p2 = new ExcelPackage(p1.Stream))
                {
                    ws = p2.Workbook.Worksheets["ErrorBars"];

                    chart = ws.Drawings[0].As.Chart.AreaChart;
                    serie = chart.Series[0];

                    Assert.IsNotNull(serie.ErrorBars);
                    Assert.IsNotNull(serie.ErrorBarsX);
                    Assert.AreEqual(eErrorBarType.Both, serie.ErrorBars.BarType);
                    Assert.AreEqual(eErrorValueType.Custom, serie.ErrorBars.ValueType);
                    Assert.AreEqual(false, serie.ErrorBars.NoEndCap);

                    Assert.AreEqual("{1}", serie.ErrorBars.Plus.ValuesSource);
                    Assert.AreEqual("ErrorBars!$A$2:$A$15", serie.ErrorBars.Minus.ValuesSource);

                    Assert.AreEqual(eErrorBarType.Both, serie.ErrorBarsX.BarType);
                    Assert.AreEqual(eErrorValueType.Custom, serie.ErrorBarsX.ValueType);
                    Assert.AreEqual(false, serie.ErrorBarsX.NoEndCap);

                    Assert.AreEqual("{2}", serie.ErrorBarsX.Plus.ValuesSource);
                    Assert.AreEqual("ErrorBars!$A$2:$A$16", serie.ErrorBarsX.Minus.ValuesSource);
                }
            }
        }
        [TestMethod]
        public void ErrorBars_Delete()
        {
            var ws = _pck.Workbook.Worksheets.Add("ErrorBar_Percentage_removed");
            LoadTestdata(ws);

            var chart = ws.Drawings.AddLineChart("LineChart1_DeletedErrorbars", eLineChartType.Line);
            var serie = chart.Series.Add("D2:D100", "A2:A100");
            serie.AddErrorBars(eErrorBarType.Plus, eErrorValueType.Percentage);
            serie.ErrorBars.Value = 5;
            chart.SetPosition(1, 0, 5, 0);

            Assert.AreEqual(eErrorBarType.Plus, serie.ErrorBars.BarType);
            Assert.AreEqual(eErrorValueType.Percentage, serie.ErrorBars.ValueType);
            Assert.AreEqual(eErrorBarDirection.Y, serie.ErrorBars.Direction);
            Assert.AreEqual(false, serie.ErrorBars.NoEndCap);

            serie.ErrorBars.Remove();
            Assert.IsNull(serie.ErrorBars);
        }
        [TestMethod]
        public void ErrorBarsScatter_Delete()
        {
            var ws = _pck.Workbook.Worksheets.Add("ErrorBar_Scatter_removed");
            LoadTestdata(ws);

            var chart = ws.Drawings.AddScatterChart("LineChart1_DeletedErrorbars", eScatterChartType.XYScatter);
            var serie = chart.Series.Add("D2:D100", "A2:A100");
            serie.AddErrorBars(eErrorBarType.Plus, eErrorValueType.Percentage);
            Assert.IsNotNull(serie.ErrorBars);
            Assert.IsNotNull(serie.ErrorBarsX);

            serie.ErrorBars.Remove();
            Assert.IsNull(serie.ErrorBars);
            Assert.IsNotNull(serie.ErrorBarsX);

            serie.ErrorBarsX.Remove();
            Assert.IsNull(serie.ErrorBars);
            Assert.IsNull(serie.ErrorBarsX);

        }

    }
}

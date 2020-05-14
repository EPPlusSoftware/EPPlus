using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.Drawing.Chart.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Drawing.Chart
{
    [TestClass]
    public class ChartExTests : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("ChartEx.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void ReadChartEx()
        {
            using (var p = OpenTemplatePackage("Chartex.xlsx"))
            {
                var chart1 = (ExcelChartEx)p.Workbook.Worksheets[0].Drawings[0];
                var chart2 = (ExcelChartEx)p.Workbook.Worksheets[0].Drawings[1];
                var chart3 = (ExcelChartEx)p.Workbook.Worksheets[0].Drawings[2];

                Assert.IsNotNull(chart1.Fill);
                Assert.IsNotNull(chart1.PlotArea);
                Assert.IsNotNull(chart1.Legend);
                Assert.IsNotNull(chart1.Title);
                Assert.IsNotNull(chart1.Title.Font);

                Assert.IsInstanceOfType(chart1.Series[0].DataDimensions[0], typeof(ExcelChartExStringData));
                Assert.AreEqual(eStringDataType.Category, ((ExcelChartExStringData)chart1.Series[0].DataDimensions[0]).Type);
                Assert.AreEqual("_xlchart.v1.0", chart1.Series[0].DataDimensions[0].Formula);
                Assert.IsInstanceOfType(chart1.Series[0].DataDimensions[1], typeof(ExcelChartExNumericData));
                Assert.AreEqual(eNumericDataType.Value, ((ExcelChartExNumericData)chart1.Series[0].DataDimensions[1]).Type);
                Assert.AreEqual("_xlchart.v1.2", chart1.Series[0].DataDimensions[1].Formula);

                Assert.IsInstanceOfType(chart1.Series[1].DataDimensions[0], typeof(ExcelChartExStringData));
                Assert.AreEqual("_xlchart.v1.0", chart1.Series[1].DataDimensions[0].Formula);
                Assert.IsInstanceOfType(chart1.Series[1].DataDimensions[1], typeof(ExcelChartExNumericData));
                Assert.AreEqual("_xlchart.v1.4", chart1.Series[1].DataDimensions[1].Formula);

            }
        }
        [TestMethod]
        public void AddSunburstChart()
        {
            var ws = _pck.Workbook.Worksheets.Add("Sunburst");
            AddHierarkiData(ws);
            var chart = ws.Drawings.AddExtendedChart("Sunburst1", eChartExType.Sunburst);
            var serie = chart.Series.Add("Sunburst!$A$2:$C$17", "Sunburst!$D$2:$D$17");
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);
            serie.DataLabel.Position = eLabelPosition.Center;
            serie.DataLabel.ShowCategory = true;
            serie.DataLabel.ShowValue=true;
            chart.StyleManager.SetChartStyle(ePresetChartStyle.Sunburst7);
        }
        [TestMethod]
        public void AddTreemapChart()
        {
            var ws = _pck.Workbook.Worksheets.Add("Treemap");
            AddHierarkiData(ws);
            var chart = ws.Drawings.AddExtendedChart("Treemap", eChartExType.Treemap);
            var serie = chart.Series.Add("Treemap!$A$2:$C$17", "Treemap!$D$2:$D$17");
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);
            serie.DataLabel.Position = eLabelPosition.Center;
            serie.DataLabel.ShowCategory = true;
            serie.DataLabel.ShowValue = true;
            serie.DataLabel.ShowSeriesName = true;
            chart.StyleManager.SetChartStyle(ePresetChartStyle.Treemap9);
        }
        [TestMethod]
        public void AddBoxWhiskerChart()
        {
            var ws = _pck.Workbook.Worksheets.Add("BoxWhisker");    
            AddHierarkiData(ws);
            var chart = ws.Drawings.AddExtendedChart("BoxWhisker", eChartExType.BoxWhisker);
            var serie = chart.Series.Add("BoxWhisker!$A$2:$C$17", "BoxWhisker!$D$2:$D$17");
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);
            chart.StyleManager.SetChartStyle(ePresetChartStyle.BoxWhiskerStyle3);
        }
        [TestMethod]
        public void AddHistogramChart()
        {
            var ws = _pck.Workbook.Worksheets.Add("Histogram");
            AddHierarkiData(ws);
            var chart = ws.Drawings.AddExtendedChart("Histogram", eChartExType.Histogram);
            var serie = chart.Series.Add("Histogram!$A$2:$C$17", "Histogram!$D$2:$D$17");
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);
            chart.StyleManager.SetChartStyle(ePresetChartStyle.HistogramStyle2);
        }
        //[TestMethod]
        //public void AddParetoChart()
        //{
        //    var ws = _pck.Workbook.Worksheets.Add("Pareto");
        //    AddHierarkiData(ws);
        //    var chart = ws.Drawings.AddExtendedChart("Pareto", eChartExType.Pareto);
        //    var serie = chart.Series.Add("Pareto!$A$2:$C$17", "Pareto!$D$2:$D$17");
        //    chart.SetPosition(2, 0, 15, 0);
        //    chart.SetSize(1600, 900);
        //    chart.StyleManager.SetChartStyle(ePresetChartStyle.HistogramStyle4);
        //}
        [TestMethod]
        public void AddWaterfallChart()
        {
            var ws = _pck.Workbook.Worksheets.Add("Waterfall");
            AddHierarkiData(ws);
            var chart = ws.Drawings.AddExtendedChart("Waterfall", eChartExType.Waterfall);
            var serie = chart.Series.Add("Waterfall!$A$2:$C$17", "Waterfall!$D$2:$D$17");
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);
            var dt = chart.Series[0].DataPoints.Add(15);
            dt.SubTotal = true;
            dt = chart.Series[0].DataPoints.Add(0);
            dt.SubTotal = true;            
            dt=chart.Series[0].DataPoints.Add(4);
            dt.Fill.Style = eFillStyle.SolidFill;
            dt.Fill.SolidFill.Color.SetSchemeColor(eSchemeColor.Accent2);
            dt = chart.Series[0].DataPoints.Add(2);
            dt.Fill.Style = eFillStyle.SolidFill;
            dt.Fill.SolidFill.Color.SetSchemeColor(eSchemeColor.Accent4);

            dt=chart.Series[0].DataPoints[0];
            dt.Border.Fill.Style = eFillStyle.GradientFill;
            dt.Border.Fill.GradientFill.Colors.AddRgb(0, Color.Green);
            dt.Border.Fill.GradientFill.Colors.AddRgb(40, Color.Blue);
            dt.Border.Fill.GradientFill.Colors.AddRgb(70, Color.Red);
            dt.Fill.Style = eFillStyle.SolidFill;
            dt.Fill.SolidFill.Color.SetSchemeColor(eSchemeColor.Accent1);

            chart.StyleManager.SetChartStyle(ePresetChartStyle.HistogramStyle4);
        }
        [TestMethod]
        public void AddFunnelChart()
        {
            var ws = _pck.Workbook.Worksheets.Add("Funnel");
            AddHierarkiData(ws);
            var chart = ws.Drawings.AddExtendedChart("Funnel", eChartExType.Funnel);
            var serie = chart.Series.Add("Funnel!$A$2:$C$17", "Funnel!$D$2:$D$17");
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);
        }
        private class SalesData
        {
            public string Continent { get; set; }
            public string Country { get; set; }
            public string State { get; set; }
            public double Sales { get; set; }

        }
        private void AddHierarkiData(ExcelWorksheet ws)
        {

            var l = new List<SalesData>
            {
                new SalesData{ Continent="Europe", Country="Sweden", State = "Stockholm", Sales = 154 },
                new SalesData{ Continent="Asia", Country="Vietnam", State = "Ho Chi Minh", Sales= 88 },
                new SalesData{ Continent="Europe", Country="Sweden", State = "Västerås", Sales = 33 },
                new SalesData{ Continent="Asia", Country="Japan", State = "Tokyo", Sales= 534 },
                new SalesData{ Continent="Europe", Country="Germany", State = "Frankfurt", Sales = 109 },
                new SalesData{ Continent="Asia", Country="Vietnam", State = "Hanoi", Sales= 322 },
                new SalesData{ Continent="Asia", Country="Japan", State = "Osaka", Sales= 88 },
                new SalesData{ Continent="North America", Country="Canada", State = "Vancover", Sales= 99 },
                new SalesData{ Continent="Asia", Country="China", State = "Peking", Sales= 205 },
                new SalesData{ Continent="North America", Country="Canada", State = "Toronto", Sales= 138 },
                new SalesData{ Continent="Europe", Country="France", State = "Lyon", Sales = 185 },
                new SalesData{ Continent="North America", Country="USA", State = "Boston", Sales= 155 },
                new SalesData{ Continent="Europe", Country="France", State = "Paris", Sales = 127 },
                new SalesData{ Continent="North America", Country="USA", State = "New York", Sales= 330 },
                new SalesData{ Continent="Europe", Country="Germany", State = "Berlin", Sales = 210 },
                new SalesData{ Continent="North America", Country="USA", State = "San Fransico", Sales= 411 },
            };

            ws.Cells["A1"].LoadFromCollection(l, true, OfficeOpenXml.Table.TableStyles.Medium12);
        }
    }
}

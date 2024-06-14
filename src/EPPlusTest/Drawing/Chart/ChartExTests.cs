using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.Drawing.Chart.Style;
using System.Drawing;
using System.Globalization;
using System.IO;

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
            var dirName = _pck.File.DirectoryName;
            var fileName = _pck.File.FullName;
            SaveAndCleanup(_pck);
            if (File.Exists(fileName))
            {
                File.Copy(fileName, dirName + "\\ChartExRead.xlsx", true);
            }
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
            LoadHierarkiTestData(ws);
            var chart = ws.Drawings.AddExtendedChart("Sunburst1", eChartExType.Sunburst);
            var serie = chart.Series.Add("Sunburst!$D$2:$D$17", "Sunburst!$A$2:$C$17");
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);
            serie.DataLabel.Position = eLabelPosition.Center;   
            serie.DataLabel.ShowCategory = true;
            serie.DataLabel.ShowValue=true;
            var dp=serie.DataPoints.Add(2);
            dp.Fill.Style = eFillStyle.PatternFill;
            dp.Fill.PatternFill.PatternType = eFillPatternStyle.DashDnDiag;
            dp.Fill.PatternFill.BackgroundColor.SetRgbColor(Color.Red);
            dp.Fill.PatternFill.ForegroundColor.SetRgbColor(Color.DarkGray);
            chart.StyleManager.SetChartStyle(ePresetChartStyle.SunburstChartStyle7);

            Assert.AreEqual(eDrawingType.Chart, chart.DrawingType);
            Assert.IsInstanceOfType(chart, typeof(ExcelSunburstChart));
            Assert.AreEqual(0, chart.Axis.Length);
            Assert.IsNull(chart.XAxis);
            Assert.IsNull(chart.YAxis);            
        }
        [TestMethod]
        public void AddSunburstChartSheet()
        {
            var ws = _pck.Workbook.Worksheets.AddChart("SunburstSheet", eChartType.Sunburst);
            var chart = ws.Chart.As.Chart.SunburstChart;
            var serie = chart.Series.Add("Sunburst!$D$2:$D$17", "Sunburst!$A$2:$C$17");
            serie.DataLabel.Position = eLabelPosition.Center;
            serie.DataLabel.ShowCategory = true;
            serie.DataLabel.ShowValue = true;
            var dp = serie.DataPoints.Add(2);
            dp.Fill.Style = eFillStyle.PatternFill;
            dp.Fill.PatternFill.PatternType = eFillPatternStyle.DashDnDiag;
            dp.Fill.PatternFill.BackgroundColor.SetRgbColor(Color.Red);
            dp.Fill.PatternFill.ForegroundColor.SetRgbColor(Color.DarkGray);
            chart.StyleManager.SetChartStyle(ePresetChartStyle.SunburstChartStyle7);

            Assert.IsInstanceOfType(chart, typeof(ExcelSunburstChart));
            Assert.AreEqual(0, chart.Axis.Length);
            Assert.IsNull(chart.XAxis);
            Assert.IsNull(chart.YAxis);

        }

        [TestMethod]
        public void AddTreemapChart()
        {
            var ws = _pck.Workbook.Worksheets.Add("Treemap");
            LoadHierarkiTestData(ws);
            var chart = ws.Drawings.AddExtendedChart("Treemap", eChartExType.Treemap);
            var serie = chart.Series.Add("Treemap!$D$2:$D$17", "Treemap!$A$2:$C$17");
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);
            serie.DataLabel.Position = eLabelPosition.Center;
            serie.DataLabel.ShowCategory = true;
            serie.DataLabel.ShowValue = true;
            serie.DataLabel.ShowSeriesName = true;
            chart.StyleManager.SetChartStyle(ePresetChartStyle.TreemapChartStyle9);
            Assert.IsInstanceOfType(chart, typeof(ExcelTreemapChart));
        }
        [TestMethod]
        public void AddBoxWhiskerChart()
        {
            var ws = _pck.Workbook.Worksheets.Add("BoxWhisker");    
            LoadHierarkiTestData(ws);
            var chart = ws.Drawings.AddBoxWhiskerChart("BoxWhisker");
            var serie = chart.Series.Add("BoxWhisker!$D$2:$D$17", "BoxWhisker!$A$2:$C$17");
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);
            chart.StyleManager.SetChartStyle(ePresetChartStyle.BoxWhiskerChartStyle3);

            Assert.IsInstanceOfType(chart, typeof(ExcelBoxWhiskerChart));
            Assert.AreEqual(2, chart.Axis.Length);
            Assert.IsNotNull(chart.XAxis);
            Assert.IsNotNull(chart.YAxis);

            Assert.IsFalse(serie.ShowMeanLine);
            Assert.IsTrue(serie.ShowMeanMarker);
            Assert.IsTrue(serie.ShowOutliers);
            Assert.IsFalse(serie.ShowNonOutliers);

            Assert.AreEqual(eQuartileMethod.Exclusive, serie.QuartileMethod);
        }
        [TestMethod]
        public void AddHistogramChart()
        {
            var ws = _pck.Workbook.Worksheets.Add("Histogram");
            LoadHierarkiTestData(ws);
            var chart = ws.Drawings.AddHistogramChart("Histogram");
            var serie = chart.Series.Add("Histogram!$D$2:$D$17", "Histogram!$A$2:$C$17");
            serie.Binning.Underflow = 1;
            serie.Binning.OverflowAutomatic = true;
            serie.Binning.Count = 3;
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);
            chart.StyleManager.SetChartStyle(ePresetChartStyle.HistogramChartStyle2);

            Assert.IsInstanceOfType(chart, typeof(ExcelHistogramChart));
        }
        [TestMethod]
        public void AddParetoChart()
        {
            var ws = _pck.Workbook.Worksheets.Add("Pareto");
            LoadHierarkiTestData(ws);
            var chart = ws.Drawings.AddHistogramChart("Pareto", true);
            var serie = chart.Series.Add("Pareto!$D$2:$D$17", "Pareto!$A$2:$C$17");
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);

            Assert.IsInstanceOfType(chart, typeof(ExcelHistogramChart));
            Assert.IsNotNull(serie.ParetoLine);
            serie.ParetoLine.Fill.Style = eFillStyle.SolidFill;
            serie.ParetoLine.Fill.SolidFill.Color.SetRgbColor(Color.FromArgb(128,255,0,0),true);
            serie.ParetoLine.Effect.SetPresetShadow(ePresetExcelShadowType.OuterBottomRight);
            Assert.AreEqual(eChartType.Pareto, chart.ChartType);
            chart.StyleManager.SetChartStyle(ePresetChartStyle.HistogramChartStyle4);
        }
        [TestMethod]
        public void AddWaterfallChart()
        {
            var ws = _pck.Workbook.Worksheets.Add("Waterfall");
            LoadHierarkiTestData(ws);
            var chart = ws.Drawings.AddWaterfallChart("Waterfall");
            var serie = chart.Series.Add("Waterfall!$D$2:$D$17", "Waterfall!$A$2:$C$17");
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);
            var dt = chart.Series[0].DataPoints.Add(15);
            dt.SubTotal = true;
            dt = serie.DataPoints.Add(0);
            dt.SubTotal = true;            
            dt = serie.DataPoints.Add(4);
            dt.Fill.Style = eFillStyle.SolidFill;
            dt.Fill.SolidFill.Color.SetSchemeColor(eSchemeColor.Accent2);
            dt = serie.DataPoints.Add(2);
            dt.Fill.Style = eFillStyle.SolidFill;
            dt.Fill.SolidFill.Color.SetSchemeColor(eSchemeColor.Accent4);

            dt = serie.DataPoints[0];
            dt.Border.Fill.Style = eFillStyle.GradientFill;
            dt.Border.Fill.GradientFill.Colors.AddRgb(0, Color.Green);
            dt.Border.Fill.GradientFill.Colors.AddRgb(40, Color.Blue);
            dt.Border.Fill.GradientFill.Colors.AddRgb(70, Color.Red);
            dt.Fill.Style = eFillStyle.SolidFill;
            dt.Fill.SolidFill.Color.SetSchemeColor(eSchemeColor.Accent1);

            chart.StyleManager.SetChartStyle(ePresetChartStyle.WaterfallChartStyle4);

            Assert.IsInstanceOfType(chart, typeof(ExcelWaterfallChart));
            Assert.AreEqual(4,serie.DataPoints.Count);
            Assert.IsTrue(serie.DataPoints[0].SubTotal);
            Assert.AreEqual(eFillStyle.GradientFill, serie.DataPoints[0].Border.Fill.Style);
            Assert.AreEqual(3, serie.DataPoints[0].Border.Fill.GradientFill.Colors.Count);
            Assert.AreEqual(eFillStyle.SolidFill, serie.DataPoints[0].Fill.Style);
            Assert.AreEqual(eSchemeColor.Accent1, serie.DataPoints[0].Fill.SolidFill.Color.SchemeColor.Color);
            Assert.IsTrue(serie.DataPoints[15].SubTotal);
        }
        [TestMethod]
        public void AddFunnelChart()
        {
            var ws = _pck.Workbook.Worksheets.Add("Funnel");
            LoadHierarkiTestData(ws);
            var chart = ws.Drawings.AddFunnelChart("Funnel");
            var serie = chart.Series.Add("Funnel!$D$2:$D$17", "Funnel!$A$2:$C$17");
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);
        }
        [TestMethod]
        public void AddRegionMapChart()
        {
            var ws = _pck.Workbook.Worksheets.Add("RegionMap");
            LoadGeoTestData(ws);
            var chart = ws.Drawings.AddRegionMapChart("RegionMap");
            var serie = chart.Series.Add("RegionMap!$C$2:$C$11", "RegionMap!$A$2:$B$11");
            serie.HeaderAddress = ws.Cells["$A$1"];
            serie.DataDimensions[0].NameFormula = "$A$1:$B$1";
            serie.DataDimensions[1].NameFormula = "$C$1";
            serie.ColorBy = eColorBy.CategoryNames;
            serie.Region = new CultureInfo("sv");
            serie.Language = new CultureInfo("sv-SE");
            serie.Colors.NumberOfColors = eNumberOfColors.ThreeColor;
            serie.Colors.MinColor.Color.SetSchemeColor(eSchemeColor.Dark1);
            serie.Colors.MinColor.ValueType = eColorValuePositionType.Number;
            serie.Colors.MinColor.PositionValue = 22;
            serie.Colors.MidColor.ValueType = eColorValuePositionType.Percent;
            serie.Colors.MidColor.PositionValue = 50.11;
            serie.Colors.MaxColor.ValueType = eColorValuePositionType.Extreme;
            serie.Colors.MaxColor.Color.SetRgbColor(Color.Red);
            serie.DataLabel.Border.Width = 1;
            serie.ViewedRegionType = eGeoMappingLevel.DataOnly;
            serie.ProjectionType = eProjectionType.Miller;

            chart.Legend.Add();
            chart.Legend.Position = eLegendPosition.Left;
            chart.Legend.PositionAlignment = ePositionAlign.Center;
            chart.Title.Text = "Sweden Region Map";
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);

            Assert.AreEqual("RegionMap!$A$2:$B$11", serie.XSeries);
            Assert.AreEqual("RegionMap!$C$2:$C$11", serie.Series);

            Assert.AreEqual("sv", serie.Region.TwoLetterISOLanguageName);
            Assert.AreEqual("sv-SE", serie.Language.Name);
        }
        [TestMethod]
        public void CopyBoxWhisker()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var package = new ExcelPackage();
            var package2 = new ExcelPackage();
            var worksheet1 = package2.Workbook.Worksheets.Add("Test_BoxWhiskers");
            ExcelChart chart3 = worksheet1.Drawings.AddBoxWhiskerChart("Status");
            var bwSerie1 = chart3.Series.Add(worksheet1.Cells[1, 1, 2, 1], null);
            chart3.SetPosition(10, 10);
            chart3.SetSize(750, 470);
            chart3.Title.Text = "Test BoxWhiskers";
            chart3.XAxis.Deleted = true;
            chart3.YAxis.AddTitle("Test");            
            chart3.Legend.Position   = eLegendPosition.TopRight;
            chart3.StyleManager.SetChartStyle(ePresetChartStyleMultiSeries.BoxWhiskerChartStyle6); //BoxWhiskerChartStyle3);
            
            var ws=package.Workbook.Worksheets.Add("Test_BoxWhiskers_New", worksheet1);
            var chart = ws.Drawings[0].As.Chart.BoxWhiskerChart;
            Assert.IsTrue(string.IsNullOrEmpty(chart.Series[0].XSeries));
        }
    }

}

using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.Drawing.Chart.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Drawing.Chart
{
    [TestClass]
    public class ChartExReadTests : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("ChartExRead.xlsx");
        }
        [ClassCleanup]
        public static void Cleanup()
        {
        }
        [TestMethod]
        public void ReadSunburstChart()
        {            
            var ws = GetWorksheet("Sunburst");
            Assert.AreEqual(1,ws.Drawings.Count);
            var chart = ws.Drawings[0].As.SunburstChart();
            Assert.AreEqual(1, chart.Series.Count);
            var serie = chart.Series[0];
            Assert.AreEqual("Sunburst!$A$2:$C$17", serie.Series);
            Assert.AreEqual("Sunburst!$D$2:$D$17", serie.XSeries);
            Assert.IsNotNull(serie.DataLabel);
            Assert.AreEqual(eLabelPosition.Center, serie.DataLabel.Position);
            Assert.IsTrue(serie.DataLabel.ShowCategory);
            Assert.IsTrue(serie.DataLabel.ShowValue);
            Assert.IsFalse(serie.DataLabel.ShowSeriesName);
            Assert.AreEqual(1, serie.DataPoints.Count);
            var dp=serie.DataPoints[2];

            Assert.AreEqual(eFillStyle.PatternFill, dp.Fill.Style);
            Assert.AreEqual(eFillPatternStyle.DashDnDiag, dp.Fill.PatternFill.PatternType);
            Assert.AreEqual(Color.Red.ToArgb(), dp.Fill.PatternFill.BackgroundColor.RgbColor.Color.ToArgb());
            Assert.AreEqual(Color.DarkGray.ToArgb(), dp.Fill.PatternFill.ForegroundColor.RgbColor.Color.ToArgb());
            Assert.AreEqual(((int)ePresetChartStyle.SunburstChartStyle7), chart.StyleManager.Style.Id);

            Assert.IsInstanceOfType(chart, typeof(ExcelSunburstChart));
            Assert.AreEqual(0, chart.Axis.Length);
            Assert.IsNull(chart.XAxis);
            Assert.IsNull(chart.YAxis);            
        }
        [TestMethod]
        public void ReadTreemapChart()
        {
            var ws = GetWorksheet("Treemap");
            Assert.AreEqual(1, ws.Drawings.Count);
            var chart = ws.Drawings[0].As.TreemapChart();
            Assert.AreEqual(1, chart.Series.Count);
            var serie = chart.Series[0];
            Assert.AreEqual("Treemap!$A$2:$C$17", serie.Series);
            Assert.AreEqual("Treemap!$D$2:$D$17", serie.XSeries);
            Assert.IsNotNull(serie.DataLabel);
            Assert.AreEqual(eLabelPosition.Center, serie.DataLabel.Position);
            Assert.IsTrue(serie.DataLabel.ShowCategory);
            Assert.IsTrue(serie.DataLabel.ShowValue);
            Assert.IsTrue(serie.DataLabel.ShowSeriesName);
            Assert.AreEqual(((int)ePresetChartStyle.TreemapChartStyle9), chart.StyleManager.Style.Id);
            chart.StyleManager.SetChartStyle(ePresetChartStyle.TreemapChartStyle9);
            Assert.IsInstanceOfType(chart, typeof(ExcelTreemapChart));
        }
        [TestMethod]
        public void ReadBoxWhiskerChart()
        {
            var ws = GetWorksheet("BoxWhisker"); 
            Assert.AreEqual(1, ws.Drawings.Count);
            var chart = ws.Drawings[0].As.BoxWhiskerChart();
            Assert.AreEqual(1, chart.Series.Count);
            var serie = chart.Series[0];
            Assert.AreEqual("BoxWhisker!$A$2:$C$17", serie.Series);
            Assert.AreEqual("BoxWhisker!$D$2:$D$17", serie.XSeries);

            Assert.AreEqual(((int)ePresetChartStyle.BoxWhiskerChartStyle3), chart.StyleManager.Style.Id);

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
        public void ReadHistogramChart()
        {
            var ws = GetWorksheet("Histogram");
            Assert.AreEqual(1, ws.Drawings.Count);
            var chart = ws.Drawings[0].As.HistogramChart();
            Assert.AreEqual(1, chart.Series.Count);
            var serie = chart.Series[0];
            Assert.AreEqual("Histogram!$A$2:$C$17", serie.Series);
            Assert.AreEqual("Histogram!$D$2:$D$17", serie.XSeries);
            Assert.AreEqual(1, serie.Binning.Underflow);
            Assert.IsTrue(serie.Binning.OverflowAutomatic);
            Assert.AreEqual(3, serie.Binning.Count);
            Assert.AreEqual(((int)ePresetChartStyle.HistogramChartStyle2), chart.StyleManager.Style.Id);

            Assert.IsInstanceOfType(chart, typeof(ExcelHistogramChart));
        }
        [TestMethod]
        public void ReadParetoChart()
        {
            var ws = GetWorksheet("Pareto");
            Assert.AreEqual(1, ws.Drawings.Count);
            var chart = ws.Drawings[0].As.Type<ExcelHistogramChart>();
            Assert.AreEqual(1, chart.Series.Count);
            var serie = chart.Series[0];
            Assert.AreEqual("Pareto!$A$2:$C$17", serie.Series);
            Assert.AreEqual("Pareto!$D$2:$D$17", serie.XSeries);

            Assert.IsInstanceOfType(chart, typeof(ExcelHistogramChart));
            Assert.IsNotNull(serie.ParetoLine);
            Assert.AreEqual(eFillStyle.SolidFill, serie.ParetoLine.Fill.Style);
            Assert.AreEqual(Color.Red.ToArgb(), serie.ParetoLine.Fill.SolidFill.Color.RgbColor.Color.ToArgb());
            Assert.AreEqual(OfficeOpenXml.Drawing.Style.Coloring.eColorTransformType.Alpha, serie.ParetoLine.Fill.SolidFill.Color.Transforms[0].Type);
            Assert.AreEqual(50.39, serie.ParetoLine.Fill.SolidFill.Color.Transforms[0].Value);

            Assert.AreEqual(45, serie.ParetoLine.Effect.OuterShadow.Direction);
            Assert.AreEqual(eRectangleAlignment.TopLeft, serie.ParetoLine.Effect.OuterShadow.Alignment);
            Assert.AreEqual(eChartType.Pareto, chart.ChartType);
            Assert.AreEqual(((int)ePresetChartStyle.HistogramChartStyle4), chart.StyleManager.Style.Id);
        }
        [TestMethod]
        public void ReadWaterfallChart()
        {
            var ws = GetWorksheet("Waterfall");

            Assert.AreEqual(1, ws.Drawings.Count);
            var chart = ws.Drawings[0].As.WaterfallChart();
            Assert.AreEqual(1, chart.Series.Count);
            var serie = chart.Series[0];
            Assert.AreEqual("Waterfall!$A$2:$C$17", serie.Series);
            Assert.AreEqual("Waterfall!$D$2:$D$17", serie.XSeries);

            Assert.AreEqual(4, serie.DataPoints.Count);
            var dt = serie.DataPoints[15];
            Assert.IsTrue(dt.SubTotal);
            dt = serie.DataPoints[0];
            Assert.AreEqual(eFillStyle.GradientFill, dt.Border.Fill.Style);
            Assert.AreEqual(3, dt.Border.Fill.GradientFill.Colors.Count);
            Assert.AreEqual(eFillStyle.SolidFill, dt.Fill.Style);
            Assert.AreEqual(eSchemeColor.Accent1, dt.Fill.SolidFill.Color.SchemeColor.Color);

            Assert.IsTrue(dt.SubTotal);
            dt = serie.DataPoints[4];
            Assert.IsFalse(dt.SubTotal);
            Assert.AreEqual(eFillStyle.SolidFill, dt.Fill.Style);
            Assert.AreEqual(eSchemeColor.Accent2, dt.Fill.SolidFill.Color.SchemeColor.Color);


            dt = serie.DataPoints[2];
            Assert.IsFalse(dt.SubTotal);
            Assert.AreEqual(eFillStyle.SolidFill, dt.Fill.Style);
            Assert.AreEqual(eSchemeColor.Accent4, dt.Fill.SolidFill.Color.SchemeColor.Color);

            Assert.AreEqual(((int)ePresetChartStyle.WaterfallChartStyle4), chart.StyleManager.Style.Id);

            Assert.IsInstanceOfType(chart, typeof(ExcelWaterfallChart));
        }
        [TestMethod]
        public void ReadFunnelChart()
        {
            var ws = GetWorksheet("Funnel");
            Assert.AreEqual(1, ws.Drawings.Count);
            var chart = ws.Drawings[0].As.FunnelChart();
            Assert.AreEqual(1, chart.Series.Count);
            var serie = chart.Series[0];
            Assert.AreEqual("Funnel!$A$2:$C$17", serie.Series);
            Assert.AreEqual("Funnel!$D$2:$D$17", serie.XSeries);
            
            Assert.AreEqual(((int)ePresetChartStyle.FunnelChartStyle1), chart.StyleManager.Style.Id);
            Assert.IsInstanceOfType(chart, typeof(ExcelFunnelChart));
        }
        [TestMethod]
        public void ReadRegionMapChart()
        {
            var ws = GetWorksheet("RegionMap");
            Assert.AreEqual(1, ws.Drawings.Count);
            var chart = ws.Drawings[0].As.RegionMapChart();
            Assert.AreEqual(1, chart.Series.Count);
            var serie = chart.Series[0];
            Assert.AreEqual("RegionMap!$A$2:$B$11", serie.XSeries);
            Assert.AreEqual("RegionMap!$C$2:$C$11", serie.Series);

            Assert.AreEqual("RegionMap!$A$1", serie.HeaderAddress.Address);
            Assert.AreEqual("RegionMap!$A$1:$B$1", serie.DataDimensions[1].NameFormula);
            Assert.AreEqual("RegionMap!$C$1", serie.DataDimensions[0].NameFormula);
            Assert.IsInstanceOfType(serie.DataDimensions[1], typeof(ExcelChartExStringData));
            Assert.AreEqual(eStringDataType.ColorString, ((ExcelChartExStringData)serie.DataDimensions[1]).Type);

            Assert.AreEqual(eNumberOfColors.ThreeColor, serie.Colors.NumberOfColors);
            Assert.AreEqual(eSchemeColor.Dark1,serie.Colors.MinColor.Color.SchemeColor.Color);
            Assert.AreEqual(eColorValuePositionType.Number, serie.Colors.MinColor.ValueType);
            Assert.AreEqual(22, serie.Colors.MinColor.PositionValue);
            Assert.AreEqual(eColorValuePositionType.Percent, serie.Colors.MidColor.ValueType);
            Assert.AreEqual(50.11, serie.Colors.MidColor.PositionValue);
            Assert.AreEqual(eColorValuePositionType.Extreme, serie.Colors.MaxColor.ValueType);
            Assert.AreEqual(Color.Red.ToArgb(), serie.Colors.MaxColor.Color.RgbColor.Color.ToArgb());

            Assert.AreEqual(1,serie.DataLabel.Border.Width);
            Assert.AreEqual(eGeoMappingLevel.DataOnly, serie.ViewedRegionType);
            Assert.AreEqual(eProjectionType.Miller, serie.ProjectionType);

            Assert.IsTrue(chart.HasLegend);
            Assert.AreEqual(eLegendPosition.Left, chart.Legend.Position);
            Assert.AreEqual(ePositionAlign.Center, chart.Legend.PositionAlignment);
            Assert.AreEqual("Sweden Region Map", chart.Title.Text);

            Assert.AreEqual("sv", serie.Region.TwoLetterISOLanguageName);
            Assert.AreEqual("sv-SE", serie.Language.Name);
        }
        private ExcelWorksheet GetWorksheet(string wsName)
        {
            if (_pck == null || _pck.Workbook.Worksheets.Count==0)
            {
                Assert.Inconclusive("ChartExRead.xlsx does not exist");
            }
            var ws = _pck.Workbook.Worksheets[wsName];
            if (ws == null)
            {
                Assert.Inconclusive($"Worksheet {wsName} does not exist");
            }
            return ws;
        }

    }
}


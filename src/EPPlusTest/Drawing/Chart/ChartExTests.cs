using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using System;
using System.Collections.Generic;
using System.Linq;
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
            //_pck = OpenPackage("ErrorBars.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            //SaveAndCleanup(_pck);
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
        public void AddChartEx()
        {
            using (var p = OpenPackage("Chartex.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Sunburst");
                LoadTestdata(ws);
                var chart = ws.Drawings.AddExtendedChart("Sunburst1", eChartExType.Sunburst);
                chart.Series.Add("A1:A5", "B1:B5");
                SaveAndCleanup(p);
            }
        }
    }
}

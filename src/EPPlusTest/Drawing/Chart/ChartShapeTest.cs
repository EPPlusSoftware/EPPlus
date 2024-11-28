using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Drawing.Chart
{
    [TestClass]
    public class ChartShapeTest : TestBase
    {
        [TestMethod]
        public void ShapeInChartTest()
        {
            using var p = OpenTemplatePackage("ShapeInChart.xlsx");
            var ws = p.Workbook.Worksheets[0];
            var chart = ws.Drawings[0] as ExcelChart;
            var cdr = chart.ChartDrawings[0];
        }
    }
}

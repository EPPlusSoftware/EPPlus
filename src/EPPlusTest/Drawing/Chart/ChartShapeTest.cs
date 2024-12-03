using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System.Drawing;

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
            //ws.Drawings.AddShape("myshape", OfficeOpenXml.Drawing.eShapeStyle.Diamond);
            var chart = ws.Drawings[0] as ExcelChart;
            var cdr = chart.ChartDrawings[0];
            var chartShape = chart.AddShape("MyShape", eShapeStyle.Plus);
            chartShape.Fill.Color = Color.Orange;
            chartShape.Effect.SetPresetShadow(ePresetExcelShadowType.OuterRight);
            chartShape.SetPosition(0.99, 0.99);
            p.SaveAs(@"c:\epplustest\testoutput\shapeInChartTest.xlsx");
        }
    }


    /*TODO
     * Pictures
     * Group Shapes
     * Properties
     * - SetPosition, Make percentage and to can't be greater than 1.
     */
}

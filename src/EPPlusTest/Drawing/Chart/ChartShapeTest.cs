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
            var cdr = chart.ChartDrawings[1];
            cdr.SetSize(200);
            var chartShape = chart.AddShape("MyShape", eShapeStyle.Diamond);
            chartShape.Fill.Color = Color.Orange;
            chartShape.Effect.SetPresetShadow(ePresetExcelShadowType.OuterRight);
            chartShape.SetPosition(0, 0);
            chartShape.SetSize(100, 200);
            var chartShape2 = chart.AddShape("MyShape2", eShapeStyle.Diamond);
            chartShape2.Fill.Color = Color.Orange;
            chartShape2.Effect.SetPresetShadow(ePresetExcelShadowType.OuterRight);
            chartShape2.SetPosition(10000, 10000);
            chartShape2.SetSize(30);
            p.SaveAs(@"c:\epplustest\testoutput\shapeInChartTest.xlsx");
        }
    }


    /*TODO
     * Pictures
     * Group Shapes
     * Properties
     */
}

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
            var dr = ws.Drawings.AddShape("myshape", OfficeOpenXml.Drawing.eShapeStyle.Diamond);
            dr.SetSize(500, 500);
            var chart = ws.Drawings[0] as ExcelChart;
            //var cdr = chart.ChartDrawings[1];
            //cdr.SetSize(200);
            //var cdr2 = chart.ChartDrawings[4];
            //cdr2.SetSize(200);

            //var chartShape = chart.AddShape("MyShape", eShapeStyle.Diamond);
            //chartShape.Fill.Color = Color.LightSeaGreen;
            //chartShape.Effect.SetPresetShadow(ePresetExcelShadowType.OuterRight);
            //chartShape.SetPosition(0, 0);
            //chartShape.SetSize(100, 200);
            //var chartShape2 = chart.AddShape("MyShape2", eShapeStyle.Diamond);
            //chartShape2.Fill.Color = Color.Orange;
            //chartShape2.Effect.SetPresetShadow(ePresetExcelShadowType.OuterRight);
            //chartShape2.SetPosition(10000, 10000);
            //chartShape2.SetSize(30);

            //var chartPic = chart.AddPicture("MyPic", @"C:\epplusTest\epplusobject.png");
            //chartPic.SetPosition(0, 5000);
            //chartPic.SetSize(200);

            //var chartPic2 = chart.AddPicture("MyPic2", @"C:\epplusTest\epplusobject.png");
            //chartPic2.SetPosition(0, 5000);

            var shp1 = chart.AddShape("level1", eShapeStyle.Star10);
            var shp2 = chart.AddShape("level2", eShapeStyle.UpDownArrow);
            var shp3 = chart.AddShape("level3", eShapeStyle.Teardrop);
            var group = shp1.Group(shp2);


            p.SaveAs(@"c:\epplustest\testoutput\shapeInChartTest.xlsx");
        }
    }


    /*TODO
     * Pictures - test Stream
     * Group Shapes - Test creating groupshape
     * Copy
     * -Copy whole drawings xml, set new name of drawings xml, create rel Id, done
     */
}

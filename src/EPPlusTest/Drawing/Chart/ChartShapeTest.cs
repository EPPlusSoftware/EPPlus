using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System.Drawing;
using System.IO;

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
            //var cdr = chart.ChartDrawings[1];
            //cdr.SetSize(200);
            //var cdr2 = chart.ChartDrawings[4];
            //cdr2.SetSize(200);

            var chartShape = chart.AddShape("MyShape", eShapeStyle.Diamond);
            chartShape.Fill.Color = Color.LightSeaGreen;
            chartShape.Effect.SetPresetShadow(ePresetExcelShadowType.OuterRight);
            chartShape.SetPosition(144, 240);
            chartShape.SetSize(240, 144);
            var chartShape2 = chart.AddShape("MyShape2", eShapeStyle.Diamond);
            chartShape2.Fill.Color = Color.Orange;
            chartShape2.Effect.SetPresetShadow(ePresetExcelShadowType.OuterRight);
            chartShape2.SetPosition(10000, 10000);
            chartShape2.SetSize(30);

            var chartPic = chart.AddPicture("MyPic", @"C:\epplusTest\epplusobject.png");
            chartPic.SetPosition(0, 5000);
            chartPic.SetSize(200);


            var myPic = Properties.Resources.GetOLEObjectFullFileName("SampleIcon.bmp");
            using (FileStream fileStream = new FileStream(myPic, FileMode.Open, FileAccess.Read))
            {
                var chartPic2 = chart.AddPicture("MyPic2", fileStream);
                chartPic2.SetPosition(0, 5000);
                chartPic2.SetSize(200);
            }

            var shp1 = chart.AddShape("level1", eShapeStyle.Star10);
            var shp2 = chart.AddShape("level2", eShapeStyle.UpDownArrow);
            shp2.SetPosition(0, 50);
            var shp3 = chart.AddShape("level3", eShapeStyle.Teardrop);
            shp3.SetPosition(50, 0);
            var group = shp1.Group(shp2);
            group.Group(shp3);
            group.SetPosition(50, 50);
            //var shap1 = ws.Drawings.AddShape("shap1", eShapeStyle.Gear6);
            //var shap2 = ws.Drawings.AddShape("shap2", eShapeStyle.LeftArrow);
            //shap1.Group(shap2);


            p.SaveAs(@"c:\epplustest\testoutput\shapeInChartTest.xlsx");
        }
    }


    /*TODO
     * Vi skapar ej en connection shape i epplus?
     * Copy
     * -Copy whole drawings xml, set new name of drawings xml, create rel Id, done
     */
}

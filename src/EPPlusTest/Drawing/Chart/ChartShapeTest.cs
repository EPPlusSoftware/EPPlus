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
            //var cdr = chart.Drawings[1];
            //cdr.SetSize(200);
            //var cdr2 = chart.Drawings[4];
            //cdr2.SetSize(200);

            //var chartShape = chart.AddShape("MyShape", eShapeStyle.Diamond);
            //chartShape.Fill.Color = Color.LightSeaGreen;
            //chartShape.Effect.SetPresetShadow(ePresetExcelShadowType.OuterRight);
            //chartShape.SetPosition(144, 240);
            ////chartShape.SetSize(240, 144);
            //var chartShape2 = chart.AddShape("MyShape2", eShapeStyle.Diamond);
            //chartShape2.Fill.Color = Color.Orange;
            //chartShape2.Effect.SetPresetShadow(ePresetExcelShadowType.OuterRight);
            //chartShape2.SetPosition(10000, 10000);
            //chartShape2.SetSize(30);

            //var chartPic = chart.AddPicture("MyPic", @"C:\epplusTest\epplusobject.png");
            //chartPic.SetPosition(0, 5000);
            //chartPic.SetSize(200);


            //var myPic = Properties.Resources.GetOLEObjectFullFileName("SampleIcon.bmp");
            //using (FileStream fileStream = new FileStream(myPic, FileMode.Open, FileAccess.Read))
            //{
            //    var chartPic2 = chart.AddPicture("MyPic2", fileStream);
            //    chartPic2.SetPosition(0, 5000);
            //    chartPic2.SetSize(200);
            //}

            var shp1 = chart.AddShape("level1", eShapeStyle.Star10);
            shp1.SetPosition(150, 250);
            var shp2 = chart.AddShape("level2", eShapeStyle.UpDownArrow);
            shp2.SetPosition(200, 200);
            var shp3 = chart.AddShape("level3", eShapeStyle.Teardrop);
            shp3.SetPosition(100, 50);
            var shp4 = chart.AddShape("level4", eShapeStyle.MathEqual);
            shp4.SetPosition(10, 60);
            var group1 = shp1.Group(shp2, shp3);
            var group2 = group1.Group(shp4);
            group2.SetPosition(0, 0);
            shp4.UnGroup();
            shp4.Copy(chart);

            var chart2 = ws.Drawings.AddChart("Chart 3", eChartType.Line);
            //chart2.Series.Add(ws.Cells["B2:B6"], ws.Cells["C2:C6"]);
            //chart2.SetSize(480, 288);
            //chart2.AddShape("hsp", eShapeStyle.Can);
            //shp4.Copy(chart2);
            //chartPic.Copy(chart);
            //chartPic.Copy(chart2);
            //group2.Copy(chart);
            //group2.Copy(chart2);

            chart.Copy(ws, 20, 0);



            //var d1 = ws.Drawings.AddShape("shape1", eShapeStyle.Rect);
            //var d2 = ws.Drawings.AddShape("shape2", eShapeStyle.QuadArrow);
            //var d3 = ws.Drawings.AddShape("shape3", eShapeStyle.Plus);
            //var f1 = d1.Group(d2);
            //var f2 = f1.Group(d3);
            //d3.UnGroup();

            p.SaveAs(@"c:\epplustest\testoutput\shapeInChartTest.xlsx");
        }
    }


    /*TODO
     * Resize bounding box for grouped objects.
     * Copy
     * -Copy whole drawings xml, set new name of drawings xml, create rel Id, done
     * Remove
     */
}

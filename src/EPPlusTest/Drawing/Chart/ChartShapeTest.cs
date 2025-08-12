using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System.IO;
using System.Drawing;
using EPPlusTest.Properties;

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
            var chart = ws.Drawings[0] as ExcelChartStandard;
            var cdr = chart.Drawings[1];
            cdr.SetSize(200);
            var cdr2 = chart.Drawings[4];
            cdr2.SetSize(200);

            var chartShape = chart.Drawings.AddShape("MyShape", eShapeStyle.Diamond);
            chartShape.Fill.Color = Color.LightSeaGreen;
            chartShape.Effect.SetPresetShadow(ePresetExcelShadowType.OuterRight);
            chartShape.SetPosition(144, 240);
            chartShape.SetSize(240, 144);
            var chartShape2 = chart.Drawings.AddShape("MyShape2", eShapeStyle.Diamond);
            chartShape2.Fill.Color = Color.Orange;
            chartShape2.Effect.SetPresetShadow(ePresetExcelShadowType.OuterRight);
            chartShape2.SetPosition(10000, 10000);
            chartShape2.SetSize(30);

            var chartPic = chart.Drawings.AddPicture("MyPic", Resources.GetImageFullFileName("epplusobject.png"));
            chartPic.SetPosition(0, 5000);
            //chartPic.SetSize(200);

            var myPic = Properties.Resources.GetOLEObjectFullFileName("SampleIcon.bmp");
            using (FileStream fileStream = new FileStream(myPic, FileMode.Open, FileAccess.Read))
            {
                var chartPic2 = chart.Drawings.AddPicture("MyPic2", fileStream);
                chartPic2.SetPosition(0, 5000);
                chartPic2.SetSize(200);
            }

            var shp1 = chart.Drawings.AddShape("level1", eShapeStyle.Star10);
            shp1.SetPosition(150, 250);
            var shp2 = chart.Drawings.AddShape("level2", eShapeStyle.UpDownArrow);
            shp2.SetPosition(200, 200);
            var shp3 = chart.Drawings.AddShape("level3", eShapeStyle.Teardrop);
            shp3.SetPosition(100, 50);
            var shp4 = chart.Drawings.AddShape("level4", eShapeStyle.MathEqual);
            shp4.SetPosition(10, 60);
            var group1 = shp1.Group(shp2, shp3);
            var group2 = group1.Group(shp4);
            group2.SetPosition(0, 0);
            shp4.UnGroup();
            shp4.Copy(chart);
            var chart2 = (ExcelChartStandard)ws.Drawings.AddChart("Chart 3", eChartType.Line);
            chart2.Series.Add(ws.Cells["B2:B6"], ws.Cells["C2:C6"]);
            chart2.SetSize(480, 288);
            chart2.Drawings.AddShape("hsp", eShapeStyle.Can);
            shp4.Copy(chart2);
            chartPic.Copy(chart);
            chartPic.Copy(chart2);
            group2.Copy(chart);
            group2.Copy(chart2);

            chart.Copy(ws, 20, 0);
            chart.Drawings.Remove(shp4);

            SaveWorkbook(@"shapeInChartTest.xlsx", p);
        }

        private void CreateChartData(ExcelWorksheet ws)
        {
            ws.Cells["A1"].Value = "Cat1";
            ws.Cells["B1"].Value = "Cat2";
            ws.Cells["C1"].Value = "Cat3";
            ws.Cells["D1"].Value = "Cat4";

            ws.Cells["A2"].Value = 10;
            ws.Cells["B2"].Value = 20;
            ws.Cells["C2"].Value = 30;
            ws.Cells["D2"].Value = 40;

            ws.Cells["A3"].Value = 100;
            ws.Cells["B3"].Value = 200;
            ws.Cells["C3"].Value = 300;
            ws.Cells["D3"].Value = 400;
        }

        private void AddDataToChart(ExcelWorksheet ws, ExcelChart chart)
        {
            chart.Series.Add(ws.Cells["A1:A3"]);
            chart.Series.Add(ws.Cells["B1:B3"]);
            chart.Series.Add(ws.Cells["C1:C3"]);
            chart.Series.Add(ws.Cells["D1:D3"]);
        }

        [TestMethod]
        public void AddShape()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            CreateChartData(ws);
            var chart = (ExcelChartStandard)ws.Drawings.AddChart("Chart 2", eChartType.Line);
            chart.SetPosition(0, 0, 5, 0);
            AddDataToChart(ws, chart);

            Assert.IsTrue(chart.Drawings.Count == 0);
            var shape = chart.Drawings.AddShape("Shape 2", eShapeStyle.DownArrow);            
            shape.SetPosition(5, 10);
            Assert.IsTrue(chart.Drawings.Count == 1);
            SaveWorkbook("ChartShape.xlsx", p);
        }
        [TestMethod]
        public void AddPicture()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            CreateChartData(ws);
            var chart = (ExcelChartStandard)ws.Drawings.AddChart("Chart 2", eChartType.Line);
            chart.SetPosition(0, 0, 5, 0);
            AddDataToChart(ws, chart);
            var pic = Properties.Resources.GetOLEObjectFullFileName("SampleIcon.bmp");
            Assert.IsTrue(chart.Drawings.Count == 0);
            chart.Drawings.AddPicture("Picture 2", pic);
            Assert.IsTrue(chart.Drawings.Count == 1);
            FileInfo picInfo = new FileInfo(pic);
            chart.Drawings.AddPicture("Picture 3", picInfo);
            Assert.IsTrue(chart.Drawings.Count == 2);
            using (FileStream fileStream = new FileStream(pic, FileMode.Open, FileAccess.Read))
            {
                chart.Drawings.AddPicture("Picture 4", fileStream);
            }
            Assert.IsTrue(chart.Drawings.Count == 3);
            SaveWorkbook("ChartPicture.xlsx", p); 
        }

        [TestMethod]
        public void GroupShapesWithShapes()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            CreateChartData(ws);
            var chart = (ExcelChartStandard)ws.Drawings.AddChart("Chart 2", eChartType.Line);
            chart.SetPosition(0, 0, 5, 0);
            AddDataToChart(ws, chart);

            var arrow = chart.Drawings.AddShape("Arrow", eShapeStyle.UpArrow);
            var equal = chart.Drawings.AddShape("Equal", eShapeStyle.MathEqual);
            var roundRect = chart.Drawings.AddShape("RoundRect", eShapeStyle.Round1Rect);
            var triangle = chart.Drawings.AddShape("Triangle", eShapeStyle.Triangle);

            Assert.IsTrue(chart.Drawings.Count == 4);
            var group = arrow.Group(equal, roundRect, triangle);
            Assert.IsTrue(chart.Drawings.Count == 1);
        }
        [TestMethod]
        public void GroupShapesWithPictures()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            CreateChartData(ws);
            var chart = (ExcelChartStandard)ws.Drawings.AddChart("Chart 2", eChartType.Line);
            chart.SetPosition(0, 0, 5, 0);
            AddDataToChart(ws, chart);

            var pic = Properties.Resources.GetOLEObjectFullFileName("SampleIcon.bmp");
            var pic1 = chart.Drawings.AddPicture("Pic 1", pic);
            pic1.SetPosition(10, 10);
            var pic2 = chart.Drawings.AddPicture("Pic 2", pic);
            pic1.SetPosition(20, 20);
            var pic3 = chart.Drawings.AddPicture("Pic 3", pic);
            pic1.SetPosition(30, 30);
            var pic4 = chart.Drawings.AddPicture("Pic 4", pic);
            pic1.SetPosition(40, 40);
            var pic5 = chart.Drawings.AddPicture("Pic 5", pic);
            pic1.SetPosition(50, 50);

            Assert.IsTrue(chart.Drawings.Count == 5);
            var group = pic1.Group(pic2, pic3, pic4, pic5);
            Assert.IsTrue(chart.Drawings.Count == 1);
        }
        [TestMethod]
        public void GroupShapesWithGroupShapes()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            CreateChartData(ws);
            var chart = (ExcelChartStandard)ws.Drawings.AddChart("Chart 2", eChartType.Line);
            chart.SetPosition(0, 0, 5, 0);
            AddDataToChart(ws, chart);

            var arrow = chart.Drawings.AddShape("Arrow", eShapeStyle.UpArrow);
            var equal = chart.Drawings.AddShape("Equal", eShapeStyle.MathEqual);
            var roundRect = chart.Drawings.AddShape("RoundRect", eShapeStyle.Round1Rect);
            var triangle = chart.Drawings.AddShape("Triangle", eShapeStyle.Triangle);

            Assert.IsTrue(chart.Drawings.Count == 4);
            var group1 = arrow.Group(equal);
            var group2 = group1.Group(roundRect, triangle);
            Assert.IsTrue(chart.Drawings.Count == 1);
        }
        [TestMethod]
        public void GroupShapesMixed()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            CreateChartData(ws);
            var chart = (ExcelChartStandard)ws.Drawings.AddChart("Chart 2", eChartType.Line);
            chart.SetPosition(0, 0, 5, 0);
            AddDataToChart(ws, chart);

            var arrow = chart.Drawings.AddShape("Arrow", eShapeStyle.UpArrow);
            var equal = chart.Drawings.AddShape("Equal", eShapeStyle.MathEqual);

            var pic = Properties.Resources.GetOLEObjectFullFileName("SampleIcon.bmp");
            var pic1 = chart.Drawings.AddPicture("Pic 1", pic);
            pic1.SetPosition(10, 10);
            var pic2 = chart.Drawings.AddPicture("Pic 2", pic);
            pic1.SetPosition(20, 20);
            Assert.IsTrue(chart.Drawings.Count == 4);
            var group1 = arrow.Group(pic1);
            var group2 = group1.Group(pic2, equal);
            Assert.IsTrue(chart.Drawings.Count == 1);
        }

        [TestMethod]
        public void CopyShape()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            CreateChartData(ws);
            var chart = (ExcelChartStandard)ws.Drawings.AddChart("Chart 2", eChartType.Line);
            chart.SetPosition(0, 0, 5, 0);
            AddDataToChart(ws, chart);
            var cshape = chart.Drawings.AddShape("Shape 2", eShapeStyle.DownArrow);
            Assert.IsTrue(chart.Drawings.Count == 1);
            cshape.Copy(chart);
            Assert.IsTrue(chart.Drawings.Count == 2);
        }
        [TestMethod]
        public void CopyPicture()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            CreateChartData(ws);
            var chart = (ExcelChartStandard)ws.Drawings.AddChart("Chart 2", eChartType.Line);
            chart.SetPosition(0, 0, 5, 0);
            AddDataToChart(ws, chart);
            var pic = Properties.Resources.GetOLEObjectFullFileName("SampleIcon.bmp");
            var cpic = chart.Drawings.AddPicture("Picture 2", pic);
            Assert.IsTrue(chart.Drawings.Count == 1);
            cpic.Copy(chart);
            Assert.IsTrue(chart.Drawings.Count == 2);
        }
        [TestMethod]
        public void CopyGroupShape()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            CreateChartData(ws);
            var chart = (ExcelChartStandard)ws.Drawings.AddChart("Chart 2", eChartType.Line);
            chart.SetPosition(0, 0, 5, 0);
            AddDataToChart(ws, chart);
            var cshape1 = chart.Drawings.AddShape("Shape 2", eShapeStyle.DownArrow);
            var cshape2 = chart.Drawings.AddShape("Shape 3", eShapeStyle.Teardrop);
            var pic = Properties.Resources.GetOLEObjectFullFileName("SampleIcon.bmp");
            var cpic = chart.Drawings.AddPicture("Picture 2", pic);
            var group = cshape1.Group(cshape2, cpic);
            Assert.IsTrue(chart.Drawings.Count == 1);
            group.Copy(chart);
            Assert.IsTrue(chart.Drawings.Count == 2);
        }
        [TestMethod]
        public void CopyWorksheet()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            CreateChartData(ws);
            var chart = (ExcelChartStandard)ws.Drawings.AddChart("Chart 2", eChartType.Line);
            chart.SetPosition(0, 0, 5, 0);
            AddDataToChart(ws, chart);
            var cshape1 = chart.Drawings.AddShape("Shape 2", eShapeStyle.DownArrow);
            var cshape2 = chart.Drawings.AddShape("Shape 3", eShapeStyle.Teardrop);
            var pic = Properties.Resources.GetOLEObjectFullFileName("SampleIcon.bmp");
            var cpic = chart.Drawings.AddPicture("Picture 2", pic);
            var group = cshape1.Group(cshape2, cpic);
            Assert.IsTrue(chart.Drawings.Count == 1);

            var ws2 = p.Workbook.Worksheets.Add("Sheet 2");
            CreateChartData(ws2);
            var chart2 = (ExcelChartStandard)ws2.Drawings.AddChart("Chart 3", eChartType.Pie);
            AddDataToChart(ws2 , chart2);
            group.Copy(chart2);
            Assert.IsTrue(chart2.Drawings.Count == 1);
        }
        [TestMethod]
        public void DeleteShape()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            CreateChartData(ws);
            var chart = (ExcelChartStandard)ws.Drawings.AddChart("Chart 2", eChartType.Line);
            chart.SetPosition(0, 0, 5, 0);
            AddDataToChart(ws, chart);
            var cshape = chart.Drawings.AddShape("Shape 2", eShapeStyle.DownArrow);
            Assert.IsTrue(chart.Drawings.Count == 1);
            chart.Drawings.Remove(cshape);
            Assert.IsTrue(chart.Drawings.Count == 0);
        }
        [TestMethod]
        public void DeletePicture()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            CreateChartData(ws);
            var chart = (ExcelChartStandard)ws.Drawings.AddChart("Chart 2", eChartType.Line);
            chart.SetPosition(0, 0, 5, 0);
            AddDataToChart(ws, chart);
            var pic = Properties.Resources.GetOLEObjectFullFileName("SampleIcon.bmp");
            var cpic = chart.Drawings.AddPicture("Picture 2", pic);
            Assert.IsTrue(chart.Drawings.Count == 1);
            chart.Drawings.Remove(cpic);
            Assert.IsTrue(chart.Drawings.Count == 0);
        }
        [TestMethod]
        public void DeleteGroupShape()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            CreateChartData(ws);
            var chart = (ExcelChartStandard)ws.Drawings.AddChart("Chart 2", eChartType.Line);
            chart.SetPosition(0, 0, 5, 0);
            AddDataToChart(ws, chart);

            var arrow = chart.Drawings.AddShape("Arrow", eShapeStyle.UpArrow);
            var equal = chart.Drawings.AddShape("Equal", eShapeStyle.MathEqual);
            var pic = Properties.Resources.GetOLEObjectFullFileName("SampleIcon.bmp");
            var pic1 = chart.Drawings.AddPicture("Pic 1", pic);
            pic1.SetPosition(10, 10);
            var pic2 = chart.Drawings.AddPicture("Pic 2", pic);
            pic1.SetPosition(20, 20);
            var group1 = arrow.Group(pic1);
            var group2 = group1.Group(pic2, equal);
            Assert.IsTrue(chart.Drawings.Count == 1);
            chart.Drawings.Remove(group2);
            Assert.IsTrue(chart.Drawings.Count == 0);
        }
    }
}

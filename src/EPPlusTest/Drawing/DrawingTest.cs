/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * Required Notice: Copyright (C) EPPlus Software AB. 
 * https://epplussoftware.com
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "" as is "" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
  Date               Author                       Change
 *******************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *******************************************************************************/
using System;
using System.Drawing;
using System.IO;
using System.Xml;
using EPPlusTest.Properties;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System.Diagnostics;
using System.Reflection;
using OfficeOpenXml.Drawing.Theme;

namespace EPPlusTest
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestClass]
    public class DrawingTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenPackage("Drawing.xlsx", true);            
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            var dirName = _pck.File.DirectoryName;
            var fileName = _pck.File.FullName;

            SaveAndCleanup(_pck);
            if(File.Exists(fileName)) File.Copy(fileName, dirName + "\\DrawingRead.xlsx", true);
        }
        [TestMethod]
        public void ReadDrawing()
        {
            using (ExcelPackage pck = new ExcelPackage(new FileInfo(_worksheetPath + @"DrawingRead.xlsx")))
            {
                var ws = pck.Workbook.Worksheets["Pyramid"];
                if (ws == null) Assert.Inconclusive("Pyramid worksheet is missing");
                Assert.AreEqual(ws.Cells["V24"].Value, 104D);
                ws = pck.Workbook.Worksheets["Scatter"];
                if (ws == null) Assert.Inconclusive("Scatter worksheet is missing");
                var cht = ws.Drawings["ScatterChart1"] as ExcelScatterChart;
                Assert.AreEqual(cht.Title.Text, "Header  Text");
                cht.Title.Text = "Test";
                Assert.AreEqual(cht.Title.Text, "Test");
            }
        }
        [TestMethod]
        public void Picture()
         {
            var ws = _pck.Workbook.Worksheets.Add("Picture");
            var pic = ws.Drawings.AddPicture("Pic1", Resources.Test1);
            Assert.AreEqual(eDrawingType.Picture, pic.DrawingType);

            pic = ws.Drawings.AddPicture("Pic2", Resources.Test1);
            pic.SetPosition(150, 200);
            pic.Border.LineStyle = eLineStyle.Solid;
            pic.Border.Fill.Color = Color.DarkCyan;
            pic.Fill.Style = eFillStyle.SolidFill;
            pic.Fill.Color = Color.White;
            pic.Fill.Transparancy = 50;

            pic = ws.Drawings.AddPicture("Pic3", Resources.Test1);
            pic.SetPosition(400, 200);
            pic.SetSize(150);

            pic = ws.Drawings.AddPicture("Pic5", GetResourceFile("BitmapImage.gif"));
            pic.SetPosition(400, 200);
            pic.SetSize(150);

            ws.Column(1).Width = 53;
            ws.Column(4).Width = 58;

            pic = ws.Drawings.AddPicture("Pic6öäå", GetResourceFile("BitmapImage.gif"));
            pic.SetPosition(400, 400);
            pic.SetSize(100);

            pic = ws.Drawings.AddPicture("PicPixelSized", Resources.Test1);
            pic.SetPosition(800, 800);
            pic.SetSize(568*2, 66*2);
            var ws2 = _pck.Workbook.Worksheets.Add("Picture2");
            var fi = GetResourceFile("BitmapImage.gif");
            if (fi.Exists)
            {
                pic = ws2.Drawings.AddPicture("Pic7", fi);
            }
            else
            {
#if (!Core)
                TestContext.WriteLine("AG00021_.GIF does not exists. Skipping Pic7.");
#endif
            }
            var wsCopy = _pck.Workbook.Worksheets.Add("Picture3", ws2);
         }
        [TestMethod]
        public void ShapeURL()
        {
            var ws = _pck.Workbook.Worksheets.Add("Shape URL");

            ExcelHyperLink hl = new ExcelHyperLink("http://epplussoftware.com");
            hl.ToolTip = "Screen Tip";

            var shape = ws.Drawings.AddShape("ShapeUrl", eShapeStyle.Rect);

            shape.Hyperlink = new ExcelHyperLink("https://epplussoftware.com", UriKind.Absolute);
        }
        [TestMethod]
        public void PictureURL()
        {
            var ws = _pck.Workbook.Worksheets.Add("Pic URL");

            ExcelHyperLink hl = new ExcelHyperLink("https://epplussoftware.com");
            hl.ToolTip = "Screen Tip";

            ws.Drawings.AddPicture("Pic URI", Properties.Resources.Test1, hl);
        }
        [TestMethod]
        public void ChartURL()
        {
            var ws = _pck.Workbook.Worksheets.Add("Chart URL");

            ExcelHyperLink hl = new ExcelHyperLink("http://epplussoftware.com");
            hl.ToolTip = "Screen Tip";

            var areaChart = ws.Drawings.AddAreaChart("ShapeUrl", eAreaChartType.Area);

            areaChart.Hyperlink = new ExcelHyperLink("Q3", "");
        }

        //[TestMethod]
        //[Ignore]
        public void DrawingSizingAndPositioning()
         {
             var ws = _pck.Workbook.Worksheets.Add("DrawingPosSize");

             var pic = ws.Drawings.AddPicture("Pic1", Resources.Test1);
             pic.SetPosition(1, 0, 1, 0);

             pic = ws.Drawings.AddPicture("Pic2", Resources.Test1);
             pic.EditAs = eEditAs.Absolute;
             pic.SetPosition(10, 5, 1, 4);

             pic = ws.Drawings.AddPicture("Pic3", Resources.Test1);
             pic.EditAs = eEditAs.TwoCell;
             pic.SetPosition(20, 5, 2, 4);


             ws.Column(1).Width = 100;
             ws.Column(3).Width = 100;
         }

        [TestMethod]
         public void BarChart()
        {
            var ws = _pck.Workbook.Worksheets.Add("BarChart");            
            var chrt = ws.Drawings.AddChart("barChart", eChartType.BarClustered) as ExcelBarChart;
            chrt.SetPosition(50, 50);
            chrt.SetSize(800, 300);
            AddTestSerie(ws, chrt);
            chrt.VaryColors = true;
            chrt.XAxis.Orientation = eAxisOrientation.MaxMin;
            chrt.XAxis.MajorTickMark = eAxisTickMark.In;
            chrt.XAxis.Format = "yyyy-MM";
            chrt.YAxis.Orientation = eAxisOrientation.MaxMin;
            chrt.YAxis.MinorTickMark = eAxisTickMark.Out;
            chrt.ShowHiddenData = true;
            chrt.DisplayBlanksAs = eDisplayBlanksAs.Zero;
            chrt.Title.RichText.Text = "Barchart Test";
            chrt.Title.RichText[0].LatinFont = "Arial";

            chrt.GapWidth = 5;
            Assert.IsTrue(chrt.ChartType == eChartType.BarClustered, "Invalid Charttype");
            Assert.IsTrue(chrt.Direction == eDirection.Bar, "Invalid Bardirection");
            Assert.IsTrue(chrt.Grouping == eGrouping.Clustered, "Invalid Grouping");
            Assert.IsTrue(chrt.Shape == eShape.Box, "Invalid Shape");
        }

        private static void AddTestSerie(ExcelWorksheet ws, ExcelChart chrt)
        {
            AddTestData(ws);
            chrt.Series.Add("'" + ws.Name + "'!V19:V24", "'" + ws.Name + "'!U19:U24");
        }

        private static void AddTestData(ExcelWorksheet ws)
        {
            ws.Cells["U19"].Value = new DateTime(2009, 12, 31);
            ws.Cells["U20"].Value = new DateTime(2010, 1, 1);
            ws.Cells["U21"].Value = new DateTime(2010, 1, 2);
            ws.Cells["U22"].Value = new DateTime(2010, 1, 3);
            ws.Cells["U23"].Value = new DateTime(2010, 1, 4);
            ws.Cells["U24"].Value = new DateTime(2010, 1, 5);
            ws.Cells["U19:U24"].Style.Numberformat.Format = "yyyy-mm-dd";

            ws.Cells["V19"].Value = 100;
            ws.Cells["V20"].Value = 102;
            ws.Cells["V21"].Value = 101;
            ws.Cells["V22"].Value = 103;
            ws.Cells["V23"].Value = 105;
            ws.Cells["V24"].Value = 104;

            ws.Cells["W19"].Value = 105;
            ws.Cells["W20"].Value = 108;
            ws.Cells["W21"].Value = 104;
            ws.Cells["W22"].Value = 121;
            ws.Cells["W23"].Value = 103;
            ws.Cells["W24"].Value = 109;


            ws.Cells["X19"].Value = "öäå";
            ws.Cells["X20"].Value = "ÖÄÅ";
            ws.Cells["X21"].Value = "üÛ";
            ws.Cells["X22"].Value = "&%#¤";
            ws.Cells["X23"].Value = "ÿ";
            ws.Cells["X24"].Value = "û";
        }
        [TestMethod]
        public void PieChart()
        {
            var expected = Color.SteelBlue;
            var ws = _pck.Workbook.Worksheets.Add("PieChart");
            var chrt = ws.Drawings.AddChart("pieChart", eChartType.Pie) as ExcelPieChart;
            
            AddTestSerie(ws, chrt);

            chrt.To.Row = 25;
            chrt.To.Column = 12;

            chrt.DataLabel.ShowPercent = true;
            chrt.Legend.Font.Color = expected;
            chrt.Title.Border.Fill.Style = eFillStyle.SolidFill;
            chrt.Legend.Position = eLegendPosition.TopRight;
            Assert.IsTrue(chrt.ChartType == eChartType.Pie, "Invalid Charttype");
            Assert.IsTrue(chrt.VaryColors);            
            var expectedArgb = expected.ToArgb() & 0xFFFFFF;    //Without alpha part
            Assert.AreEqual(expectedArgb, chrt.Legend.Font.Color.ToArgb());
            Assert.AreEqual(chrt.Legend.Font.Fill.Style, eFillStyle.SolidFill);
            Assert.AreEqual(chrt.Legend.Font.Fill.SolidFill.Color.ColorType, eDrawingColorType.Rgb);
            Assert.AreEqual(expected.ToArgb(), chrt.Legend.Font.Fill.SolidFill.Color.RgbColor.Color.ToArgb());
            chrt.Title.Text = "Piechart";
        }
        [TestMethod]
        public void PieOfChart()
        {
            var ws = _pck.Workbook.Worksheets.Add("PieOfChart");
            var chrt = ws.Drawings.AddChart("pieOfChart", eChartType.BarOfPie) as ExcelOfPieChart;

            AddTestSerie(ws, chrt);

            chrt.To.Row = 25;
            chrt.To.Column = 12;

            chrt.DataLabel.ShowPercent = true;
            chrt.Legend.Font.Color = Color.SteelBlue;
            chrt.Title.Border.Fill.Style = eFillStyle.SolidFill;
            chrt.Legend.Position = eLegendPosition.TopRight;
            Assert.IsTrue(chrt.ChartType == eChartType.BarOfPie, "Invalid Charttype");
            chrt.Title.Text = "Piechart";
        }
        [TestMethod]
        public void PieChart3D()
        {
            var ws = _pck.Workbook.Worksheets.Add("PieChart3d");
            var chrt = ws.Drawings.AddChart("pieChart3d", eChartType.Pie3D) as ExcelPieChart;
            AddTestSerie(ws, chrt);

            chrt.To.Row = 25;
            chrt.To.Column = 12;

            chrt.DataLabel.ShowValue = true;
            chrt.Legend.Position = eLegendPosition.Left;
            chrt.ShowHiddenData = false;
            chrt.DisplayBlanksAs = eDisplayBlanksAs.Gap;
            chrt.Title.RichText.Add("Pie RT Title add");
            Assert.IsTrue(chrt.ChartType == eChartType.Pie3D, "Invalid Charttype");
            Assert.IsTrue(chrt.VaryColors);

        }
        [TestMethod]
        public void Scatter()
        {
            var ws = _pck.Workbook.Worksheets.Add("Scatter");
            var chrt = ws.Drawings.AddChart("ScatterChart1", eChartType.XYScatterSmoothNoMarkers) as ExcelScatterChart;
            AddTestSerie(ws, chrt);
           // chrt.Series[0].Marker = eMarkerStyle.Diamond;
            chrt.To.Row = 23;
            chrt.To.Column = 12;
            //chrt.Title.Text = "Header Text";
            var r1=chrt.Title.RichText.Add("Header");
            r1.Bold = true;
            var r2=chrt.Title.RichText.Add("  Text");
            r2.UnderLine = eUnderLineType.WavyHeavy;
            
            chrt.Title.Fill.Style = eFillStyle.SolidFill;
            chrt.Title.Fill.Color = Color.LightBlue;
            chrt.Title.Fill.Transparancy = 50;
            chrt.VaryColors = true;
            ExcelScatterChartSerie ser = chrt.Series[0] as ExcelScatterChartSerie;
            ser.DataLabel.Position = eLabelPosition.Center;
            ser.DataLabel.ShowValue = true;
            ser.DataLabel.ShowCategory = true;
            ser.DataLabel.Fill.Color = Color.BlueViolet;
            ser.DataLabel.Font.Color = Color.White;
            ser.DataLabel.Font.Italic = true;
            ser.DataLabel.Font.SetFromFont("bookman old style", 8);
            Assert.IsTrue(chrt.ChartType == eChartType.XYScatterSmoothNoMarkers, "Invalid Charttype");
            chrt.Series[0].Header = "Test serie";
            chrt = ws.Drawings.AddChart("ScatterChart2", eChartType.XYScatterSmooth) as ExcelScatterChart;
            chrt.Series.Add("U19:U24", "V19:V24");
            
            chrt.From.Column = 0;
            chrt.From.Row=25;
            chrt.To.Row = 53;
            chrt.To.Column = 12;
            chrt.Legend.Position = eLegendPosition.Bottom;
            
            ////chrt.Series[0].DataLabel.Position = eLabelPosition.Center;
            //Assert.IsTrue(chrt.ChartType == eChartType.XYScatter, "Invalid Charttype");

        }
       [TestMethod]
        public void Bubble()
        {
            var ws = _pck.Workbook.Worksheets.Add("Bubble");
            var chrt = ws.Drawings.AddChart("Bubble", eChartType.Bubble) as ExcelBubbleChart;
            AddTestData(ws);

            chrt.Series.Add("V19:V24", "U19:U24");

            chrt = ws.Drawings.AddChart("Bubble3d", eChartType.Bubble3DEffect) as ExcelBubbleChart;
            ws.Cells["W19"].Value = 1;
            ws.Cells["W20"].Value = 1;
            ws.Cells["W21"].Value = 2;
            ws.Cells["W22"].Value = 2;
            ws.Cells["W23"].Value = 3;
            ws.Cells["W24"].Value = 4;

            chrt.Series.Add("V19:V24", "U19:U24", "W19:W24");
            chrt.Style = eChartStyle.Style25;

            // chrt.Series[0].Marker = eMarkerStyle.Diamond;
            chrt.From.Row = 23;
            chrt.From.Column = 12;
            chrt.To.Row = 33;
            chrt.To.Column = 22;
            chrt.Title.Text = "Header Text";
            

        }
        [TestMethod]
        public void Radar()
        {
            var ws = _pck.Workbook.Worksheets.Add("Radar");
            AddTestData(ws);

            var chrt = ws.Drawings.AddChart("Radar1", eChartType.Radar) as ExcelRadarChart;
            var s=chrt.Series.Add("V19:V24", "U19:U24");
            s.Header = "serie1";
            // chrt.Series[0].Marker = eMarkerStyle.Diamond;
            chrt.From.Row = 23;
            chrt.From.Column = 12;
            chrt.To.Row = 38;
            chrt.To.Column = 22;
            chrt.Title.Text = "Radar Chart 1";

            chrt = ws.Drawings.AddChart("Radar2", eChartType.RadarFilled) as ExcelRadarChart;
            s = chrt.Series.Add("V19:V24", "U19:U24");
            s.Header = "serie1";
            // chrt.Series[0].Marker = eMarkerStyle.Diamond;
            chrt.From.Row = 43;
            chrt.From.Column = 12;
            chrt.To.Row = 58;
            chrt.To.Column = 22;
            chrt.Title.Text = "Radar Chart 2";

            chrt = ws.Drawings.AddChart("Radar3", eChartType.RadarMarkers) as ExcelRadarChart;
            var rs = (ExcelRadarChartSerie)chrt.Series.Add("V19:V24", "U19:U24");
            rs.Header = "serie1";
            rs.Marker.Style = eMarkerStyle.Star;
            rs.Marker.Size = 14;

            // chrt.Series[0].Marker = eMarkerStyle.Diamond;
            chrt.From.Row = 63;
            chrt.From.Column = 12;
            chrt.To.Row = 78;
            chrt.To.Column = 22;
            chrt.Title.Text = "Radar Chart 3";
        }
        [TestMethod]
        public void Surface()
        {
            var ws = _pck.Workbook.Worksheets.Add("Surface");
            AddTestData(ws);

            var chrt = ws.Drawings.AddChart("Surface1", eChartType.Surface) as ExcelSurfaceChart;
            var s = chrt.Series.Add("V19:V24", "U19:U24");
            var s2 = chrt.Series.Add("W19:W24", "U19:U24");
            s.Header = "serie1";
            // chrt.Series[0].Marker = eMarkerStyle.Diamond;
            chrt.From.Row = 23;
            chrt.From.Column = 12;
            chrt.To.Row = 38;
            chrt.To.Column = 22;
            chrt.Title.Text = "Surface Chart 1";

            //chrt = ws.Drawings.AddChart("Surface", eChartType.RadarFilled) as ExcelRadarChart;
            //s = chrt.Series.Add("V19:V24", "U19:U24");
            //s.Header = "serie1";
            //// chrt.Series[0].Marker = eMarkerStyle.Diamond;
            //chrt.From.Row = 43;
            //chrt.From.Column = 12;
            //chrt.To.Row = 58;
            //chrt.To.Column = 22;
            //chrt.Title.Text = "Radar Chart 2";

            //chrt = ws.Drawings.AddChart("Radar3", eChartType.RadarMarkers) as ExcelRadarChart;
            //var rs = (ExcelRadarChartSerie)chrt.Series.Add("V19:V24", "U19:U24");
            //rs.Header = "serie1";
            //rs.Marker = eMarkerStyle.Star;
            //rs.MarkerSize = 14;

            //// chrt.Series[0].Marker = eMarkerStyle.Diamond;
            //chrt.From.Row = 63;
            //chrt.From.Column = 12;
            //chrt.To.Row = 78;
            //chrt.To.Column = 22;
            //chrt.Title.Text = "Radar Chart 3";
        }
        [TestMethod]
        public void Pyramid()
        {
            var ws = _pck.Workbook.Worksheets.Add("Pyramid");
            var chrt = ws.Drawings.AddChart("Pyramid1", eChartType.PyramidCol) as ExcelBarChart;
            AddTestSerie(ws, chrt);
            chrt.VaryColors = true;
            chrt.To.Row = 23;
            chrt.To.Column = 12;
            chrt.Title.Text = "Header Text";
            chrt.Title.Fill.Style= eFillStyle.SolidFill;
            chrt.Title.Fill.Color = Color.DarkBlue;
            chrt.DataLabel.ShowValue = true;

            chrt.Border.LineCap = eLineCap.Round;            
            chrt.Border.LineStyle = eLineStyle.LongDashDotDot;
            chrt.Border.Fill.Style = eFillStyle.SolidFill;
            chrt.Border.Fill.Color = Color.Blue;

            chrt.Fill.Color = Color.LightCyan;
            chrt.PlotArea.Fill.Color = Color.White;
            chrt.PlotArea.Border.Fill.Style = eFillStyle.SolidFill;
            chrt.PlotArea.Border.Fill.Color = Color.Beige;
            chrt.PlotArea.Border.LineStyle = eLineStyle.LongDash;

            chrt.Legend.Fill.Color = Color.Aquamarine;
            chrt.Legend.Position = eLegendPosition.Top;
            chrt.Axis[0].Fill.Style = eFillStyle.SolidFill;
            chrt.Axis[0].Fill.Color = Color.Black;
            chrt.Axis[0].Font.Color = Color.White;

            chrt.Axis[1].Fill.Style = eFillStyle.SolidFill;
            chrt.Axis[1].Fill.Color = Color.LightSlateGray;
            chrt.Axis[1].Font.Color = Color.DarkRed;

            chrt.DataLabel.Font.Bold = true;
            chrt.DataLabel.Fill.Color = Color.LightBlue;
            chrt.DataLabel.Border.Fill.Style=eFillStyle.SolidFill;
            chrt.DataLabel.Border.Fill.Color=Color.Black;
            chrt.DataLabel.Border.LineStyle = eLineStyle.Solid;
        }
        [TestMethod]
        public void Cone()
        {
            var ws = _pck.Workbook.Worksheets.Add("Cone");
            var chrt = ws.Drawings.AddChart("Cone1", eChartType.ConeBarClustered) as ExcelBarChart;
            AddTestSerie(ws, chrt);
            chrt.VaryColors = true;
            chrt.SetSize(200);
            chrt.Title.Text = "Cone bar";
            chrt.Series[0].Header = "Serie 1";
            chrt.Legend.Position = eLegendPosition.Right;
            chrt.Axis[1].DisplayUnit = 100000;
            Assert.AreEqual(chrt.Axis[1].DisplayUnit, 100000);
        }
        [TestMethod]
        public void Column()
        {
            var ws = _pck.Workbook.Worksheets.Add("Column");
            var chrt = ws.Drawings.AddChart("Column1", eChartType.ColumnClustered3D) as ExcelBarChart;
            AddTestSerie(ws, chrt);
            chrt.VaryColors = true;
            chrt.View3D.RightAngleAxes = true;
            chrt.View3D.DepthPercent = 99;
            //chrt.View3D.HeightPercent = 99;
            chrt.View3D.RightAngleAxes = true;
            chrt.SetSize(200);
            chrt.Title.Text = "Column";
            chrt.Series[0].Header = "Serie 1";
            chrt.Locked = false;
            chrt.Print = false;
            chrt.EditAs = eEditAs.TwoCell;
            chrt.Axis[1].DisplayUnit = 10020;
            Assert.AreEqual(chrt.Axis[1].DisplayUnit, 10020);
        }
        [TestMethod]
        public void Doughnut()
        {
            var ws = _pck.Workbook.Worksheets.Add("Doughnut");
            var chrt = ws.Drawings.AddChart("Doughnut1", eChartType.DoughnutExploded) as ExcelDoughnutChart;
            AddTestSerie(ws, chrt);
            chrt.SetSize(200);
            chrt.Title.Text = "Doughnut Exploded";
            chrt.Series[0].Header = "Serie 1";
            chrt.EditAs = eEditAs.Absolute;
        }
        [TestMethod]
        public void Line()
        {
            var ws = _pck.Workbook.Worksheets.Add("Line");
            var chrt = ws.Drawings.AddChart("Line1", eChartType.Line) as ExcelLineChart;
            Assert.AreEqual(eDrawingType.Chart, chrt.DrawingType);
            AddTestSerie(ws, chrt);
            chrt.SetSize(150);
            chrt.VaryColors = true;
            chrt.Smooth = false;
            chrt.Title.Text = "Line 3D";
            chrt.Series[0].Header = "Line serie 1";
            chrt.Axis[0].MajorGridlines.Fill.Color = Color.Black;
            chrt.Axis[0].MajorGridlines.LineStyle = eLineStyle.Dot;

            var tl = chrt.Series[0].TrendLines.Add(eTrendLine.Polynomial);
            tl.Name = "Test";
            tl.DisplayRSquaredValue = true;
            tl.DisplayEquation = true;
            tl.Forward = 15;
            tl.Backward = 1;
            tl.Intercept = 6;
            tl.Order = 5;
            
            tl = chrt.Series[0].TrendLines.Add(eTrendLine.MovingAvgerage);
            chrt.Fill.Color = Color.LightSteelBlue;
            chrt.Border.LineStyle = eLineStyle.Dot;
            chrt.Border.Fill.Color=Color.Black;

            chrt.Legend.Font.Color = Color.Red;
            chrt.Legend.Font.Strike = eStrikeType.Double;
            chrt.Title.Font.Color = Color.DarkGoldenrod;
            chrt.Title.Font.LatinFont = "Arial";
            chrt.Title.Font.Bold = true;
            chrt.Title.Fill.Color = Color.White;
            chrt.Title.Border.Fill.Style = eFillStyle.SolidFill;
            chrt.Title.Border.LineStyle = eLineStyle.LongDashDotDot;
            chrt.Title.Border.Fill.Color = Color.Tomato;
            chrt.DataLabel.ShowSeriesName = true;
            chrt.DataLabel.ShowLeaderLines=true;
            chrt.EditAs = eEditAs.OneCell;
            chrt.DisplayBlanksAs = eDisplayBlanksAs.Span;
            chrt.Axis[0].Title.Text = "Axis 0";
            chrt.Axis[0].Title.Rotation = 90;
            chrt.Axis[0].Title.Overlay = true;
            chrt.Axis[1].Title.Text = "Axis 1";
            chrt.Axis[1].Title.AnchorCtr = true;
            chrt.Axis[1].Title.TextVertical = eTextVerticalType.Vertical270;
            chrt.Axis[1].Title.Border.LineStyle=eLineStyle.LongDashDotDot;

            chrt.Series[0].CreateCache();

            chrt.StyleManager.CreateEmptyStyle(eChartStyle.Style2);

            Assert.AreEqual(0, chrt.StyleManager.Style.AxisTitle.BorderReference.Index);
            Assert.AreEqual(eDrawingColorType.None, chrt.StyleManager.Style.AxisTitle.BorderReference.Color.ColorType);

            Assert.AreEqual(eThemeFontCollectionType.Minor, chrt.StyleManager.Style.AxisTitle.FontReference.Index);
            Assert.AreEqual(eDrawingColorType.None, chrt.StyleManager.Style.AxisTitle.FontReference.Color.ColorType);
        }
        [TestMethod]
        public void LineMarker()
        {
            var ws = _pck.Workbook.Worksheets.Add("LineMarker1");
            var chrt = ws.Drawings.AddChart("Line1", eChartType.LineMarkers) as ExcelLineChart;
            AddTestSerie(ws, chrt);
            chrt.SetSize(150);
            chrt.Title.Text = "Line Markers";
            chrt.Series[0].Header = "Line serie 1";
            ((ExcelLineChartSerie)chrt.Series[0]).Marker.Style = eMarkerStyle.Plus;

            var chrt2 = ws.Drawings.AddChart("Line2", eChartType.LineMarkers) as ExcelLineChart;
            AddTestSerie(ws, chrt2);
            chrt2.SetPosition(500,0);
            chrt2.SetSize(150);
            chrt2.Title.Text = "Line Markers";
            var serie = (ExcelLineChartSerie)chrt2.Series[0];
            serie.Marker.Style = eMarkerStyle.X;

        }
        [TestMethod]
        public void Drawings()
        {
            var ws = _pck.Workbook.Worksheets.Add("Shapes");
            int y=100, i=1;
            foreach(eShapeStyle style in Enum.GetValues(typeof(eShapeStyle)))
            {
                var shape = ws.Drawings.AddShape("shape"+i.ToString(), style);
                Assert.AreEqual(eDrawingType.Shape, shape.DrawingType);
                shape.SetPosition(y, 100);
                shape.SetSize(300, 300);
                y += 400;
                shape.Text = style.ToString();
                i++;
            }

            (ws.Drawings["shape1"] as ExcelShape).TextAnchoring = eTextAnchoringType.Top;            
            var rt = (ws.Drawings["shape1"] as ExcelShape).RichText.Add("Added formatted richtext");
            (ws.Drawings["shape1"] as ExcelShape).LockText = false;
            rt.Bold = true;
            rt.Color = Color.Aquamarine;
            rt.Italic = true;
            rt.Size = 17;
            (ws.Drawings["shape2"] as ExcelShape).TextVertical = eTextVerticalType.Vertical;
            rt = (ws.Drawings["shape2"] as ExcelShape).RichText.Add("\r\nAdded formatted richtext");
            rt.Bold = true;
            rt.Color = Color.DarkGoldenrod ;
            rt.SetFromFont("Times new roman", 18, false, false, true);
            rt.UnderLineColor = Color.Green;


            (ws.Drawings["shape3"] as ExcelShape).TextAnchoring=eTextAnchoringType.Bottom;
            (ws.Drawings["shape3"] as ExcelShape).TextAnchoringControl=true ;

            (ws.Drawings["shape4"] as ExcelShape).TextVertical = eTextVerticalType.Vertical270;
            (ws.Drawings["shape4"] as ExcelShape).TextAnchoring = eTextAnchoringType.Top;

            (ws.Drawings["shape5"] as ExcelShape).Fill.Style=eFillStyle.SolidFill;
            (ws.Drawings["shape5"] as ExcelShape).Fill.Color=Color.Red;
            (ws.Drawings["shape5"] as ExcelShape).Fill.Transparancy = 50;

            (ws.Drawings["shape6"] as ExcelShape).Fill.Style = eFillStyle.NoFill;
            (ws.Drawings["shape6"] as ExcelShape).Font.Color = Color.Black;
            (ws.Drawings["shape6"] as ExcelShape).Border.Fill.Color = Color.Black;

            (ws.Drawings["shape7"] as ExcelShape).Fill.Style = eFillStyle.SolidFill;
            (ws.Drawings["shape7"] as ExcelShape).Fill.Color=Color.Gray;
            (ws.Drawings["shape7"] as ExcelShape).Border.Fill.Style=eFillStyle.SolidFill;
            (ws.Drawings["shape7"] as ExcelShape).Border.Fill.Color = Color.Black;
            (ws.Drawings["shape7"] as ExcelShape).Border.Fill.Transparancy=43;
            (ws.Drawings["shape7"] as ExcelShape).Border.LineCap=eLineCap.Round;
            (ws.Drawings["shape7"] as ExcelShape).Border.LineStyle = eLineStyle.LongDash;
            (ws.Drawings["shape7"] as ExcelShape).Font.UnderLineColor = Color.Blue;
            (ws.Drawings["shape7"] as ExcelShape).Font.Color = Color.Black;
            (ws.Drawings["shape7"] as ExcelShape).Font.Bold = true;
            (ws.Drawings["shape7"] as ExcelShape).Font.LatinFont = "Arial";
            (ws.Drawings["shape7"] as ExcelShape).Font.ComplexFont = "Arial";
            (ws.Drawings["shape7"] as ExcelShape).Font.Italic = true;
            (ws.Drawings["shape7"] as ExcelShape).Font.UnderLine = eUnderLineType.Dotted;

            (ws.Drawings["shape8"] as ExcelShape).Fill.Style = eFillStyle.SolidFill;
            (ws.Drawings["shape8"] as ExcelShape).Font.LatinFont = "Miriam";
            (ws.Drawings["shape8"] as ExcelShape).Font.UnderLineColor = Color.CadetBlue;
            (ws.Drawings["shape8"] as ExcelShape).Font.UnderLine = eUnderLineType.Single;

            (ws.Drawings["shape9"] as ExcelShape).TextAlignment = eTextAlignment.Right;

            (ws.Drawings["shape120"] as ExcelShape).TailEnd.Style = eEndStyle.Oval;
            (ws.Drawings["shape120"] as ExcelShape).TailEnd.Width = eEndSize.Large;
            (ws.Drawings["shape120"] as ExcelShape).TailEnd.Height = eEndSize.Large;
            (ws.Drawings["shape120"] as ExcelShape).HeadEnd.Style= eEndStyle.Arrow;
            (ws.Drawings["shape120"] as ExcelShape).HeadEnd.Height = eEndSize.Small;
            (ws.Drawings["shape120"] as ExcelShape).HeadEnd.Width = eEndSize.Small;
        }
        [TestMethod]
        //[Ignore]
        public void DrawingWorksheetCopy()
        {
            using (var pck = OpenPackage("Drawingread.xlsx"))
            {
                var ws = pck.Workbook.Worksheets["Shapes"];
                if (ws == null) Assert.Inconclusive("Shapes worksheet is missing");

                var wsShapes = pck.Workbook.Worksheets.Add("Copy Shapes", ws);
                Assert.AreEqual(187, wsShapes.Drawings.Count);

                ws = pck.Workbook.Worksheets["Scatter"];
                if (ws == null) Assert.Inconclusive("Scatter worksheet is missing");
                var wsScatterChart = pck.Workbook.Worksheets.Add("Copy Scatter", ws);
                Assert.AreEqual(2, wsScatterChart.Drawings.Count);
                var chart1 = wsScatterChart.Drawings[0].As.Chart.ScatterChart;
                Assert.AreEqual(1, chart1.Series.Count);
                Assert.AreEqual("'Copy Scatter'!V19:V24", chart1.Series[0].Series);

                ws = pck.Workbook.Worksheets["Picture"];
                if (ws == null) Assert.Inconclusive("Picture worksheet is missing");
                var wsPicture = pck.Workbook.Worksheets.Add("Copy Picture", ws);

                pck.SaveAs(new FileInfo(_worksheetPath+"DrawingCopied.xlsx"));
            }
        }
        [TestMethod]
        public void Line2Test()
        {
           ExcelWorksheet worksheet = _pck.Workbook.Worksheets.Add("LineIssue");

           ExcelChart chart = worksheet.Drawings.AddChart("LineChart", eChartType.Line);
           
           worksheet.Cells["A1"].Value=1;
           worksheet.Cells["A2"].Value=2;
           worksheet.Cells["A3"].Value=3;
           worksheet.Cells["A4"].Value=4;
           worksheet.Cells["A5"].Value=5;
           worksheet.Cells["A6"].Value=6;

           worksheet.Cells["B1"].Value=10000;
           worksheet.Cells["B2"].Value=10100;
           worksheet.Cells["B3"].Value=10200;
           worksheet.Cells["B4"].Value=10150;
           worksheet.Cells["B5"].Value=10250;
           worksheet.Cells["B6"].Value=10200;

           chart.Series.Add(ExcelRange.GetAddress(1, 2, worksheet.Dimension.End.Row, 2),
                            ExcelRange.GetAddress(1, 1, worksheet.Dimension.End.Row, 1));

           var Series = chart.Series[0];
           chart.Axis[0].MinorGridlines.Fill.Color = Color.Red;
           chart.Axis[0].MinorGridlines.LineStyle = eLineStyle.LongDashDot;

           chart.Series[0].Header = "Blah";
        }
        [TestMethod]
        public void MultiChartSeries()
        {
            ExcelWorksheet worksheet = _pck.Workbook.Worksheets.Add("MultiChartTypes");

            ExcelChart chart = worksheet.Drawings.AddChart("chtPie", eChartType.LineMarkers);
            chart.SetPosition(100, 100);
            chart.SetSize(800,600);
            AddTestSerie(worksheet, chart);
            chart.Series[0].Header = "Serie5";
            chart.Style = eChartStyle.Style27;
            worksheet.Cells["W19"].Value = 120;
            worksheet.Cells["W20"].Value = 122;
            worksheet.Cells["W21"].Value = 121;
            worksheet.Cells["W22"].Value = 123;
            worksheet.Cells["W23"].Value = 125;
            worksheet.Cells["W24"].Value = 124;

            worksheet.Cells["X19"].Value = 90;
            worksheet.Cells["X20"].Value = 52;
            worksheet.Cells["X21"].Value = 88;
            worksheet.Cells["X22"].Value = 75;
            worksheet.Cells["X23"].Value = 77;
            worksheet.Cells["X24"].Value = 99;
            
            var cs2 = chart.PlotArea.ChartTypes.Add(eChartType.ColumnClustered);
            var s = cs2.Series.Add(worksheet.Cells["W19:W24"], worksheet.Cells["U19:U24"]);
            s.Header = "Serie4";
            cs2.YAxis.MaxValue = 300;
            cs2.YAxis.MinValue = -5.5;
            var cs3 = chart.PlotArea.ChartTypes.Add(eChartType.Line);
            s=cs3.Series.Add(worksheet.Cells["X19:X24"], worksheet.Cells["U19:U24"]);
            s.Header = "Serie1";
            cs3.UseSecondaryAxis = true;
                        
            cs3.XAxis.Deleted = false;
            cs3.XAxis.MajorUnit = 20;
            cs3.XAxis.MinorUnit = 3;
            cs3.XAxis.MinorUnit = null;

            cs3.XAxis.TickLabelPosition = eTickLabelPosition.High;
            cs3.YAxis.LogBase = 10.2;

            var chart2 = worksheet.Drawings.AddChart("scatter1", eChartType.XYScatterSmooth);
            s=chart2.Series.Add(worksheet.Cells["W19:W24"], worksheet.Cells["U19:U24"]);
            s.Header = "Serie2";

            var c2ct2 = chart2.PlotArea.ChartTypes.Add(eChartType.XYScatterSmooth);
            s=c2ct2.Series.Add(worksheet.Cells["X19:X24"], worksheet.Cells["V19:V24"]);
            s.Header="Serie3";
            s=c2ct2.Series.Add(worksheet.Cells["W19:W24"], worksheet.Cells["V19:V24"]);
            s.Header = "Serie4";

            c2ct2.UseSecondaryAxis = true;
            c2ct2.XAxis.Deleted = false;
            c2ct2.XAxis.TickLabelPosition = eTickLabelPosition.High;

            ExcelChart chart3 = worksheet.Drawings.AddChart("chart", eChartType.LineMarkers);
            chart3.SetPosition(300, 1000);
            var s31=chart3.Series.Add(worksheet.Cells["W19:W24"], worksheet.Cells["U19:U24"]);
            s31.Header = "Serie1";

            var c3ct2 = chart3.PlotArea.ChartTypes.Add(eChartType.LineMarkers);
            var c32 = c3ct2.Series.Add(worksheet.Cells["X19:X24"], worksheet.Cells["V19:V24"]);
            c3ct2.UseSecondaryAxis = true;
            c32.Header = "Serie2";
            
            XmlNamespaceManager ns=new XmlNamespaceManager(new NameTable());
            ns.AddNamespace("c","http://schemas.openxmlformats.org/drawingml/2006/chart");
            var element = chart.ChartXml.SelectSingleNode("//c:plotVisOnly", ns);
            if (element!=null) element.ParentNode.RemoveChild(element);
        }
        [TestMethod]
        public void DeleteDrawing()
        {
            var ws=_pck.Workbook.Worksheets.Add("DeleteDrawing1");
            var chart1 = ws.Drawings.AddChart("Chart1", eChartType.Line);
            var chart2 = ws.Drawings.AddChart("Chart2", eChartType.Line);
            var shape1 = ws.Drawings.AddShape("Shape1", eShapeStyle.ActionButtonBackPrevious);
            var pic1 = ws.Drawings.AddPicture("Pic1", Resources.Test1);
            ws.Drawings.Remove(2);
            ws.Drawings.Remove(chart2);
            ws.Drawings.Remove("Pic1");

            ws = _pck.Workbook.Worksheets.Add("DeleteDrawing2");
            chart1 = ws.Drawings.AddChart("Chart1", eChartType.Line);
            chart2 = ws.Drawings.AddChart("Chart2", eChartType.Line);
            shape1 = ws.Drawings.AddShape("Shape1", eShapeStyle.ActionButtonBackPrevious);
            pic1 = ws.Drawings.AddPicture("Pic1", Resources.Test1);

            ws.Drawings.Remove("chart1");

            ws = _pck.Workbook.Worksheets.Add("ClearDrawing2");
            chart1 = ws.Drawings.AddChart("Chart1", eChartType.Line);
            chart2 = ws.Drawings.AddChart("Chart2", eChartType.Line);
            shape1 = ws.Drawings.AddShape("Shape1", eShapeStyle.ActionButtonBackPrevious);
            pic1 = ws.Drawings.AddPicture("Pic1", Resources.Test1);
            ws.Drawings.Clear();
        }   
        [TestMethod]
        public void ReadDocument()
        {
            var fi=new FileInfo(_worksheetPath + "drawingread.xlsx");
            if (!fi.Exists)
            {
                Assert.Inconclusive("Drawing.xlsx is not created. Skippng");
            }
            var pck = new ExcelPackage(fi, true);

            foreach(var ws in pck.Workbook.Worksheets)
            {
                foreach(ExcelDrawing d in ws.Drawings)
                {
                    if (d is ExcelChart c)
                    {
#if (!Core)
                        TestContext.WriteLine(c.ChartType.ToString());
#endif
                    }
                }
            }
            pck.Dispose();
        }
        [TestMethod]
        [Ignore]
        public void ReadMultiChartSeries()
        {
            ExcelPackage pck = new ExcelPackage(new FileInfo("c:\\temp\\chartseries.xlsx"), true);

            var ws = pck.Workbook.Worksheets[1];
            ExcelChart c = ws.Drawings[0] as ExcelChart;


            var p = c.PlotArea;
            p.ChartTypes[1].Series[0].Series = "S7:S15";

            var c2=ws.Drawings.AddChart("NewChart", eChartType.ColumnClustered);
            var serie1 = c2.Series.Add("R7:R15", "Q7:Q15");
            c2.SetSize(800, 800);
            serie1.Header = "Column Clustered";

            var subChart = c2.PlotArea.ChartTypes.Add(eChartType.LineMarkers);
            var serie2 = subChart.Series.Add("S7:S15", "Q7:Q15");
            serie2.Header = "Line";

            //var subChart2 = c2.PlotArea.ChartTypes.Add(eChartType.DoughnutExploded);
            //var serie3 = subChart2.Series.Add("S7:S15", "Q7:Q15");
            //serie3.Header = "Doughnut";

            var subChart3 = c2.PlotArea.ChartTypes.Add(eChartType.Area);
            var serie4 = subChart3.Series.Add("R7:R15", "Q7:Q15");
            serie4.Header = "Area";
            subChart3.UseSecondaryAxis = true;

            var serie5 = subChart.Series.Add("R7:R15","Q7:Q15");
            serie5.Header = "Line 2";

            pck.SaveAs(new FileInfo("c:\\temp\\chartseriesnew.xlsx"));
        }
        [TestMethod]
        public void ChartWorksheet()
        {
            var wsChart = _pck.Workbook.Worksheets.AddChart("ChartWorksheet", eChartType.Bubble3DEffect);
            var ws = _pck.Workbook.Worksheets.Add("data");
            AddTestSerie(ws, wsChart.Chart);
            wsChart.Chart.Style = eChartStyle.Style23;
            wsChart.Chart.Title.Text = "Chart worksheet";
            wsChart.Chart.Series[0].Header = "Serie";
        }
        [TestMethod]
        public void ReadChartWorksheet()
        {
            //Setup
            var wsName = "ChartWorksheet";
            var pck = OpenPackage("DrawingRead.xlsx");
            var ws = ((ExcelChartsheet)pck.Workbook.Worksheets[wsName]);
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var chart = ws.Chart;

            Assert.AreEqual(eChartStyle.Style23, chart.Style);
            Assert.AreEqual("Chart worksheet", chart.Title.Text);
            Assert.AreEqual("Serie", chart.Series[0].Header);
        }
        [TestMethod]
        public void TestHeaderaddress()
        {
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Draw");
            var chart = ws.Drawings.AddChart("NewChart1",eChartType.Area) as ExcelChart;
            var ser1 = chart.Series.Add("A1:A2", "B1:B2");
            ser1.HeaderAddress = new ExcelAddress("A1:A2");
            ser1.HeaderAddress = new ExcelAddress("A1:B1");
            ser1.HeaderAddress = new ExcelAddress("A1");
            pck.Dispose();
        }
        public void DrawingRowheightDynamic()
        {
            var ws = _pck.Workbook.Worksheets.Add("PicResize");
            ws.Cells["A1"].Value = "test";
            ws.Cells["A1"].Style.Font.Name = "Symbol";
            ws.Cells["A1"].Style.Font.Size = 39;
            ws.Workbook.Styles.NamedStyles[0].Style.Font.Name = "Symbol";
            ws.Workbook.Styles.NamedStyles[0].Style.Font.Size = 16;
            var pic = ws.Drawings.AddPicture("Pic1", Resources.Test1);
        }
        [TestMethod]
        public void ChangeToOneCellAnchor()
        {
            var ws = _pck.Workbook.Worksheets.Add("PicOneCellAnchor");
            var pic = ws.Drawings.AddPicture("Pic1", Resources.Test1);

            pic.ChangeCellAnchor(eEditAs.OneCell, 600, 500, (int)pic._width, (int)pic._height);
            //AssertPic(pic, 600, 500);
        }
        [TestMethod]
        public void ChangeToAbsoluteAnchor()
        {
            var ws = _pck.Workbook.Worksheets.Add("PicAbsoluteAnchor");
            var pic = ws.Drawings.AddPicture("Pic1", Resources.Test1);

            pic.ChangeCellAnchor(eEditAs.Absolute, 600, 500, (int)pic._width, (int)pic._height);
            //AssertPic(pic, 600, 500);
        }
        [TestMethod]
        public void ChangeToTwoCellAnchor()
        {
            var ws = _pck.Workbook.Worksheets.Add("PicTwoCellAnchor");
            var pic = ws.Drawings.AddPicture("Pic1", Resources.Test1);

            pic.ChangeCellAnchor(eEditAs.OneCell, 600, 500, (int)pic._width, (int)pic._height);
            //AssertPic(pic, 600, 500);
            pic.ChangeCellAnchor(eEditAs.TwoCell, 600, 500, (int)pic._width, (int)pic._height);
            //AssertPic(pic, 600, 500);
        }
        [TestMethod]
        public void ChangeToOneCellAnchorNoPositionAndSize()
        {
            var ws = _pck.Workbook.Worksheets.Add("PicOneCellAnchorNoPosAndSize");
            var pic = ws.Drawings.AddPicture("Pic1", Resources.Test1);

            pic.SetPosition(600, 500);

            //One Cell
            pic.ChangeCellAnchor(eEditAs.OneCell);

            //AssertPic(pic, 600, 500);

            pic.ChangeCellAnchor(eEditAs.TwoCell);

            //AssertPic(pic, 600, 500);

            pic.ChangeCellAnchor(eEditAs.Absolute);

            //AssertPic(pic, 600, 500);
        }

        //private static void AssertPic(ExcelPicture pic, int top, int left)
        //{
        //    Assert.AreEqual(Resources.Test1.Width, pic._width);
        //    Assert.AreEqual(Resources.Test1.Height, pic._height);
        //    Assert.AreEqual(top, pic._top);
        //    Assert.AreEqual(left, pic._left);
        //}

        [TestMethod]
        public void ChangeToAbsoluteAnchorNoPositionAndSize()
        {
            var ws = _pck.Workbook.Worksheets.Add("PicAbsoluteAnchorNoPosChange");
            var pic = ws.Drawings.AddPicture("Pic1", Resources.Test1);

            pic.ChangeCellAnchor(eEditAs.OneCell, 500, 500, (int)pic._width, (int)pic._height);
        }
        [TestMethod]
        public void ChangeToTwoCellAnchorNoPositionAndSize()
        {
            var ws = _pck.Workbook.Worksheets.Add("PicTwoCellAnchorNoPosChange");
            var pic = ws.Drawings.AddPicture("Pic1", Resources.Test1);

            pic.ChangeCellAnchor(eEditAs.OneCell, 500, 500, (int)pic._width, (int)pic._height);
            pic.ChangeCellAnchor(eEditAs.TwoCell, 500, 500, (int)pic._width, (int)pic._height);
        }
        [TestMethod]
        public void ValidateTextBody()
        {
            var ws = _pck.Workbook.Worksheets.Add("TextBody_RightInsert");
            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.TextBody.RightInsert = 1;
            
            Assert.AreEqual(1, shape.TextBody.RightInsert);
            shape.ChangeCellAnchor(eEditAs.OneCell);

            Assert.AreEqual(1, shape.TextBody.RightInsert);
        }

        [TestMethod]
        public void SendToBack()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("SendToBack_Shape3");
            var shape1 = ws.Drawings.AddShape("First",eShapeStyle.Rect);
            shape1.Text = "First";
            shape1.SetPosition(1, 0, 1, 0);
            var shape2 = ws.Drawings.AddShape("Second", eShapeStyle.Rect);
            shape2.Text = "Second";
            shape2.SetPosition(2, 0, 2, 0);
            var shape3 = ws.Drawings.AddShape("Third", eShapeStyle.Rect);
            shape3.SetPosition(3, 0, 3, 0);
            shape3.Text = "Third";
            var shape4 = ws.Drawings.AddShape("Fourth", eShapeStyle.Rect);
            shape4.SetPosition(4, 0, 4, 0);
            shape4.Text = "Fourth";

            //Act
            shape3.SendToBack();

            //Assert
            Assert.AreEqual("Third", ws.Drawings[0].Name);
            Assert.AreEqual("First", ws.Drawings[1].Name);
            Assert.AreEqual("Second", ws.Drawings[2].Name);
            Assert.AreEqual("Fourth", ws.Drawings[3].Name);

            Assert.AreEqual("First", ws.Drawings["First"].Name);
            Assert.AreEqual("Third", ws.Drawings["Third"].Name);
            Assert.AreEqual("Second", ws.Drawings["Second"].Name);
            Assert.AreEqual("Fourth", ws.Drawings["Fourth"].Name);
        }
        [TestMethod]
        public void BringToFront()
        {
            //Setup
            var ws = _pck.Workbook.Worksheets.Add("BringToFront_Shape2");
            var shape1 = ws.Drawings.AddShape("First", eShapeStyle.Rect);
            shape1.Text = "First";
            shape1.SetPosition(1, 0, 1, 0);
            var shape2 = ws.Drawings.AddShape("Second", eShapeStyle.Rect);
            shape2.Text = "Second";
            shape2.SetPosition(2, 0, 2, 0);
            var shape3 = ws.Drawings.AddShape("Third", eShapeStyle.Rect);
            shape3.SetPosition(3, 0, 3, 0);
            var shape4 = ws.Drawings.AddShape("Fourth", eShapeStyle.Rect);
            shape4.SetPosition(4, 0, 4, 0);
            shape4.Text = "Fourth";

            //Act
            shape2.BringToFront();

            //Assert
            Assert.AreEqual("First", ws.Drawings[0].Name);
            Assert.AreEqual("Third", ws.Drawings[1].Name);
            Assert.AreEqual("Fourth", ws.Drawings[2].Name);
            Assert.AreEqual("Second", ws.Drawings[3].Name);

            Assert.AreEqual("First", ws.Drawings["First"].Name);
            Assert.AreEqual("Second", ws.Drawings["Second"].Name);
            Assert.AreEqual("Third", ws.Drawings["Third"].Name);
            Assert.AreEqual("Fourth", ws.Drawings["Fourth"].Name);
        }
        [TestMethod]
        public void ReadControls()
        {
            using (var p = OpenTemplatePackage("Controls.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                Assert.AreEqual(3, ws.Drawings.Count);

            }
        }
        [TestMethod]
        public void DrawingSetFont()
        {
            var ws = _pck.Workbook.Worksheets.Add("DrawingChangeFont");
            var shape = ws.Drawings.AddShape("FontChange", eShapeStyle.Rect);
            shape.Font.SetFromFont("Arial", 20);
            shape.Text = "Font";
            shape.RichText[0].SetFromFont("Calibri", 8);  //works
            shape.RichText.Add("New Line", true);

            Assert.AreEqual("Arial", shape.Font.LatinFont);
            Assert.AreEqual("Arial", shape.Font.ComplexFont);
            Assert.AreEqual(20 , shape.Font.Size);

            Assert.AreEqual("Calibri", shape.RichText[0].LatinFont);
            Assert.AreEqual("Calibri", shape.RichText[0].ComplexFont);
            Assert.AreEqual(8, shape.RichText[0].Size);
        }
        [TestMethod]
        public void PictureChangeCellAnchor()
        {
            var ws = _pck.Workbook.Worksheets.Add("PictureChangeCellAnchore");
            var pic = ws.Drawings.AddPicture("Pic1", Resources.Test1);
            pic.ChangeCellAnchor(eEditAs.TwoCell);
        }
    }
}

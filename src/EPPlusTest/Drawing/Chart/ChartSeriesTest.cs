using Castle.Components.DictionaryAdapter.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace EPPlusTest.Drawing.Chart
{
    [TestClass]
    public class ChartSeriesTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("ChartSingleSerie.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        #region Single Serie
        [TestMethod]
        public void AddSunburstChartSingleSerie()
        {
            var ws = _pck.Workbook.Worksheets.Add("Sunburst");
            LoadHierarkiTestData(ws);
            var chart = ws.Drawings.AddSunburstChart("Sunburst1");
            var serie = chart.Series.Add(ws.Cells["D2:D17"]);
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);
            serie.DataLabel.Position = eLabelPosition.Center;
            serie.DataLabel.ShowCategory = true;
            serie.DataLabel.ShowValue = true;
            var dp = serie.DataPoints.Add(2);

            Assert.AreEqual(eDrawingType.Chart, chart.DrawingType);
            Assert.IsInstanceOfType(chart, typeof(ExcelSunburstChart));
            Assert.AreEqual(0, chart.Axis.Length);
            Assert.IsNull(chart.XAxis);
            Assert.IsNull(chart.YAxis);

        }
        [TestMethod]
        public void AddPieChartSingleSerie()
        {
            var ws = _pck.Workbook.Worksheets.Add("Pie");
            LoadHierarkiTestData(ws);
            var chart = ws.Drawings.AddPieChart("Pie1", ePieChartType.Pie);
            var serie = chart.Series.Add(ws.Cells["D2:D17"]);
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);
            serie.DataLabel.Position = eLabelPosition.Center;
            serie.DataLabel.ShowCategory = true;
            serie.DataLabel.ShowValue = true;
            var dp = serie.DataPoints.Add(2);

            Assert.AreEqual(eDrawingType.Chart, chart.DrawingType);
            Assert.IsInstanceOfType(chart, typeof(ExcelPieChart));
            Assert.AreEqual(0, chart.Axis.Length);
            Assert.IsNull(chart.XAxis);
            Assert.IsNull(chart.YAxis);

        }
        [TestMethod]
        public void AddColumnChartSingleSerieWithSecondSerieWithCategory()
        {
            var ws = _pck.Workbook.Worksheets.Add("Column");
            LoadHierarkiTestData(ws);
            var chart = ws.Drawings.AddBarChart("Bar1", eBarChartType.Column3D);
            var serie1 = chart.Series.Add(ws.Cells["D2:D17"]);
            var serie2 = chart.Series.Add(ws.Cells["D2:D17"], ws.Cells["C2:C17"]);
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);

            Assert.AreEqual(ws.Cells["D2:D17"].FullAddressAbsolute, serie1.Series);
            Assert.AreEqual("", serie1.XSeries);
            Assert.AreEqual(ws.Cells["D2:D17"].FullAddressAbsolute, serie2.Series);
            Assert.AreEqual(ws.Cells["C2:C17"].FullAddressAbsolute, serie2.XSeries);

            Assert.AreEqual(eDrawingType.Chart, chart.DrawingType);
            Assert.IsInstanceOfType(chart, typeof(ExcelBarChart));
            Assert.AreEqual(2, chart.Axis.Length);
            Assert.IsNotNull(chart.XAxis);
            Assert.IsNotNull(chart.YAxis);

        }

        [TestMethod]
        public void AddColumnChartSingleSerieWithSecondSerieWithCategoryWithLinear()
        {
            var ws = _pck.Workbook.Worksheets.Add("ColumnWithinLinear");
            LoadHierarkiTestData(ws);
            var chart = ws.Drawings.AddBarChart("Bar1", eBarChartType.Column3D);

            //Change chart colorMethod from Cylce to WithinLinear
            chart.StyleManager.SetChartStyle(OfficeOpenXml.Drawing.Chart.Style.ePresetChartStyleMultiSeries.Column3dChartStyle1,
                OfficeOpenXml.Drawing.Chart.Style.ePresetChartColors.ColorfulPalette1);
            chart.StyleManager.ColorsManager.Method = OfficeOpenXml.Drawing.Chart.Style.eChartColorStyleMethod.WithinLinear;

            //make series only have range of 1 so that the serie2(index=1) is the same as the number of cells in the range
            //which causes System.ArgumentException: Negative percentage not allowed
            var serie1 = chart.Series.Add(ws.Cells["D2"]);
            var serie2 = chart.Series.Add(ws.Cells["D2"], ws.Cells["C2"]);
        }

        [TestMethod]
        public void AddChartWithLegendEntries()
        {
            var ws = _pck.Workbook.Worksheets.Add("LegendEntries");
            LoadHierarkiTestData(ws);
            var chart = ws.Drawings.AddBarChart("Bar1", eBarChartType.Column3D);
            var serie1 = chart.Series.Add(ws.Cells["D2:D17"]);
            var serie2 = chart.Series.Add(ws.Cells["D2:D17"], ws.Cells["C2:C17"]);
            var serie3 = chart.Series.Add(ws.Cells["D2:D17"], ws.Cells["B2:B17"]);

            serie1.Header = "Serie 1";
            serie2.Header = "Serie 2-Deleted";
            serie3.Header = "Serie 3-Font Changed";

            chart.Legend.Entries[2].Font.Fill.Style = eFillStyle.SolidFill;
            chart.Legend.Entries[2].Font.Fill.SolidFill.Color.SetRgbColor(Color.Red);

            chart.Legend.Entries[1].Deleted = true;

            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);

            Assert.AreEqual(ws.Cells["D2:D17"].FullAddressAbsolute, serie1.Series);
            Assert.AreEqual("", serie1.XSeries);
            Assert.AreEqual(ws.Cells["D2:D17"].FullAddressAbsolute, serie2.Series);
            Assert.AreEqual(ws.Cells["C2:C17"].FullAddressAbsolute, serie2.XSeries);

            Assert.AreEqual("Serie 1", serie1.Header);
            Assert.AreEqual("Serie 2-Deleted", serie2.Header);
            Assert.IsTrue(chart.Legend.Entries[1].Deleted);
            Assert.AreEqual("Serie 3-Font Changed", serie3.Header);

            Assert.AreEqual(eFillStyle.SolidFill,chart.Legend.Entries[2].Font.Fill.Style);
            Assert.AreEqual(Color.Red.ToArgb(), chart.Legend.Entries[2].Font.Fill.SolidFill.Color.RgbColor.Color.ToArgb());
        }

        [TestMethod]
        public void ReadChartWithLegendEntries()
        {
            using (var p1 = new ExcelPackage())
            {
                var ws = p1.Workbook.Worksheets.Add("LegendEntries");
                LoadHierarkiTestData(ws);
                var chart = ws.Drawings.AddBarChart("Bar1", eBarChartType.Column3D);
                var serie1 = chart.Series.Add(ws.Cells["D2:D17"]);
                var serie2 = chart.Series.Add(ws.Cells["D2:D17"], ws.Cells["C2:C17"]);
                var serie3 = chart.Series.Add(ws.Cells["D2:D17"], ws.Cells["B2:B17"]);

                serie1.Header = "Serie 1";
                serie2.Header = "Serie 2-Deleted";
                serie3.Header = "Serie 3-Font Changed";

                chart.Legend.Entries[2].Font.Fill.Style = eFillStyle.SolidFill;
                chart.Legend.Entries[2].Font.Fill.SolidFill.Color.SetRgbColor(Color.Red);

                chart.Legend.Entries[1].Deleted = true;

                chart.SetPosition(2, 0, 15, 0);
                chart.SetSize(1600, 900);

                //Assert p1
                Assert.AreEqual(ws.Cells["D2:D17"].FullAddressAbsolute, serie1.Series);
                Assert.AreEqual("", serie1.XSeries);
                Assert.AreEqual(ws.Cells["D2:D17"].FullAddressAbsolute, serie2.Series);
                Assert.AreEqual(ws.Cells["C2:C17"].FullAddressAbsolute, serie2.XSeries);

                Assert.AreEqual("Serie 1", serie1.Header);
                Assert.AreEqual("Serie 2-Deleted", serie2.Header);
                Assert.IsTrue(chart.Legend.Entries[1].Deleted);
                Assert.AreEqual("Serie 3-Font Changed", serie3.Header);

                Assert.AreEqual(eFillStyle.SolidFill, chart.Legend.Entries[2].Font.Fill.Style);
                Assert.AreEqual(Color.Red.ToArgb(), chart.Legend.Entries[2].Font.Fill.SolidFill.Color.RgbColor.Color.ToArgb());

                p1.Save();

                using (var p2 = new ExcelPackage(p1.Stream))
                {
                    ws = p2.Workbook.Worksheets[0];
                    chart = ws.Drawings[0].As.Chart.BarChart;
                    serie1 = chart.Series[0];
                    serie2 = chart.Series[1];
                    serie3 = chart.Series[2];

                    //Assert p2 
                    Assert.AreEqual(ws.Cells["D2:D17"].FullAddressAbsolute, serie1.Series);
                    Assert.AreEqual("", serie1.XSeries);
                    Assert.AreEqual(ws.Cells["D2:D17"].FullAddressAbsolute, serie2.Series);
                    Assert.AreEqual(ws.Cells["C2:C17"].FullAddressAbsolute, serie2.XSeries);

                    Assert.AreEqual("Serie 1", serie1.Header);
                    Assert.AreEqual("Serie 2-Deleted", serie2.Header);
                    Assert.IsTrue(chart.Legend.Entries[1].Deleted);
                    Assert.AreEqual("Serie 3-Font Changed", serie3.Header);

                    Assert.IsFalse(chart.Legend.Entries[0].Deleted);
                    Assert.AreEqual(eFillStyle.SolidFill, chart.Legend.Entries[2].Font.Fill.Style);
                    Assert.AreEqual(Color.Red.ToArgb(), chart.Legend.Entries[2].Font.Fill.SolidFill.Color.RgbColor.Color.ToArgb());
                }
            }
        }
        #endregion

        [TestMethod]
        public void SimpleChartDataLabels()
        {
            using (var p = OpenPackage("ChartDataLabels.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("newWs");
                var table = ws.Tables.Add(ws.Cells["A1:K20"], "TestTable");

                var dataRange = ws.Cells["A2:K20"];
                dataRange.Formula = "ROW() + COLUMN()";

                ws.Calculate();

                var chart = ws.Drawings.AddBarChart("barChart", eBarChartType.ColumnClustered);

                chart.ShowDataLabelsOverMaximum = false;

                for (int i = 1; i < table.Columns.Count; i++)
                {
                    var series = chart.Series.Add(dataRange.TakeSingleColumn(i), dataRange.TakeRowsBetween(0, dataRange.Rows));
                    series.DataLabel.ShowValue = true;
                    series.DataLabel.DataLabels.Add(0).ShowValue = true;
                }

                SaveAndCleanup(p);
            }
        }

        //[TestMethod]
        //public void GenerateExample()
        //{
        //    using (var package = OpenPackage("manualLayoutChartCommentRange.xlsx", true))
        //    {
        //        var ws = package.Workbook.Worksheets.Add("ManualLayout");

        //        //Create some values
        //        ws.Cells["A1:A2"].Value = 5;
        //        ws.Cells["B1:B2"].Value = 10;

        //        //Create a column chart
        //        var sChart = ws.Drawings.AddBarChart("ColumnChart", eBarChartType.ColumnClustered);

        //        //Add series (clustered columns) to the chart. In this case 2 per series
        //        var s1 = sChart.Series.Add(ws.Cells["A1:A2"]);
        //        var s2 = sChart.Series.Add(ws.Cells["B1:B2"]);

        //        //Add a general datalabel
        //        var label = s1.DataLabel;
        //        label.ShowValue = true;

        //        //Add a specific datalabel to the first column in the cluster
        //        var dl = label.DataLabels.Add(0);

        //        var myLabel = s1.DataLabel.DataLabels[0];

        //        //Offset the data label 10% of the charts width to the left
        //        //AKA Remove 10 from x coordinate
        //        dl.Layout.ManualLayout.Left = -10;

        //        //Offset the data label 10% of the charts height to the top
        //        //AKA remove 10 from y coordinate
        //        dl.Layout.ManualLayout.Top = -10;

        //        //Save the package at a path
        //        package.SaveAs(@"C:\temp\manualLayoutChart.xlsx");
        //    }
        //}



        [TestMethod]
        public void ReadSimpleFile()
        {
            using (var package = OpenTemplatePackage("editedDataLabel.xlsx"))
            {
                var ws = package.Workbook.Worksheets[0];

                var myChart = ws.Drawings[0].As.Chart.BarChart;

                var lbl = myChart.Series[0].DataLabel.DataLabels[0];

                var lblTxtBody = myChart.Series[0].DataLabel.DataLabels[0].TextBody;

                

                //var serializedParent = JsonConvert.SerializeObject(test.Font);
                //ExcelParagraph castedFont = JsonConvert.DeserializeObject<ExcelParagraph>(serializedParent);


                //castedFont.Text = "My Overriding Comment";

                //for (int i = 1; i < 4; i++)
                //{
                //    ws.Cells[i, 2].Value = $"Comment {i}";
                //}

                //var serie = myChart.Series.Add(ws.Cells["A1:A3"], ws.Cells["C1:C3"]);
                //serie.DataLabel.ShowValue = true;

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void ReadFile()
        {
            using (var package = OpenTemplatePackage("S1008_NoComment.xlsx"))
            {
                var ws = package.Workbook.Worksheets[0];

                var chart = ws.Drawings[0].As.Chart.LineChart;

                //var label = chart.Series[0].DataLabel.DataLabels[20];

                //var paraStr = label.GetExistingParagraphStrings();
                //label.OverWriteText("Mislabeled resistors - wrong part defects");

                //var paraStr2 = label.GetExistingParagraphStrings();

                chart.Series[0].DataLabel.SelectRange(ws.Cells["E2:E53"]);

                //chart.Series[0]

                //chart.Series[0].XSeries = $"{{'Test{1}','Test{2}'}}";

                //string myComment = "Mislabeled resistors - wrong part defects";

                //CommentOnDataLabel(myChart, 0,20, myComment);


                //Save workbook

                //ready namespace manager
                //var nsm = new XmlNamespaceManager(myChart.NameTable);
                //nsm.AddNamespace(@"c", @"http://schemas.openxmlformats.org/drawingml/2006/chart");
                //nsm.AddNamespace(@"a", @"http://schemas.openxmlformats.org/drawingml/2006/main");

                //int serieIdx = 0;

                //var serieNode = myChart.ChildNodes[2].SelectSingleNode(@"c:chart/c:plotArea/c:lineChart/c:ser[c:idx[@val='" + serieIdx + "']]", nsm);
                //int dlblIdx = 20;
                //var dataLabelToComment = serieNode.SelectSingleNode(@"c:dLbls/c:dLbl[c:idx[@val='" + dlblIdx + "']]", nsm);


                //dataLabelToComment.InnerXml = $"<a:r><a:rPr lang=\"en-US\"/><a:t>{myComment}</a:t></a:r>";

                //"Mislabeled resistors - wrong part defects";
                //var firstSeries = myChart.SelectSingleNode("//c:ser", myChart.namespa);
                //var idxNode = firstSeries.SelectSingleNode("//c:idx[@val=\"21\"]");

                //firstSeries.SelectSingleNode()

                //var theSeries = chart.Series[0].DataPoints
                //aNode.
                //var myLitterals = chart.Series[0].StringLiteralsX;
                //var myLitteralsOther = chart.Series[0].StringLiteralsY;

                //chart.Series[0].CreateCache();

                //var dataPoint = chart.Series[0].DataPoints[21];

                //var element = label21.TextBody.PathElement;
                //label26.
                //var fldParagraph = label21.TextBody.Paragraphs[0];
                //fldParagraph.Text

                SaveAndCleanup(package);
            }
        }

        static void CommentOnDataLabel(XmlDocument chartXml, int seriesIdx, int datalabelIdx, string myComment)
        {
            var nsm = new XmlNamespaceManager(chartXml.NameTable);
            nsm.AddNamespace(@"c", @"http://schemas.openxmlformats.org/drawingml/2006/chart");
            nsm.AddNamespace(@"a", @"http://schemas.openxmlformats.org/drawingml/2006/main");

            int serieIdx = 0;

            var serieNode = chartXml.ChildNodes[2].SelectSingleNode(@"c:chart/c:plotArea/c:lineChart/c:ser[c:idx[@val='" + serieIdx + "']]", nsm);
            
            var dataLabelToComment = serieNode.SelectSingleNode(@"c:dLbls/c:dLbl[c:idx[@val='" + datalabelIdx + "']]", nsm);

            var paragraphNodeForTheText = dataLabelToComment.SelectSingleNode("c:tx/c:rich/a:p", nsm);
            paragraphNodeForTheText.InnerXml = $"<a:r><a:rPr lang=\"en-US\"/><a:t>{myComment}</a:t></a:r>";
        }
    }
}

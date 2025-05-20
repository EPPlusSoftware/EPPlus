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
using Microsoft.VisualStudio.TestPlatform.PlatformAbstractions.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace EPPlusTest.Drawing.Chart
{
    [TestClass]
    public class ExcelChartAxisTest : TestBase
    {
        private ExcelChartAxis axis;

        [TestInitialize]
        public void Initialize()
        {
            var xmlDoc = new XmlDocument();
            xmlDoc.LoadXml("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ></c:chartSpace>");
            var xmlNsm = new XmlNamespaceManager(new NameTable());
            xmlNsm.AddNamespace("c", ExcelPackage.schemaChart);
            xmlNsm.AddNamespace("a", ExcelPackage.schemaDrawings);
            var node = xmlDoc.CreateElement("axis");
            xmlDoc.DocumentElement.AppendChild(node);
            axis = new ExcelChartAxisStandard(null, xmlNsm, node, "c");
        }

        [TestMethod]
        public void CrossesAt_SetTo2_Is2()
        {
            axis.CrossesAt = 2;
            Assert.AreEqual(axis.CrossesAt, 2);
        }

        [TestMethod]
        public void CrossesAt_SetTo1EMinus6_Is1EMinus6()
        {
            axis.CrossesAt = 1.2e-6;
            Assert.AreEqual(axis.CrossesAt, 1.2e-6);
        }

        [TestMethod]
        public void MinValue_SetTo2_Is2()
        {
            axis.MinValue = 2;
            Assert.AreEqual(axis.MinValue, 2);
        }

        [TestMethod]
        public void MinValue_SetTo1EMinus6_Is1EMinus6()
        {
            axis.MinValue = 1.2e-6;
            Assert.AreEqual(axis.MinValue, 1.2e-6);
        }

        [TestMethod]
        public void MaxValue_SetTo2_Is2()
        {
            axis.MaxValue = 2;
            Assert.AreEqual(axis.MaxValue, 2);
        }

        [TestMethod]
        public void MaxValue_SetTo1EMinus6_Is1EMinus6()
        {
            axis.MaxValue = 1.2e-6;
            Assert.AreEqual(axis.MaxValue, 1.2e-6);
        }
        [TestMethod]
        public void Gridlines_Set_IsNotNull()
        {
            var major = axis.MajorGridlines;
            major.Width = 1;
            Assert.IsTrue(axis.ExistsNode("c:majorGridlines"));

            var minor = axis.MinorGridlines;
            minor.Width = 1;
            Assert.IsTrue(axis.ExistsNode("c:minorGridlines"));
        }

        [TestMethod]
        public void Gridlines_Remove_IsNull()
        {
            var major = axis.MajorGridlines;
            major.Width = 1;
            var minor = axis.MinorGridlines;
            minor.Width = 1;

            axis.RemoveGridlines();

            Assert.IsFalse(axis.ExistsNode("c:majorGridlines"));
            Assert.IsFalse(axis.ExistsNode("c:minorGridlines"));

            major = axis.MajorGridlines;
            major.Width = 1;
            minor = axis.MinorGridlines;
            minor.Width = 1;

            axis.RemoveGridlines(true, false);

            Assert.IsFalse(axis.ExistsNode("c:majorGridlines"));
            Assert.IsTrue(axis.ExistsNode("c:minorGridlines"));

            major = axis.MajorGridlines;
            major.Width = 1;
            minor = axis.MinorGridlines;
            minor.Width = 1;

            axis.RemoveGridlines(false, true);

            Assert.IsTrue(axis.ExistsNode("c:majorGridlines"));
            Assert.IsFalse(axis.ExistsNode("c:minorGridlines"));
        }

        [TestMethod]
        public void LoadChartTest()
        {
            using (var p = OpenTemplatePackage("testAxisLabels.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];

                var chart1 = p.Workbook.Worksheets[0].Drawings[0].As.Chart.BarChart;
                var ser4 = chart1.Series[4];
                var someSeries = chart1.Series[4].XSeries;

                var strLits = chart1.Series[4].StringLiteralsX;

                chart1.YAxis.ChangeAxisTypeReal(eAxisType.Date);

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void TestAxisEpplus()
        {
            using (var p = OpenPackage("EpplusAxisCase1.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = "Col1";
                ws.Cells["B1"].Value = "Col2";

                var datRange = ws.Cells["A2:B3"];

                datRange.Formula = "ROW() + COLUMN()";

                ws.Calculate();

                var barChart = ws.Drawings.AddBarChart("testChart", eBarChartType.ColumnClustered);

                barChart.Series.Add(datRange.TakeSingleColumn(0), ws.Cells["A1:B1"]);

                ws.Calculate();

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void TestAxisEpplus2()
        {
            using (var p = OpenPackage("EpplusAxisCase2.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var dt = DateTime.Now;
                ws.Cells["A1"].Value = dt;
                ws.Cells["B1"].Value = dt.AddDays(-1);

                ws.Cells["A1:B1"].Style.Numberformat.Format = "d-mmm";

                var datRange = ws.Cells["A2:B3"];

                datRange.Formula = "ROW() + COLUMN()";

                ws.Calculate();

                var barChart = ws.Drawings.AddBarChart("testChart", eBarChartType.ColumnClustered);

                var ser = barChart.Series.Add(ws.Cells["A2:B2"], ws.Cells["A1:B1"]);
                ser.Header = "Ser1";
                var ser2 = barChart.Series.Add(ws.Cells["A3:B3"], ws.Cells["A1:B1"]);
                ser2.Header = "Ser2";

                barChart.XAxis.ChangeAxisTypeReal(eAxisType.Date);

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void SecondaryAxis()
        {
            using (var p = OpenPackage("doubleAxis.xlsx", true))
            {
                var Worksheet = p.Workbook.Worksheets.Add("Sheet1");

                ExcelChart chart = Worksheet.Drawings.AddLineChart("chtLine", eLineChartType.LineMarkers);
                var serie1 = chart.Series.Add(Worksheet.Cells["B1:B4"], Worksheet.Cells["A1:A4"]);
                var chartType2 = chart.PlotArea.ChartTypes.AddLineChart(eLineChartType.LineMarkers);
                var serie2 = chartType2.Series.Add(Worksheet.Cells["C1:C4"], Worksheet.Cells["A1:A4"]);
                chartType2.UseSecondaryAxis = true;

                //By default the secondary X axis is hidden. If you what to show it, try this...
                chartType2.XAxis.Deleted = false;
                chartType2.XAxis.TickLabelPosition = eTickLabelPosition.High;

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void TestLineChartSimpleChange()
        {
            using (var p = OpenPackage("EpplusLineChartSimple.xlsx", true))
            {
                var now = DateTime.Now;

                var ws = p.Workbook.Worksheets.Add("Sheet1");

                //var headers = new List<string> { "Dates", "Sales total", "Spending Total", "Day total"};

                //for (int i = 0; i < headers.Count; i++)
                //{
                //    ws.Cells[1, i + 1].Value = headers[i];
                //}

                var dateRange = ws.Cells["$A$2:$A$4"];

                var dList = new List<DateTime> { now, now.AddDays(1), now.AddDays(2) };

                var retRange = ws.Cells["A1"].LoadFromCollection(dList);

                //var headers = new List<string> { "Item Type", "Date Sold", "Amount Sold", "Shop Name" };

                //var headRange = ws.Cells["A1:D1"];

                //var retRange = headRange.LoadFromCollection(headers);

                var testRetRange = dateRange.LoadFromCollection(new List<DateTime> { now, now.AddDays(1), now.AddDays(2) });


                //var salesRange = ws.Cells["B2:B4"].LoadFromCollection(new List<double> { 0d, 500d, 1500d });
                //var spendRange = ws.Cells["C2:C4"].LoadFromCollection(new List<double> { 200d, 10d, 400d });
                //var totalRange = ws.Cells["D2:D4"];
                //totalRange.Formula = "B2-C2";
                //totalRange.Calculate();

                //ws.Cells["B2:D4"].Style.Numberformat.Format = "#,##0kr";
                //ws.Cells["A2:A4"].Style.Numberformat.Format = "dd/mm/yyyy";

                //var lineChart = ws.Drawings.AddLineChart("testLineChart", eLineChartType.Line);

                //var ser1 = lineChart.Series.Add(salesRange, dateRange);
                //ser1.HeaderAddress = ws.Cells["B1"];
                //var ser2 = lineChart.Series.Add(spendRange, dateRange);
                //ser2.HeaderAddress = ws.Cells["C1"];
                //var ser3 = lineChart.Series.Add(totalRange, dateRange);
                //ser3.HeaderAddress = ws.Cells["D1"];
                ////lineChart.XAxis.ChangeAxisTypeReal(eAxisType.Date);

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void TestBarChartDateOnMulti()
        {
            using (var p = OpenPackage("EpplusAmountSold.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");

                var now = DateTime.Now;

                var headers = new List<string> { "Item Type", "Date Sold", "Amount Sold", "Shop Name" };

                for(int i = 0; i< headers.Count; i++)
                {
                    ws.Cells[1, i + 1].Value = headers[i];
                }

                var headerRange = ws.Cells["A1:D1"];
                var itemRange = ws.Cells["A2:A4"].LoadFromCollection(new List<string> { "Hammer", "Tongs", "Sickle" });
                var dateRange = ws.Cells["B2:B4"].LoadFromCollection(new List<DateTime> { now, now.AddDays(-1), now.AddDays(-2) });
                var amountRange = ws.Cells["C2:C4"].LoadFromCollection(new List<int> { 1, 2, 3 });
                var shopRange = ws.Cells["D2:D4"].LoadFromCollection(new List<string> { "Hammer Shop", "Tong Shop", "Sickle Shop" });

                ws.Cells["B2:B4"].Style.Numberformat.Format = "d-mmm";

                var barChart = ws.Drawings.AddBarChart("testChart", eBarChartType.BarClustered);

                //Multi-Str ref since multiple rows/Cols "A2:B4"
                var serAmount = barChart.Series.Add(amountRange, ws.Cells["A2:B4"]);
                serAmount.HeaderAddress = ws.Cells["A1"];

                //Changing to date type is still allowed.
                //TODO: Why is XAxis actually the YAxis visually?
                barChart.XAxis.ChangeAxisTypeReal(eAxisType.Date);

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void TestAxisEpplusChangeToDate()
        {
            using (var p = OpenPackage("EpplusAxisCaseChangeToDate.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = "Col1";
                ws.Cells["B1"].Value = "Col2";

                var datRange = ws.Cells["A2:B3"];

                datRange.Formula = "ROW() + COLUMN()";

                ws.Calculate();

                var barChart = ws.Drawings.AddBarChart("testChart", eBarChartType.ColumnClustered);

                barChart.Series.Add(datRange.TakeSingleColumn(0), ws.Cells["A1:B1"]);

                ws.Calculate();

                var dt = DateTime.Now;
                ws.Cells["A1"].Value = dt;
                ws.Cells["B1"].Value = dt.AddDays(-1);

                ws.Cells["A1:B1"].Style.Numberformat.Format = "d-mmm";

                ws.Calculate();

                barChart.YAxis.ChangeAxisTypeReal(eAxisType.Cat);
                barChart.XAxis.ChangeAxisTypeReal(eAxisType.Val);

                barChart.Series.Delete(0);
                barChart.Series.Add(datRange.TakeSingleColumn(0), ws.Cells["A1:B1"]);

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void ChangingAxisOfChart()
        {
            using (var p = OpenPackage("ChartAxisChange.xlsx", true))
            {
                var wb = p.Workbook;
                var ws = p.Workbook.Worksheets.Add("ChartWs");

                ws.Cells["A1:A10"].Formula = "ROW()+ 5";
                ws.Cells["C1:C10"].Formula = "COLUMN() + ROW() - 5";
                ws.Cells["D1:D10"].Formula = "COLUMN() + ROW() + 10";


                ws.Calculate();

                var barChart = ws.Drawings.AddBarChart("testChart", eBarChartType.ColumnClustered);

                var series = barChart.Series.Add("C1:C10");
                series.XSeries = ws.Cells["A1:A10"].Address;

                barChart.Title.Text = "OriginalChart";

                ws.Cells["A1:A10"].Formula = "\"Category\" & ROW()";

                var copiedChartDrawing = barChart.Copy(ws, 2, 2);
                var copiedChart = copiedChartDrawing.As.Chart.BarChart;

                ws.Calculate();

                //barChart.Axis[0].AxisType; 

                //copiedChart.

                copiedChart.Title.Text = "CopiedChart";

                //var someAxis = copiedChart.Axis[0];
                //copiedChart.Axis[0] = copiedChart.Axis[1];
                //copiedChart.Axis[1] = someAxis;

                //ws.Cells["A1"].SetCellValue(0, 0, "123");

                var axisType1 = barChart.Axis[0].AxisType;
                var axisType2 = barChart.Axis[1].AxisType;

                var barChart3D = ws.Drawings.AddBarChart("testChart3D", eBarChartType.ColumnClustered3D);

                var series2 = barChart3D.Series.Add("C1:C10");
                series2.XSeries = ws.Cells["A1:A10"].Address;

                barChart3D.Series.Add("D1:D10");

                //copiedChart.YAxis.ChangeAxisTypeReal(eAxisType.Serie);

                //copiedChart.Axis[0].ChangeAxisType(axisType2);
                //copiedChart.Axis[1].ChangeAxisType(axisType1);

                //copiedChart.Series.Delete(0);
                //var series2 = copiedChart.Series.Add("A1:A10");
                //series2.XSeries = ws.Cells["C1:C10"].Address;

                SaveAndCleanup(p);
            }
        }

        //[TestMethod]
        //public void TryDisableRecyclable()
        //{
        //    using (FileStream fStream = File.OpenRead("C:\\epplusTest\\Workbooks\\SimpleSortTest.xlsx"))
        //    {
        //        //Set up a memory stream from existing file without using epplus
        //        MemoryStream ms = new MemoryStream();
        //        ms.SetLength(fStream.Length);
        //        var msBuff = ms.GetBuffer();
        //        fStream.Read(msBuff, 0, (int)fStream.Length);

        //        ms.Seek(0, SeekOrigin.Begin);

        //        //Start using epplus
        //        ExcelPackage.MemorySettings.UseRecyclableMemory = false;

        //        //ExcelPackageSettings settings = new ExcelPackageSettings() { }

        //        var ms3 = new MemoryStream();
        //        var ms2 = new MemoryStream();

        //        var aPackage = new ExcelPackage(new FileInfo("C:\\epplusTest\\Workbooks\\SimpleSortTest.xlsx"), false);

        //        using (var p = new ExcelPackage(ms2,"Epplus"))
        //        {
        //            var someWs = p.Workbook.Worksheets.Add("TestWs");

        //            someWs.Cells["A1"].Value = "123";

        //            //Input your destination filepath here instead
        //            p.SaveAs(@"C:/temp/NoRecyclable.xlsx");
        //        }
        //    }


        //    //ExcelPackage.MemorySettings.UseRecyclableMemory = false;

        //    //var byteArr = File.ReadAllBytes("SimpleSortTest.xlsx");
        //    //var ms = new MemoryStream()

        //    //using (var ms = new MemoryStream())
        //    //{
        //    //    using (var p = new ExcelPackage(ms))
        //    //    {
        //    //        var someWs = p.Workbook.Worksheets.Add("TestWs");

        //    //        someWs.Cells["A1"].Value = "123";

        //    //        //Input your destination filepath here instead
        //    //        p.SaveAs(@"C:/temp/NoRecyclable.xlsx");
        //    //    }
        //    //}

        //    //using (var p = new ExcelPackage())
        //    //{
        //    //    var someWs = p.Workbook.Worksheets.Add("TestWs");

        //    //    someWs.Cells["A1"].Value = "123";

        //    //    var name = GetOutputFile("", "SomePackage.xlsx").FullName;
        //    //    p.SaveAs(name);
        //    //}
        //}
    }
}

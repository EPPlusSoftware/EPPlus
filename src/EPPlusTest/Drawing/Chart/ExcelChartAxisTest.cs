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
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.Xml;

namespace EPPlusTest.Drawing.Chart
{
    [TestClass]
    public class ExcelChartAxisTest : TestBase
    {
        private ExcelChartAxis axis;
        private ExcelPackage p;
        private ExcelWorksheet axisWsDates;

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

            p = OpenPackage("AxisDataSheet.xlsx", true);
            
            var now = DateTime.Now;

            axisWsDates = p.Workbook.Worksheets.Add("Sheet1");

            var headers = new List<string> { "Dates", "Sales", "Spending", "Net Profit" };

            var headRange = axisWsDates.Cells["A1:D1"];
            headRange.FillList(headers, x =>
            {
                x.Direction = eFillDirection.Row;
            });

            var dList = new List<DateTime> { now, now.AddDays(1), now.AddDays(2) };

            var dateRange = axisWsDates.Cells["$A$2:$A$4"].LoadFromCollection(dList);

            var salesRange = axisWsDates.Cells["B2:B4"].LoadFromCollection(new List<double> { 0d, 500d, 1500d });
            var spendRange = axisWsDates.Cells["C2:C4"].LoadFromCollection(new List<double> { 200d, 10d, 400d });
            var totalRange = axisWsDates.Cells["D2:D4"];

            totalRange.Formula = "B2-C2";
            totalRange.Calculate();

            axisWsDates.Cells["B2:D4"].Style.Numberformat.Format = "#,##0kr";
            axisWsDates.Cells["A2:A4"].Style.Numberformat.Format = "dd/mm/yyyy";
        }

        [TestCleanup]
        public void CleanUp()
        {
            SaveAndCleanup(p);
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
        public void LineChartChangeToDateAxis()
        {
            using (var p = OpenPackage("EpplusLineChartSimple.xlsx", true))
            {
                var ws = axisWsDates;

                var dateRange = axisWsDates.Cells["$A$2:$A$4"];

                var salesRange = axisWsDates.Cells["B2:B4"];
                var spendRange = axisWsDates.Cells["C2:C4"];
                var totalRange = axisWsDates.Cells["D2:D4"];

                var lineChart = ws.Drawings.AddLineChart("testLineChart", eLineChartType.Line);

                var ser1 = lineChart.Series.Add(salesRange, dateRange);
                ser1.HeaderAddress = ws.Cells["B1"];
                var ser2 = lineChart.Series.Add(spendRange, dateRange);
                ser2.HeaderAddress = ws.Cells["C1"];
                var ser3 = lineChart.Series.Add(totalRange, dateRange);
                ser3.HeaderAddress = ws.Cells["D1"];

                ws.Cells.AutoFitColumns();

                lineChart.XAxis.ChangeAxisTypeLimited(eAxisType.Date);
            }
        }

        [TestMethod]
        public void BarChartChangeToDateAxisThenResave()
        {
            //Create simple category Bar Chart with Epplus
            using (var p = OpenPackage("SimpleBarChart.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = "Categories";
                var header1Address = ws.Cells["B1"];
                var header2Address = ws.Cells["C1"];

                header1Address.Value = "Col1";
                header2Address.Value = "Col2";

                var dataRange = ws.Cells["B2:C3"];
                dataRange.Formula = "ROW() + COLUMN()";

                var catRange = ws.Cells["A2:A3"];

                catRange.Formula = "\"Row \" & (ROW()-1)";

                ws.Calculate();

                var barChart = ws.Drawings.AddBarChart("testChart", eBarChartType.ColumnClustered);

                var ser = barChart.Series.Add(dataRange.TakeSingleColumn(0), catRange);
                ser.HeaderAddress = header1Address;
                var ser2 = barChart.Series.Add(dataRange.TakeSingleColumn(1), catRange);
                ser2.HeaderAddress = header2Address;
                ws.Calculate();

                SaveAndCleanup(p);
            }

            //Read the file again and change axis type and data in the category column appropriately.
            using (var p = OpenPackage("SimpleBarChart.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];

                var barChartRead = ws.Drawings[0].As.Chart.BarChart;

                barChartRead.XAxis.ChangeAxisTypeLimited(eAxisType.Date);

                var dataRange = ws.Cells["A2:A3"];

                dataRange.ClearFormulas();

                var dtNow = DateTime.Now;

                ws.Cells["A2"].Value = dtNow;
                ws.Cells["A3"].Value = dtNow.AddDays(1);

                dataRange.Style.Numberformat.Format = "d-mmm";

                var outFile = GetOutputFile("", "SimpleBarChart_Resaved.xlsx").FullName;

                p.SaveAs(outFile);
            }
        }

        [TestMethod]
        public void MultipleCategoriesOnAxesColumn()
        {
            using (var p = OpenPackage("CreateEmployeesAndSales.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var headers = new List<string> { "Hire Date", "Employee Name", "Employee Position", "Sales This Year" };

                var headRange = ws.Cells["A1:D1"];
                headRange.FillList(headers, x =>
                {
                    x.Direction = eFillDirection.Row;
                });

                var dtNow = DateTime.Now;

                ws.Cells["A2"].Value = dtNow;
                ws.Cells["A3"].Value = dtNow.AddDays(1);

                ws.Cells["B2"].Value = "Ossian";
                ws.Cells["B3"].Value = "Mats";

                ws.Cells["C2"].Value = "Grunt";
                ws.Cells["C3"].Value = "Senior Developer";

                ws.Cells["D2"].Value = 5200d;
                ws.Cells["D3"].Value = 100d;
  
                ws.Cells["A2:A3"].Style.Numberformat.Format = "dd/mm/yyyy";
                ws.Cells["D2:D3"].Style.Numberformat.Format = "###,###0kr";

                var colChart = ws.Drawings.AddBarChart("colChart", eBarChartType.ColumnClustered);

                var ser = colChart.Series.Add(ws.Cells["D2:D3"], ws.Cells["A2:C3"]);
                ser.HeaderAddress = ws.Cells["D1"];
                var xSer = ser.XSeries;

                ws.Cells.Calculate();
                ws.Cells.AutoFitColumns();

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void CatToValLine()
        {
            using (var p = OpenPackage("CatToValLine.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("SomeSheet");

                var headers = new List<string> { "Categories", "Sales" };
                ws.Cells["A1:B1"].FillList(headers, x =>
                {
                    x.Direction = eFillDirection.Row;
                });

                var catRange = ws.Cells["A2:A3"];
                ws.Cells["A2"].Value = "Cat1";
                ws.Cells["A3"].Value = "Cat2";

                var valueRange = ws.Cells["B2:B3"];
                valueRange.Formula = "75 * 2 * ROW()";
                valueRange.Calculate();

                catRange.Formula = "ROW()+5";
                catRange.Calculate();

                var lChart = ws.Drawings.AddLineChart("lineChartOne", eLineChartType.Line);

                var ser = lChart.Series.Add(valueRange, ws.Cells["A2:A3"]);
                ser.HeaderAddress = ws.Cells["B1"];

                var axis = lChart.XAxis;
                axis.ChangeAxisTypeLimited(eAxisType.Date);
                //axis.ChangeAxisTypeLimited(eAxisType.Val);
                axis.CrossBetween = eCrossBetween.MidCat;
                lChart.DisplayBlanksAs = eDisplayBlanksAs.Gap;
                lChart.ShowDataLabelsOverMaximum = false;

                var xyScatter = ws.Drawings.AddScatterChart("xyChart1", eScatterChartType.XYScatter);

                var ser2 = xyScatter.Series.Add(valueRange, ws.Cells["A2:A3"]);
                ser2.HeaderAddress = ws.Cells["B1"];

                var axis2 = xyScatter.XAxis;
                axis2.ChangeAxisTypeLimited(eAxisType.Date);
                axis2.CrossBetween = eCrossBetween.MidCat;
                xyScatter.DisplayBlanksAs = eDisplayBlanksAs.Gap;
                xyScatter.ShowDataLabelsOverMaximum = false;

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void LineShouldNotAllowSeriesNon3D()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("sheetTest");
                var chart = ws.Drawings.AddLineChart("non3DLine", eLineChartType.LineStacked100);

                chart.XAxis.ChangeAxisTypeLimited(eAxisType.Serie);
            }
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void LineChartShouldNotAllowVal()
        {
            using (var p = OpenPackage("CatToValException.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("SomeSheet");

                var headers = new List<string> { "Categories", "Sales" };
                ws.Cells["A1:B1"].FillList(headers, x =>
                {
                    x.Direction = eFillDirection.Row;
                });

                var catRange = ws.Cells["A2:A3"];
                ws.Cells["A2"].Value = "Cat1";
                ws.Cells["A3"].Value = "Cat2";

                var valueRange = ws.Cells["B2:B3"];
                valueRange.Formula = "75 * 2 * ROW()";
                valueRange.Calculate();

                catRange.Formula = "ROW()+5";
                catRange.Calculate();

                var lChart = ws.Drawings.AddLineChart("lineChartOne", eLineChartType.Line);

                var ser = lChart.Series.Add(valueRange, ws.Cells["A2:A3"]);
                ser.HeaderAddress = ws.Cells["B1"];

                //We throw on val axisType. For Linecharts
                lChart.XAxis.ChangeAxisTypeLimited(eAxisType.Val);
            }
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void ColumnBarChartShouldNotAllowValX()
        {
            //Create simple category Bar Chart with Epplus
            using (var p = OpenPackage("ColumnBarCharts.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = "Categories";
                var header1Address = ws.Cells["B1"];
                var header2Address = ws.Cells["C1"];

                header1Address.Value = "Col1";
                header2Address.Value = "Col2";

                var dataRange = ws.Cells["B2:C3"];
                dataRange.Formula = "ROW() + COLUMN()";

                var catRange = ws.Cells["A2:A3"];

                catRange.Formula = "\"Row \" & (ROW()-1)";

                ws.Calculate();

                var barChart = ws.Drawings.AddBarChart("testChart", eBarChartType.BarClustered);

                var ser = barChart.Series.Add(dataRange.TakeSingleColumn(0), catRange);
                ser.HeaderAddress = header1Address;
                var ser2 = barChart.Series.Add(dataRange.TakeSingleColumn(1), catRange);
                ser2.HeaderAddress = header2Address;
                ws.Calculate();

                ws.Cells["A2:A3"].Formula = "ROW()+20";
                ws.Calculate();

                barChart.XAxis.ChangeAxisTypeLimited(eAxisType.Val);

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void AreaChart()
        {
            var ws = axisWsDates;

            var dateRange = axisWsDates.Cells["$A$2:$A$4"];

            var salesRange = axisWsDates.Cells["B2:B4"];
            var spendRange = axisWsDates.Cells["C2:C4"];
            var totalRange = axisWsDates.Cells["D2:D4"];

            var anArea = ws.Drawings.AddAreaChart("AnAreaChart", eAreaChartType.Area);

            var ser1 = anArea.Series.Add(salesRange, dateRange);
            ser1.HeaderAddress = ws.Cells["B1"];
            var ser2 = anArea.Series.Add(spendRange, dateRange);
            ser2.HeaderAddress = ws.Cells["C1"];
            var ser3 = anArea.Series.Add(totalRange, dateRange);
            ser3.HeaderAddress = ws.Cells["D1"];

            anArea.XAxis.ChangeAxisTypeLimited(eAxisType.Val);
        }

        [TestMethod]
        public void XYChart()
        {
            var ws = axisWsDates;

            var dateRange = axisWsDates.Cells["$A$2:$A$4"];

            var salesRange = axisWsDates.Cells["B2:B4"];
            var spendRange = axisWsDates.Cells["C2:C4"];
            var totalRange = axisWsDates.Cells["D2:D4"];

            var xyChart = ws.Drawings.AddScatterChart("ScatterChart", eScatterChartType.XYScatter);

            //var ser1 = xyChart.Series.Add(salesRange, dateRange);
            //ser1.HeaderAddress = ws.Cells["B1"];
            var ser2 = xyChart.Series.Add(dateRange, spendRange);
            ser2.HeaderAddress = ws.Cells["C1"];
            //var ser3 = xyChart.Series.Add(totalRange, dateRange);
            //ser3.HeaderAddress = ws.Cells["D1"];

            //dateRange["A2"].Value = "SomeValue";
            //dateRange["A3"].Value = "Some2Value";
            //dateRange["A4"].Value = "Some3Value";

            xyChart.XAxis.ChangeAxisTypeLimited(eAxisType.Val);
            xyChart.YAxis.ChangeAxisTypeLimited(eAxisType.Val);
            ws.Calculate();
            //xyChart.XAxis.ChangeAxisTypeLimited(eAxisType.Cat);
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

                var headers = new List<string> { "Categories", "Sales" };
                ws.Cells["A1:B1"].FillList(headers);

                ws.Cells["A1"].Value = "Categories";
                ws.Cells["B1"].Value = "Col1";
                ws.Cells["C1"].Value = "Col2";

                var dataRange = ws.Cells["B2:C3"];
                dataRange.Formula = "ROW() + COLUMN()";

                var catRange = ws.Cells["A2:A3"];

                catRange.Formula = "\"Row \" & (ROW()-1)";

                ws.Calculate();

                var barChart = ws.Drawings.AddBarChart("testChart", eBarChartType.ColumnClustered);

                barChart.Series.Add(dataRange.TakeSingleColumn(0), catRange);
                barChart.Series.Add(dataRange.TakeSingleColumn(1), catRange);

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
        public void LoadFromCollectionDateTime()
        {
            using (var p = OpenPackage("LoadFromCollectionDateTime.xlsx", true))
            {
                var now = DateTime.Now;

                var ws = p.Workbook.Worksheets.Add("Sheet1");

                var dateRange = ws.Cells["$A$2:$A$4"];

                var dList = new List<DateTime> { now, now.AddDays(1), now.AddDays(2) };

                dList.Add(DateTime.Today);

                var retRange = ws.Cells["A2"].LoadFromCollection(dList);

                Assert.AreEqual("A2:A5", retRange.Address);

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void FillListFirstRow()
        {
            using (var p = OpenPackage("FillListFirstRow.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");

                var headers = new List<string> { "Dates", "Sales total", "Spending Total", "Day total" };

                var headRange = ws.Cells["A1:D1"];
                headRange.FillList(headers, x =>
                {
                    x.Direction = eFillDirection.Row;
                });

                Assert.AreEqual(ws.Cells["A1"].Text, headers[0]);
                Assert.AreEqual(ws.Cells["B1"].Text, headers[1]);
                Assert.AreEqual(ws.Cells["C1"].Text, headers[2]);
                Assert.AreEqual(ws.Cells["D1"].Text, headers[3]);

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
    }
}

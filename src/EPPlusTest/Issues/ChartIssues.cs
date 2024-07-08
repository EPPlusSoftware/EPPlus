﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml;
using System.IO;
using System.Drawing;
using System.Text;
using System;
using System.Security.Principal;

namespace EPPlusTest.Issues
{
	[TestClass]
	public class ChartIssues : TestBase
	{
		[ClassInitialize]
		public static void Init(TestContext context)
		{
		}
		[ClassCleanup]
		public static void Cleanup()
		{
		}
		[TestInitialize]
		public void Initialize()
		{
		}
		[TestMethod]
		public void s578()
		{
			using (var p = OpenPackage("s578.xlsx", true))
			{
				var sheet = p.Workbook.Worksheets.Add("Sheet1");
				// do work here
				sheet.Cells["P11"].Value = "2023/10/01";
				sheet.Cells["Q11"].Value = "2023/10/02";
				sheet.Cells["R11"].Value = "2023/10/03";
				sheet.Cells["S11"].Value = "2023/10/04";
				sheet.Cells["P12"].Value = 3.0;
				sheet.Cells["Q12"].Value = 4.0;
				sheet.Cells["R12"].Value = 5.0;
				sheet.Cells["S12"].Value = 4.5;
				sheet.Cells["P13"].Value = 4.0;
				sheet.Cells["Q13"].Value = 6.0;
				sheet.Cells["R13"].Value = 7.0;
				sheet.Cells["S13"].Value = 6.0;
				sheet.Cells["P14"].Value = 5.0;
				sheet.Cells["Q14"].Value = 2.0;
				sheet.Cells["R14"].Value = 5.0;
				sheet.Cells["S14"].Value = 2.0;

				ExcelLineChart chart = sheet.Drawings.AddLineChart("test chart", eLineChartType.Line);

				var LabelRange = sheet.Cells["P11:S11"];
				var DataRange = sheet.Cells["P12:S12"];

				var chartSerie = chart.Series.Add(DataRange, LabelRange);
				chartSerie.Header = "test";
				chart.Legend.Border.LineStyle = eLineStyle.Solid;
				chart.Legend.Border.Width = 1;
				chart.Legend.Position = eLegendPosition.Right;
				chart.Legend.TextSettings.Effect.SetPresetReflection(ePresetExcelReflectionType.FullTouching);
				chart.XAxis.TextSettings.Effect.SetPresetReflection(ePresetExcelReflectionType.HalfTouching);
				chart.XAxis.TextSettings.Fill.Style = eFillStyle.GradientFill;
				chart.XAxis.TextSettings.Fill.GradientFill.Colors.AddRgb(0, System.Drawing.Color.DarkSeaGreen);
				chart.XAxis.TextSettings.Fill.GradientFill.Colors.AddRgb(50, System.Drawing.Color.LightCoral);
				chart.XAxis.TextSettings.Outline.Fill.Style = eFillStyle.SolidFill;
				chart.XAxis.TextSettings.Outline.LineStyle = eLineStyle.Dash;
				chart.Title.Text = "Title 1";
				chart.Title.TextSettings.Effect.SetPresetGlow(ePresetExcelGlowType.Accent1_5Pt);
				SaveAndCleanup(p);
			}
		}
		[TestMethod]
		public void s598()
		{
			using (var p = OpenPackage("s598.xlsx", true))
			{
				var sheet = p.Workbook.Worksheets.Add("Sheet1");
				// do work here
				sheet.Cells["P11"].Value = "2023/10/01";
				sheet.Cells["Q11"].Value = "2023/10/02";
				sheet.Cells["R11"].Value = "2023/10/03";
				sheet.Cells["S11"].Value = "2023/10/04";
				sheet.Cells["P12"].Value = 3.0;
				sheet.Cells["Q12"].Value = 4.0;
				sheet.Cells["R12"].Value = 5.0;
				sheet.Cells["S12"].Value = 4.5;
				sheet.Cells["P13"].Value = 4.0;
				sheet.Cells["Q13"].Value = 6.0;
				sheet.Cells["R13"].Value = 7.0;
				sheet.Cells["S13"].Value = 6.0;
				sheet.Cells["P14"].Value = 5.0;
				sheet.Cells["Q14"].Value = 2.0;
				sheet.Cells["R14"].Value = 5.0;
				sheet.Cells["S14"].Value = 2.0;

				ExcelLineChart chart = sheet.Drawings.AddLineChart("test chart", eLineChartType.Line);

				var LabelRange = sheet.Cells["P11:S11"];
				var DataRange = sheet.Cells["P12:S12"];

				var chartSerie = chart.Series.Add(DataRange, LabelRange);
				chartSerie.Header = "test";
				chart.Title.Text = "test Graph";
				chart.Title.TextSettings.Fill.Style = OfficeOpenXml.Drawing.eFillStyle.SolidFill;
				chart.Title.TextSettings.Fill.SolidFill.Color.SetRgbColor(System.Drawing.Color.Black);
				chart.Legend.Position = eLegendPosition.Right;

				chart.Legend.Font.UnderLine = OfficeOpenXml.Style.eUnderLineType.Single;

				/* if you remove the following line, reflection setting is OK */
				chart.Legend.Font.UnderLineColor = System.Drawing.Color.Red;

				chart.Legend.TextSettings.Fill.Style = OfficeOpenXml.Drawing.eFillStyle.SolidFill;
				chart.Legend.TextSettings.Fill.SolidFill.Color.SetRgbColor(System.Drawing.Color.Black);
				chart.Legend.TextSettings.Fill.Transparancy = 0;
				chart.Legend.TextSettings.Effect.SetPresetReflection(OfficeOpenXml.Drawing.ePresetExcelReflectionType.FullTouching);

				SaveAndCleanup(p);
			}

		}
		[TestMethod]
		public void s599()
		{
			using (var p = OpenPackage("s599.xlsx", true))
			{
				var sheet = p.Workbook.Worksheets.Add("Sheet1");
				// do work here
				sheet.Cells["P11"].Value = "2023/10/01";
				sheet.Cells["Q11"].Value = "2023/10/02";
				sheet.Cells["R11"].Value = "2023/10/03";
				sheet.Cells["S11"].Value = "2023/10/04";
				sheet.Cells["P12"].Value = 3.0;
				sheet.Cells["Q12"].Value = 4.0;
				sheet.Cells["R12"].Value = 5.0;
				sheet.Cells["S12"].Value = 4.5;
				sheet.Cells["P13"].Value = 4.0;
				sheet.Cells["Q13"].Value = 6.0;
				sheet.Cells["R13"].Value = 7.0;
				sheet.Cells["S13"].Value = 6.0;
				sheet.Cells["P14"].Value = 5.0;
				sheet.Cells["Q14"].Value = 2.0;
				sheet.Cells["R14"].Value = 5.0;
				sheet.Cells["S14"].Value = 2.0;

				ExcelLineChart chart = sheet.Drawings.AddLineChart("test chart", eLineChartType.Line);

				var LabelRange = sheet.Cells["P11:S11"];
				var DataRange = sheet.Cells["P12:S12"];

				var chartSerie = chart.Series.Add(DataRange, LabelRange);
				chartSerie.Header = "test";
				chart.Title.Text = "test Graph";
				chart.Title.TextSettings.Fill.Style = OfficeOpenXml.Drawing.eFillStyle.SolidFill;
				chart.Title.TextSettings.Fill.SolidFill.Color.SetRgbColor(System.Drawing.Color.Black);

				chart.DataLabel.ShowValue = true;

				/* the following 2 lines make Excel unable to open the file */
				chart.DataLabel.TextSettings.Fill.Style = OfficeOpenXml.Drawing.eFillStyle.SolidFill;
				chart.DataLabel.TextSettings.Fill.SolidFill.Color.SetRgbColor(System.Drawing.Color.Blue);

				chart.Legend.Position = eLegendPosition.Right;

				SaveAndCleanup(p);
			}
		}
		[TestMethod]
		public void s643()
		{
			using (var p = OpenTemplatePackage("s643.xlst"))
			{
				SaveWorkbook("s643.xlsx", p);
			}
		}
		[TestMethod]
		public void i1401()
		{
			using (var p = OpenPackage("i1401.xlsx", true))
			{
				var chartWorksheet = p.Workbook.Worksheets.Add("Sheet1");
				LoadTestdata(chartWorksheet);
				var chart = chartWorksheet.Drawings.AddBarChart("chart1", eBarChartType.ColumnClustered);
				chart.Series.Add("B1:B100", "A1:A100");
				chart.SetPosition(1, 10, 12, 0);
				chart.SetSize(1200, 580);
				chart.Legend.Remove();
				chart.Title.Text = "t";
				chart.Title.Font.Bold = true;
				chart.Title.Font.UnderLine = OfficeOpenXml.Style.eUnderLineType.Single;
				chart.Title.Font.Size = 16;

				chart.XAxis.LabelPosition = eTickLabelPosition.NextTo;
				chart.XAxis.TextBody.WrapText = eTextWrappingType.Square;
				chart.XAxis.TextBody.Rotation = 45D;
				chart.DataLabel.ShowValue = true;
				chart.DataLabel.Position = eLabelPosition.OutEnd;
				chart.DataLabel.TextBody.Rotation = 45D; //<= This line causes the error.

				SaveAndCleanup(p);
			}
		}

		[TestMethod]
		public void s694()
		{
			using (var p = OpenPackage("s694.xlsx", true))
			{
				// Add a new worksheet to the empty workbook
				var worksheet = p.Workbook.Worksheets.Add("Sheet1");

				// Add some data for the pie chart
				worksheet.Cells["A1"].Value = "Category";
				worksheet.Cells["B1"].Value = "Value";
				worksheet.Cells["A2"].Value = "Category 1";
				worksheet.Cells["B2"].Value = 20;
				worksheet.Cells["A3"].Value = "Category 2";
				worksheet.Cells["B3"].Value = 30;
				worksheet.Cells["A4"].Value = "Category 3";
				worksheet.Cells["B4"].Value = 50;
				worksheet.Cells["A5"].Value = "Category 4";
				worksheet.Cells["B5"].Value = 34;
				worksheet.Cells["A6"].Value = "Category 5";
				worksheet.Cells["B6"].Value = 12;

				// Add a pie chart to the worksheet
				var pieChart = worksheet.Drawings.AddChart("pieChart", eChartType.Pie) as ExcelPieChart;
				pieChart.Title.Text = "Pie Chart Example";
				pieChart.SetPosition(1, 0, 3, 0);
				pieChart.SetSize(600, 400);

				// Define the data series for the pie chart
				var series = pieChart.Series.Add(worksheet.Cells["B2:B6"], worksheet.Cells["A2:A6"]);

				series.DataPoints.Add(0);
				series.DataPoints.Add(1);
				series.DataPoints.Add(2);

				series.DataPoints[0].Fill.Style = eFillStyle.SolidFill;
				series.DataPoints[0].Fill.Color = Color.Red;
				series.DataPoints[1].Fill.Style = eFillStyle.SolidFill;
				series.DataPoints[1].Fill.Color = Color.Blue;
				series.DataPoints[2].Fill.Style = eFillStyle.SolidFill;
				series.DataPoints[2].Fill.Color = Color.Green;
				//            pieChart.Series.Add(worksheet.Cells["A2:B2"]);
				//            pieChart.Series.Add(worksheet.Cells["A3:B3"]);
				//            pieChart.Series.Add(worksheet.Cells["A4:B4"]);
				//            pieChart.Series.Add(worksheet.Cells["A5:B5"]);
				//            pieChart.Series.Add(worksheet.Cells["A6:B6"]);

				//for(int i = 2; i < 7; i++)
				//{
				//                var series = pieChart.Series.Add(worksheet.Cells[$"A{i}:B{i}"]);
				//	series.Fill.Style = eFillStyle.SolidFill;
				//	series.Fill.Color = Color.BlueViolet;
				//                //pieChart.Series[i].Fill.Style = eFillStyle.SolidFill;
				//                //               pieChart.Series[i].Fill.Color = Co;
				//            }

				// Apply some styling to the chart
				pieChart.DataLabel.ShowCategory = true;
				pieChart.DataLabel.ShowPercent = true;
				pieChart.DataLabel.ShowLeaderLines = true;
				//pieChart.Style = eChartStyle.None;

				//pieChart.Style = eChartStyle.Style18;

				//pieChart.Series[0] = pieChart.


				////pieChart.Series[0].Fill.Style = eFillStyle.SolidFill;
				////pieChart.Series[0].Fill.Color = Color.Red;

				////pieChart.Series[0].

				//////pieChart.Series[1].Fill.Style = eFillStyle.SolidFill;
				//////pieChart.Series[1].Fill.Color = Color.Blue;

				SaveAndCleanup(p);
			}
		}

		[TestMethod]
		public void s694_2()
		{
			// Ensure ExcelPackage works with non-commercial license
			ExcelPackage.LicenseContext = LicenseContext.Commercial;

			// Create a new Excel package
			using (ExcelPackage package = OpenPackage("s694_2.xlsx", true))
			{
				// Add a new worksheet to the empty workbook
				ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

				// Add some data for the pie chart
				worksheet.Cells["A1"].Value = "Category";
				worksheet.Cells["B1"].Value = "Value";
				worksheet.Cells["A2"].Value = "Category 1";
				worksheet.Cells["B2"].Value = 20;
				worksheet.Cells["A3"].Value = "Category 2";
				worksheet.Cells["B3"].Value = 30;
				worksheet.Cells["A4"].Value = "Category 3";
				worksheet.Cells["B4"].Value = 50;
				worksheet.Cells["A5"].Value = "Category 4";
				worksheet.Cells["B5"].Value = 34;
				worksheet.Cells["A6"].Value = "Category 5";
				worksheet.Cells["B6"].Value = 12;

				// Add a pie chart to the worksheet
				//var pieChart = worksheet.Drawings.AddChart("pieChart", eChartType.Pie) as ExcelPieChart;
				using (FileStream template = new FileStream(@"C:\epplusTest\Workbooks\Chart1.crtx", FileMode.Open, FileAccess.Read))
				{
					var pieChart = worksheet.Drawings.AddChartFromTemplate(template, "pieChart").As.Chart.PieChart;

					pieChart.Title.Text = "Pie Chart Example";
					pieChart.SetPosition(1, 0, 3, 0);
					pieChart.SetSize(600, 400);

					pieChart.DataLabel.ShowCategory = false;
					pieChart.DataLabel.ShowPercent = false;

					//pieChart.Series[0].Series = "sheet1!$B$2:$B$6";

					//pieChart.Series[0].Series = "$B$2:$B$6";
					//pieChart.Series[0].XSeries = "$A$2:$A$6";
					//// Define the data series for the pie chart
					///

					pieChart.ApplyThemeOverrideOnFirstSeries = true;
                    var series = pieChart.Series.Add(worksheet.Cells["B2:B6"], worksheet.Cells["A2:A6"]);

					//series.xs
					//series.

                    // Apply some styling to the chart
                    pieChart.DataLabel.ShowCategory = true;
					pieChart.DataLabel.ShowPercent = true;
					pieChart.DataLabel.ShowLeaderLines = true;
					//pieChart.StyleManager.LoadTemplateStyles(template);
				}

				SaveAndCleanup(package);
			}
		}
	}
}

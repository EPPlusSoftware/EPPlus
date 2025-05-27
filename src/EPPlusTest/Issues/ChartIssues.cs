using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Security.Principal;
using System.Text;

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
				worksheet.Cells["A1"].Value = "Cat";
				worksheet.Cells["B1"].Value = "Value";
				worksheet.Cells["A2"].Value = "Cat 1";
				worksheet.Cells["B2"].Value = 15;
				worksheet.Cells["A3"].Value = "Cat 2";
				worksheet.Cells["B3"].Value = 24;
				worksheet.Cells["A4"].Value = "Cat 3";
				worksheet.Cells["B4"].Value = 40;
				worksheet.Cells["A5"].Value = "Cat 4";
				worksheet.Cells["B5"].Value = 23;
				worksheet.Cells["A6"].Value = "Cat 5";
				worksheet.Cells["B6"].Value = 4;

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

				pieChart.DataLabel.ShowCategory = true;
				pieChart.DataLabel.ShowPercent = true;
				pieChart.DataLabel.ShowLeaderLines = true;

				SaveAndCleanup(p);
			}
		}

		[TestMethod]
		public void s694_2()
		{
			// Create a new Excel package
			using (ExcelPackage package = OpenPackage("s694_2.xlsx", true))
			{
				// Add a new worksheet to the empty workbook
				ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

				// Add some data for the pie chart
				worksheet.Cells["A1"].Value = "Cat";
				worksheet.Cells["B1"].Value = "Value";
				worksheet.Cells["A2"].Value = "Cat 1";
				worksheet.Cells["B2"].Value = 15;
				worksheet.Cells["A3"].Value = "Cat 2";
				worksheet.Cells["B3"].Value = 24;
				worksheet.Cells["A4"].Value = "Cat 3";
				worksheet.Cells["B4"].Value = 40;
				worksheet.Cells["A5"].Value = "Cat 4";
				worksheet.Cells["B5"].Value = 23;
				worksheet.Cells["A6"].Value = "Cat 5";
				worksheet.Cells["B6"].Value = 4;

				var currDir = Directory.GetCurrentDirectory();

				// Add a pie chart to the worksheet
				using (FileStream template = new FileStream($@"{currDir}\Resources\PieChartTemplate2.crtx", FileMode.Open, FileAccess.Read))
				{
					var pieChart = worksheet.Drawings.AddChartFromTemplate(template, "pieChart").As.Chart.PieChart;

					pieChart.Title.Text = "Pie Chart Example";
					pieChart.SetPosition(1, 0, 3, 0);
					pieChart.SetSize(600, 400);

					pieChart.DataLabel.ShowCategory = false;
					pieChart.DataLabel.ShowPercent = false;

					var series2 = pieChart.Series.Add(worksheet.Cells["B2:B6"], worksheet.Cells["A2:A6"]);

					// Apply some styling to the chart-/
					pieChart.DataLabel.ShowCategory = false;
					pieChart.DataLabel.ShowPercent = false;
					pieChart.DataLabel.ShowLeaderLines = false;

					Assert.AreEqual(pieChart.Series[0].DataPoints[1].Fill.Color, Color.FromArgb(255, 165, 234, 54));
				}

				SaveAndCleanup(package);
			}
		}

		[TestMethod]
		public void s694_3()
		{

			// Create a new Excel package
			using (ExcelPackage package = OpenPackage("s694_3.xlsx", true))
			{
				// Add a new worksheet to the empty workbook
				ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

				// Add some data for the pie chart
				worksheet.Cells["A1"].Value = "Cat";
				worksheet.Cells["B1"].Value = "Value";
				worksheet.Cells["A2"].Value = "Cat 1";
				worksheet.Cells["B2"].Value = 15;
				worksheet.Cells["C2"].Value = 25;
				worksheet.Cells["A3"].Value = "Cat 2";
				worksheet.Cells["B3"].Value = 24;
				worksheet.Cells["C3"].Value = 33;
				worksheet.Cells["A4"].Value = "Cat 3";
				worksheet.Cells["B4"].Value = 40;
				worksheet.Cells["C4"].Value = 47;
				worksheet.Cells["A5"].Value = "Cat 4";
				worksheet.Cells["B5"].Value = 23;
				worksheet.Cells["C5"].Value = 13;
				worksheet.Cells["A6"].Value = "Cat 5";
				worksheet.Cells["B6"].Value = 4;
				worksheet.Cells["C6"].Value = 12;

				var currDir = Directory.GetCurrentDirectory();

				// Add a pie chart to the worksheet
				using (FileStream template = new FileStream($@"{currDir}\Resources\StackedColumnChart.crtx", FileMode.Open, FileAccess.Read))
				{
					var barChart = worksheet.Drawings.AddChartFromTemplate(template, "colChart").As.Chart.BarChart;

					barChart.Title.Text = "Stacked Column Example";
					barChart.SetPosition(0, 0, 6, 0);
					barChart.SetSize(1200, 400);

					var range = worksheet.Cells["A2:C6"];

					var series1 = barChart.Series.Add(range.TakeSingleColumn(1), range.TakeSingleColumn(0));
					var series2 = barChart.Series.Add(range.TakeSingleColumn(2), range.TakeSingleColumn(0));
				}

				SaveAndCleanup(package);
			}
		}
		[TestMethod]
		public void s754()
		{
			using (var package = OpenTemplatePackage("s754.xlsx"))
			{
				var workbook = package.Workbook;
				SaveAndCleanup(package);
			}
		}
		//i1886 handling
		[TestMethod]
		public void i886()
		{
			using (var package = OpenTemplatePackage("LiteralsExample.xlsx"))
			{
				var wb = package.Workbook;
				var ws = wb.Worksheets[0];

				var numLitChart = ws.Drawings[0].As.Chart.BarChart;

				var serie = numLitChart.Series[0];
				var numLits = serie.NumberLiteralsX;
				var numlitsY = serie.NumberLiteralsY;

				serie.Series = "{10,20,50}";
				serie.XSeries = "{'col1','col2','col3','col4'}";

				SaveAndCleanup(package);
			}
		}

        [TestMethod]
        public void CreateStringLitterals()
        {
			using (var package = OpenPackage("LitteralsSetting.xlsx", true))
			{
				var wb = package.Workbook;
				var ws = wb.Worksheets.Add("NewWork");

				var numLitChart = ws.Drawings.AddBarChart("bar", eBarChartType.ColumnClustered);

				var serie = numLitChart.Series.Add("{10,20,50}");
				serie.XSeries = "{'col1','col2','col3'}";
                SaveAndCleanup(package);
            }
        }


        [TestMethod]
        public void SC870_ALT()
		{
			using (var package = OpenTemplatePackage("s870.xlsx"))
			{
				var wb = package.Workbook;

				wb.FullCalcOnLoad = false;

				var worksheet = package.Workbook.Worksheets["MASTER"];

				var originalFormulaPart = "VLOOKUP(B11, Salgsfragt!B:C, 2, TRUE)";
				var changedFormulaPart = "VLOOKUP(B11, Salgsfragt!B6:C65, 2, TRUE)";

				var cellC19 = worksheet.Cells["C19"];


				var cellC25 = worksheet.Cells["C25"];
				cellC25.Formula = $"IF(B7=\"Denmark\", IF(VLOOKUP(B10, Produkter!A:Q, 13, FALSE)=\"PL2\", Salgsfragt!F6 * B11, ({changedFormulaPart} * B11) + IF(VLOOKUP(B10, Produkter!A:Q, 13, FALSE)=\"PT7\", VLOOKUP(\"PT7\", Salgsfragt!E:G, 2, FALSE) * B11, 0) + IF(Kalkulator!C52=\"Ja\", Salgsfragt!F5 * B11, 0)), Kalkulator!F21)";

				cellC19.Formula = $"IF(B7=\"Denmark\", IF(VLOOKUP(B10, Produkter!A:Q, 13, FALSE)=\"PL2\", Salgsfragt!F6 * B11, ({changedFormulaPart} * B11) + IF(VLOOKUP(B10, Produkter!A:Q, 13, FALSE)=\"PT7\", VLOOKUP(\"PT7\", Salgsfragt!E:G, 2, FALSE) * B11, 0) + IF(Kalkulator!C52=\"Ja\", Salgsfragt!F5 * B11, 0)), Kalkulator!F21)/7.46";

				//Alternative Slightly more efficent solution:
				//cellC19.Formula = "C25/7.46";
				//Since repeating the formula should be unnecesary.

				wb.Calculate();

				var val = cellC19.Value;
				var val2 = cellC25.Value;

				//Save workbook

                SaveAndCleanup(package);
            }
		}


			[TestMethod]
		public void SC870_EpplusOnly()
		{
			using(var p = new ExcelPackage())
			{
				var wb = p.Workbook;
				var ws = wb.Worksheets.Add("VLookupTest");
				List<int> col1Values = new List<int>{ 1, 2, 4, 7, 11, 16, 21, 27 };
				List<int> col2Values = new List<int> {  400, 365, 315, 280, 250, 215, 200, 170};

				ws.Cells["B6:B13"].LoadFromCollection(col1Values);
				ws.Cells["C6:C13"].LoadFromCollection(col2Values);

				ws.Cells["A11"].Value = 1;

				ws.Cells["F5"].Formula = "VLOOKUP(A11, B:C, 2, TRUE)";

				ws.Calculate();

				//Epplus returns N/A here but it appears to calculate correctly in excel. Why?
				var outputValue = ws.Cells["F5"].Value;

				//Save Workbook
			}
        }

        [TestMethod]
        public void SC870()
        {
			using (var package = OpenTemplatePackage("s870.xlsx"))
			{
				var wb = package.Workbook;

                var worksheet = package.Workbook.Worksheets[0];

				worksheet.Cells["B7"].Value = "Denmark";
				worksheet.Cells["B8"].Value = (int)9000;
				worksheet.Cells["B10"].Value = "18L BIOBED bioactive bedding ORGANIC (full pallet)";
				worksheet.Cells["B11"].Value = (int)1;

				foreach (var sheet in package.Workbook.Worksheets)
				{
					sheet.Hidden = eWorkSheetHidden.Visible;
					//sheet.Calculate();
				}

                //package.Workbook.Calculate();

                //wb.wo
                //"B6:C13"
                //worksheet.Cells["F15"].Formula = "VLOOKUP(B11, Salgsfragt!B6:C13, 2, TRUE)";

                worksheet.Cells["F15"].Formula = "VLOOKUP(B11, Salgsfragt!B:C, 2, TRUE)";

				var sWs = package.Workbook.Worksheets.GetByName("Salgsfragt");
				sWs.Cells["B4"].Value = null;
                sWs.Cells["B2"].Value = null;


                // Output from the logger will be written to the following file
                var logfile = new FileInfo(@"c:\temp\logfile.txt");
                // Attach the logger before the calculation is performed.
                package.Workbook.FormulaParserManager.AttachLogger(logfile);
				worksheet.Cells["F15"].Calculate();
                package.Workbook.FormulaParserManager.DetachLogger();

				var someVal = worksheet.Cells["F15"].Value;

                var errorText = worksheet.Cells["D8"].Text;

                var cellC19 = worksheet.Cells["C19"];
                var cellC25 = worksheet.Cells["C25"];

				worksheet.Calculate();

				var val1 = cellC19.Value;
				var val2 = cellC25.Value;


                //// Output from the logger will be written to the following file
                //var logfile = new FileInfo(@"c:\temp\logfile.txt");
                //// Attach the logger before the calculation is performed.
                //package.Workbook.FormulaParserManager.AttachLogger(logfile);
                //worksheet.Cells["C19"].Calculate();
                //// The following method removes any logger attached to the workbook.
                //package.Workbook.FormulaParserManager.DetachLogger();

                //var transportPriceVal = worksheet.Cells["C19"].Value;

                //if (!string.IsNullOrEmpty(errorText))
                //{
                //	return BadRequest(new List<string>() { errorText });
                //}

                // Save

                SaveAndCleanup(package);
				//var savePath = Path.Combine(Directory.GetCurrentDirectory(), "PriceData", "calculator_result.xlsx");
				//package.SaveAs(new FileInfo(savePath));

				//var totalM3Text = worksheet.Cells["B14"].Text;
				//var expectedTransitTime = worksheet.Cells["B15"].Text;
				//var itemPriceEurText = worksheet.Cells["C18"].Value.ToString();
				//var transportPriceEurText = worksheet.Cells["C19"].Value.ToString();
				//var totalPriceEurText = worksheet.Cells["C20"].Value.ToString();

				//double? totalM3 = null;
				//double? itemPriceEur = null;
				//double? transportPriceEur = null;
				//double? totalPriceEur = null;

				//if (!string.IsNullOrEmpty(totalM3Text))
				//{
				//	totalM3 = double.Parse(totalM3Text.Replace(",", "."), CultureInfo.InvariantCulture);
				//}

				//if (!string.IsNullOrEmpty(itemPriceEurText))
				//{
				//	itemPriceEur = double.Parse(itemPriceEurText.Replace(",", "."), CultureInfo.InvariantCulture);
				//}

				//if (!string.IsNullOrEmpty(transportPriceEurText))
				//{
				//	transportPriceEur = double.Parse(transportPriceEurText.Replace(",", "."),
				//		CultureInfo.InvariantCulture);
				//}

				//if (!string.IsNullOrEmpty(totalPriceEurText))
				//{
				//	totalPriceEur = double.Parse(totalPriceEurText.Replace(",", "."), CultureInfo.InvariantCulture);
				//}

				//return Ok(new
				//{
				//	TotalM3 = totalM3,
				//	ExpectedTransitTime = expectedTransitTime,
				//	ItemPriceEur = itemPriceEur,
				//	TransportPriceEur = transportPriceEur,
				//	TotalPriceEur = totalPriceEur,
				//	ItemName = worksheet.Cells["B10"].Text,
				//});
			}
        }
    }
}

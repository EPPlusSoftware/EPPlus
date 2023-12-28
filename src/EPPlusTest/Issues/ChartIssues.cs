using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
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
				//chart.Legend.Border.LineStyle = OfficeOpenXml.Drawing.eLineStyle.Solid;
				//chart.Legend.Border.Width=1;
				//chart.Legend.Position = eLegendPosition.Right;
				//chart.Legend.Effect.SetPresetReflection(OfficeOpenXml.Drawing.ePresetExcelReflectionType.FullTouching);
				//chart.Legend.Effect.Reflection.Distance = 1.0;


				SaveAndCleanup(p);
			}
		}
	}
}

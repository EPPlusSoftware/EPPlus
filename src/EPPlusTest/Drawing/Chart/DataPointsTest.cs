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
using EPPlusTest.SaveFunctions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Style;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Drawing.Theme;
using System.Drawing;
using System.IO;

namespace EPPlusTest.Drawing.Chart
{
    [TestClass]
    public class DataPointsTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("DataPoints.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }

        [TestMethod]
        public void LineChart()
        {
            var ws = _pck.Workbook.Worksheets.Add("LineChart");
            LoadTestdata(ws);

            var chart=ws.Drawings.AddLineChart("LineChart1", eLineChartType.Line);
            var serie = chart.Series.Add("D2:D100", "A2:A100");
            var point = serie.DataPoints.Add(3);
            point.Border.Fill.Color = Color.Red;
            point.Border.Fill.Style = eFillStyle.SolidFill;
            point.Fill.Color = Color.Green;
            chart.SetPosition(1, 0, 5, 0);
            var svg = chart.ToSvg();
            
            File.WriteAllText($"{_worksheetPath}svg\\EPPlusLineChart1.svg", svg);
        }
        [TestMethod]
        public void PieChart()
        {
            var ws = _pck.Workbook.Worksheets.Add("PieChart");
            LoadTestdata(ws);

            var chart = ws.Drawings.AddPieChart("PieChart1", ePieChartType.Pie);
            var serie = chart.Series.Add("D2:D6", "A2:A6");
            var point = serie.DataPoints.Add(0);
            point.Border.Fill.Color = Color.Red;
            point.Border.Fill.Style = eFillStyle.SolidFill;
            point.Fill.Color = Color.Green;
            chart.SetPosition(1, 0, 5, 0);

            var svg = chart.ToSvg();

            //SaveAndCleanup(_pck);

            //File.WriteAllText($"{_worksheetPath}svg\\EPPlusPieChart1.svg", svg);
        }
        [TestMethod]
        public void BarChart()  
        {
            var ws = _pck.Workbook.Worksheets.Add("BarChart");
            LoadTestdata(ws);

            var chart = ws.Drawings.AddBarChart("BarChart1", eBarChartType.BarStacked);
            var serie = chart.Series.Add("D2:D5", "A2:A5");
            var point = serie.DataPoints.Add(0);
            point.Border.Fill.Color = Color.Blue;
            point.Border.Fill.Style = eFillStyle.SolidFill;
            point.Fill.Style = eFillStyle.SolidFill;
            point.Fill.SolidFill.Color.SetRgbColor(Color.Yellow);
            point.Fill.Transparency = 5;            
            Assert.AreEqual(eColorTransformType.Alpha, point.Fill.SolidFill.Color.Transforms[0].Type);
            Assert.AreEqual(95, point.Fill.SolidFill.Color.Transforms[0].Value);
            
            chart.SetPosition(1, 0, 5, 0);

            var svg = chart.ToSvg();

            File.WriteAllText($"{_worksheetPath}svg\\EPPlusBarChart1.svg", svg);
        }


        [TestMethod]
        public void GradientPieChart()
        {
            using (var pck = OpenTemplatePackage("2.4-CreateAFileSystemReport.xlsx"))
            {
                var ws = pck.Workbook.Worksheets[1];

                int idx = 0;
                foreach(var drawing in ws.Drawings)
                {
                    var file = GetOutputFile("svg", $"{idx}_2.4-CreateAFileSystemReport.svg");
                    File.WriteAllText(file.FullName, drawing.ToSvg());
                    idx++;
                }
            }
        }

                [TestMethod]
        public void DataLabelsMultipleOneSeriesExport()
        {
            using (var pck = OpenPackage("DataLabelsMultipleOneSeriesExport.xlsx", true))
            {
                var cSheet = pck.Workbook.Worksheets.Add("ColumnChartSheet");

                var range = cSheet.Cells["A1:C3"];
                var table = cSheet.Tables.Add(range, "DataTable");
                table.ShowHeader = false;

                range.Formula = "ROW() + COLUMN()";

                cSheet.Calculate();

                var sChart = cSheet.Drawings.AddBarChart("simpleChart", eBarChartType.ColumnStacked);

                sChart.Series.Add(cSheet.Cells["A1:A3"]);
                sChart.Series.Add(cSheet.Cells["B1:B3"]);
                sChart.Series.Add(cSheet.Cells["C1:C3"]);

                sChart.Series[2].DataLabel.DataLabels.Add(0);
                var dlbl = sChart.Series[2].DataLabel.DataLabels.Add(2);
                sChart.Series[2].DataLabel.DataLabels.Add(1);


                dlbl.ShowSeriesName = true;
                dlbl.Fill.Color = Color.Red;

                var mySvg = sChart.ToSvg();

                var file = GetOutputFile("svg", "DataLabelsMultipleOneSeriesExport.svg");

                File.WriteAllText(file.FullName, mySvg);

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void Column3DChart()
        {
            var ws = _pck.Workbook.Worksheets.Add("Col3dChart");
            LoadTestdata(ws);

            var chart = ws.Drawings.AddBarChart("Col3DChart1", eBarChartType.Column3D);
            var serie = chart.Series.Add("D2:D5", "A2:A5");
            var point = serie.DataPoints.Add(0);
            point.Border.Fill.Color = Color.Blue;
            point.Border.Fill.Style = eFillStyle.SolidFill;
            point.Fill.Style = eFillStyle.SolidFill;
            point.Fill.SolidFill.Color.SetRgbColor(Color.Yellow);
            point.Fill.Transparency = 5;
            Assert.AreEqual(eColorTransformType.Alpha, point.Fill.SolidFill.Color.Transforms[0].Type);
            Assert.AreEqual(95, point.Fill.SolidFill.Color.Transforms[0].Value);
            chart.SetPosition(1, 0, 5, 0);
        }
    }
}

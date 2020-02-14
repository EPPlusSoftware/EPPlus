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
using System;

namespace EPPlusTest.Drawing.Chart
{
    [TestClass]
    public class SeriesTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("SeriesTest.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void ChartSeriesRangeAddress()
        {
            var ws = _pck.Workbook.Worksheets.Add("SeriesAddress");
            LoadTestdata(ws);
            var lineChart = ws.Drawings.AddLineChart("LineChart1", OfficeOpenXml.Drawing.Chart.eLineChartType.Line);
            var serie = lineChart.Series.Add("A1:A12", "B1:B12");
            Assert.AreEqual("'SeriesAddress'!A1:A12", serie.Series);
            Assert.AreEqual("'SeriesAddress'!B1:B12", serie.XSeries);
        }
        [TestMethod]
        public void ChartSeriesFullRangeAddress()
        {
            var ws = _pck.Workbook.Worksheets.Add("SeriesFullAddress");
            LoadTestdata(ws);
            var lineChart = ws.Drawings.AddLineChart("LineChart1", OfficeOpenXml.Drawing.Chart.eLineChartType.Line);
            var serie = lineChart.Series.Add("SeriesFullAddress!A1:A12", "SeriesFullAddress!B1:B12");
            Assert.AreEqual("SeriesFullAddress!A1:A12", serie.Series);
            Assert.AreEqual("SeriesFullAddress!B1:B12", serie.XSeries);
        }
        [TestMethod]
        public void ChartSeriesName()
        {
            var ws = _pck.Workbook.Worksheets.Add("SeriesName");
            LoadTestdata(ws);
            var lineChart = ws.Drawings.AddLineChart("LineChart1", OfficeOpenXml.Drawing.Chart.eLineChartType.Line);
            var serie = lineChart.Series.Add("SeriesName!Name1", "Name2");
            Assert.AreEqual("SeriesName!Name1", serie.Series);
            Assert.AreEqual("Name2", serie.XSeries);
        }
        [TestMethod]
        public void ChartSeriesLitStringXNumY()
        {
            var ws = _pck.Workbook.Worksheets.Add("StrLitAndNumLit");
            LoadTestdata(ws);
            var lineChart = ws.Drawings.AddLineChart("LineChart1", OfficeOpenXml.Drawing.Chart.eLineChartType.Line);
            var serie=lineChart.Series.Add("{120.3,14,5000.0005}", "{\"Label1\",\"Label 2\",\"Something else\"}");
            Assert.AreEqual(serie.StringLiteralsX.Length, 3);
            Assert.AreEqual(serie.StringLiteralsX[0], "Label1");
            Assert.AreEqual(serie.StringLiteralsX[1], "Label 2");
            Assert.AreEqual(serie.StringLiteralsX[2], "Something else");
            Assert.AreEqual(serie.NumberLiteralsY.Length, 3);
            Assert.AreEqual(serie.NumberLiteralsY[0], 120,3);
            Assert.AreEqual(serie.NumberLiteralsY[1], 14);
            Assert.AreEqual(serie.NumberLiteralsY[2], 5000.0005);

        }
        [TestMethod]
        public void ChartSeriesLitNumXNumY()
        {
            var ws = _pck.Workbook.Worksheets.Add("NumLit");
            LoadTestdata(ws);
            var lineChart = ws.Drawings.AddScatterChart("ScatterChart1", OfficeOpenXml.Drawing.Chart.eScatterChartType.XYScatter);
            var serie = lineChart.Series.Add("{120.3,14,5000.0005}", "{1.3,5,4.333}");
            Assert.AreEqual(serie.NumberLiteralsX.Length, 3);
            Assert.AreEqual(serie.NumberLiteralsX[0], 1.3);
            Assert.AreEqual(serie.NumberLiteralsX[1], 5);
            Assert.AreEqual(serie.NumberLiteralsX[2], 4.333);
            Assert.AreEqual(serie.NumberLiteralsY.Length, 3);
            Assert.AreEqual(serie.NumberLiteralsY[0], 120, 3);
            Assert.AreEqual(serie.NumberLiteralsY[1], 14);
            Assert.AreEqual(serie.NumberLiteralsY[2], 5000.0005);

        }
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void ChartSeriesLitStringYNumX()
        {
            using(var p=new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("InvalidStrLit");
                var lineChart = ws.Drawings.AddLineChart("LineChart1", OfficeOpenXml.Drawing.Chart.eLineChartType.Line);
                var serie = lineChart.Series.Add("{\"Label1\",\"Label 2\",\"Something else\"}", "{120.3,14,5000.0005}");
            }
        }
    }
}


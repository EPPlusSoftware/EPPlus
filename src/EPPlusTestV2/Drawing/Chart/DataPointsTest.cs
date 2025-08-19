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
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Style;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Drawing.Theme;
using System.Drawing;

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
        }
        [TestMethod]
        public void BarChart()  
        {
            var ws = _pck.Workbook.Worksheets.Add("BarChart");
            LoadTestdata(ws);

            var chart = ws.Drawings.AddBarChart("BarChart1", eBarChartType.Column3D);
            var serie = chart.Series.Add("D2:D5", "A2:A5");
            var point = serie.DataPoints.Add(0);
            point.Border.Fill.Color = Color.Blue;
            point.Border.Fill.Style = eFillStyle.SolidFill;
            point.Fill.Style = eFillStyle.SolidFill;
            point.Fill.SolidFill.Color.SetRgbColor(Color.Yellow);
            point.Fill.Transparancy = 5;            
            Assert.AreEqual(eColorTransformType.Alpha, point.Fill.SolidFill.Color.Transforms[0].Type);
            Assert.AreEqual(95, point.Fill.SolidFill.Color.Transforms[0].Value);
            chart.SetPosition(1, 0, 5, 0);
        }
    }
}

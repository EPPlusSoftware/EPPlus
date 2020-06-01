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
using EPPlusTest.Properties;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Style;
using System.Drawing;
using System.IO;
using System.Xml;

namespace EPPlusTest.Drawing.Chart.Styling
{
    [TestClass]
    public class ChartTemplateTests : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("ChartTemplate.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void LoadChartStyle()
        {
            var ws = _pck.Workbook.Worksheets.Add("ChartTemplate");
            LoadTestdata(ws);

            var chart=ws.Drawings.AddLineChart("LineChart1", eLineChartType.Line);
            var serie = chart.Series.Add("D2:D100", "A2:A100");

            chart.StyleManager.LoadTemplateStyles(Resources.TestLine3Crtx);
        }
        [TestMethod]
        public void AddChartFromTemplate()
        {
            var ws = _pck.Workbook.Worksheets.Add("NewChartFromTemplate");
            LoadTestdata(ws);            
            var chart = ws.Drawings.AddChartFromTemplate(Resources.TestLine3Crtx, "LineChart1", null);
            chart.Series.Add("D2:D100","A2:A100");
            chart.Series.Add("c2:c100", "A2:A100");
            chart.StyleManager.ApplyStyles();
        }
    }
}
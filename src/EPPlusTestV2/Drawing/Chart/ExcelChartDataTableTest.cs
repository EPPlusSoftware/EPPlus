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
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace EPPlusTest.Drawing.Chart
{

    [TestClass]
    public class ExcelChartDataTableTest : TestBase
    {
        /// <summary>
        /// Basic test to check output with excel. need enhanced to be stand alone checking
        /// </summary>
        [TestMethod,Ignore]
        public void DataTableFile()
        {
            string outfile = Path.Combine(_worksheetPath, "DataTableFile.xlsx");
            var fileinfo = new FileInfo(outfile);
            using (ExcelPackage pkg = new ExcelPackage(fileinfo))
            {
                // Add worksheet with sample data
                var worksheet = pkg.Workbook.Worksheets.Add("TestData");
                worksheet.Cells["A1"].Value = "Data";
                worksheet.Cells["B1"].Value = "Values";
                for (int x = 1; x < 12; ++x)
                {

                    worksheet.Cells[x + 1, 1].Value = $"Sample {x}";
                    worksheet.Cells[x + 1, 2].Value = (double)x / 3.0;
                }

                // Add chart from sample data
                var chartsheet = pkg.Workbook.Worksheets.AddChart("TestChart", eChartType.Line);
                var chart = chartsheet.Chart as ExcelLineChart;
                chart.Series.Add(worksheet.Cells["B2:B12"], worksheet.Cells["A2:A12"]).Header = "Data Test";

                Assert.AreEqual(null, chart.PlotArea.DataTable);
                chart.PlotArea.CreateDataTable();
                Assert.AreEqual(true, chart.PlotArea.DataTable.ShowVerticalBorder);
                chart.PlotArea.RemoveDataTable();
                Assert.AreEqual(null, chart.PlotArea.DataTable);
                chart.PlotArea.CreateDataTable();
                chart.PlotArea.DataTable.ShowOutline = false;
                pkg.Save();

                XmlDocument xmldoc = chart.ChartXml;
                string xml = xmldoc.InnerXml;
                Console.WriteLine(xml);
                Assert.IsTrue(xml.Contains("c:dTable"));
                Assert.IsTrue(xml.Contains("/c:dTable"));
            }
        }
    }
}

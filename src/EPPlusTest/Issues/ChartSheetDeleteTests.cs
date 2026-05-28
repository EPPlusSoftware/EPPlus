/*******************************************************************************
 * Required Notice: Copyright (C) EPPlus Software AB. 
 * https://epplussoftware.com
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *******************************************************************************/
using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace EPPlusTest.Issues
{
    [TestClass]
    public class ChartSheetDeleteTests
    {
        [TestMethod]
        public void DeleteChartSheetShouldNotThrowNotSupportedException()
        {
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Data");
                var chartSheet = package.Workbook.Worksheets.AddChart("ChartSheet", eChartType.ColumnClustered);
                
                // This call previously threw NotSupportedException because it tried to access chartSheet.PivotTables
                package.Workbook.Worksheets.Delete(chartSheet);
                
                Assert.AreEqual(1, package.Workbook.Worksheets.Count);
            }
        }
    }
}

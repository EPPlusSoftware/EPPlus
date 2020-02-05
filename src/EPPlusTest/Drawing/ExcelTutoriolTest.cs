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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace EPPlusTest.Drawing
{
    /// <summary>
    /// This test class requires the Excel tuturiol charts to be copied to the msCharts folder in the templates folder.
    /// Ignored if the folder or templates are missing.
    /// </summary>
    [TestClass]
    public class ExcelTutorialTest : TestBase
    {       
        [TestInitialize]
        public void Initialize()
        {

        }
        [TestMethod]
        public void ReadBeyondPieChartsTutorial()
        {
            using (var p = OpenTemplatePackage(@"msCharts\Beyond pie charts tutorial.xlsx"))
            {
                Assert.AreEqual(27, p.Workbook.Worksheets.Count);
                Assert.AreEqual(2, p.Workbook.Worksheets[0 + p._worksheetAdd].Drawings.Count);
                Assert.AreEqual(3, p.Workbook.Worksheets[1 + p._worksheetAdd].Drawings.Count);
                var grpShp = ((ExcelGroupShape)p.Workbook.Worksheets[1 + p._worksheetAdd].Drawings[0]);
                Assert.AreEqual(4, grpShp.Drawings.Count);
                Assert.AreEqual(eShapeStyle.Rect, ((ExcelShape)grpShp.Drawings[0]).Style);
                Assert.AreEqual(6,((ExcelPieChart)p.Workbook.Worksheets[1 + p._worksheetAdd].Drawings[1]).StyleManager.ColorsManager.Colors.Count);
            }
        }
        [TestMethod]
        public void GetMoreOutOfPivotTables()
        {
            using(var p = OpenTemplatePackage(@"msCharts\Get more out of PivotTables.xltx"))
            {
                Assert.AreEqual(26, p.Workbook.Worksheets.Count);
                Assert.AreEqual(4, p.Workbook.Worksheets[0 + p._worksheetAdd].Drawings.Count);
                Assert.AreEqual(6, p.Workbook.Worksheets[1 + p._worksheetAdd].Drawings.Count);
            }
        }
        [TestMethod]    
        public void FormulaTutorial()
        {
            using (var p = OpenTemplatePackage(@"msCharts\Formula tutorial.xltx"))
            {
                Assert.AreEqual(13, p.Workbook.Worksheets.Count);
                Assert.AreEqual(2, p.Workbook.Worksheets[0 + p._worksheetAdd].Drawings.Count);
                Assert.AreEqual(37, p.Workbook.Worksheets[1 + p._worksheetAdd].Drawings.Count);
            }
        }
    }
}

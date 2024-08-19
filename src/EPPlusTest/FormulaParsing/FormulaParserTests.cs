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
using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;
using FakeItEasy;
using OfficeOpenXml;
using System.Diagnostics;
using System.Drawing;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class FormulaParserTests
    {
        [TestInitialize]
        public void Setup()
        {

        }

        [TestCleanup]
        public void Cleanup()
        {

        }
        [TestMethod]
        public void ParseAtShouldCallExcelDataProvider()
        {            
            using (var p = new ExcelPackage())
            {
                var parser = p.Workbook.FormulaParser;
                var ws = p.Workbook.Worksheets.Add("test");
                ws.Cells["A1"].Formula = "Sum(1,2)";
                var result = parser.ParseAt("A1");
                Assert.AreEqual(3d, result);
            }
        }

        [TestMethod]
        public void Validate_shared_formula_expressions_are_cleared_when_inserting_row()
        {
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = 1D;
                ws.Cells["A2"].Value = 2D;
                ws.Cells["B1:B2"].Formula = "A1";
                package.Workbook.Calculate();
                Assert.AreEqual(1D, ws.Cells["A1"].Value);
                ws.InsertRow(1, 1);

                Assert.AreEqual(1D, ws.Cells["B2"].Value);
                ws.Cells["A2"].Value = 3D;
                package.Workbook.Calculate();
                Assert.AreEqual(3D, ws.Cells["B2"].Value);
                Assert.AreEqual(2D, ws.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void Validate_shared_formula_expressions_are_cleared_when_inserting_column()
        {
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = 1D;
                ws.Cells["B1"].Value = 2D;
                ws.Cells["A2:B2"].Formula = "A1";
                package.Workbook.Calculate();
                Assert.AreEqual(1D, ws.Cells["A1"].Value);
                ws.InsertColumn(1, 1);

                Assert.AreEqual(1D, ws.Cells["B2"].Value);
                ws.Cells["B1"].Value = 3D;
                package.Workbook.Calculate();
                Assert.AreEqual(3D, ws.Cells["B1"].Value);
                Assert.AreEqual(2D, ws.Cells["C1"].Value);
            }
        }

        [TestMethod]
        public void CalculateAfterClearFormulas()
        {
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Sheet1");
                /* 
                This reference to a custom function is a simulation of my use-case.
                It doesn't appear to matter what the formula is, it just has to be set to something
                ws.Cells["A3"].Formula = "1"; // this works just as well as "@SomeCustomVbaFunction(A1,A2)"
                */
                ws.Cells["A3"].Formula = "@SomeCustomVbaFunction(A1,A2)";
                /* 
                 * clear the formulas so that EPPlus doesn't go looking for SomeCustomVbaFunction
                 I have purposefully chosen not to implement this function as a class extending ExcelFunction                
                */
                ws.Cells["A3"].ClearFormulas();
                //ws.Cells["A3"].Formula = "0"; //This may be a workaround for now
                ws.Cells["A3"].Value = "2000"; 
                ws.Cells["A4"].Formula = "ROUNDUP(A3/1609.334,0)";

                ws.Calculate();
                Assert.AreEqual(2D, ws.Cells["A4"].Value);

            }
        }
    }
}

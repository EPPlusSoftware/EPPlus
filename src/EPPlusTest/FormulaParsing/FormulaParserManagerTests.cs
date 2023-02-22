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
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn.FunctionCompilers;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class FormulaParserManagerTests
    {
        #region test classes

        private class MyFunction : ExcelFunction
        {
            public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
            {
                throw new NotImplementedException();
            }
        }

        private class MyModule : IFunctionModule
        {
            public MyModule()
            {
                Functions = new Dictionary<string, ExcelFunction>();
                Functions.Add("MyFunction", new MyFunction());

                CustomCompilers = new Dictionary<Type, RpnFunctionCompiler>();
            }
            public IDictionary<string, ExcelFunction> Functions { get; }
            public IDictionary<Type, RpnFunctionCompiler> CustomCompilers { get; }
        }
        #endregion

        [TestMethod]
        public void FunctionsShouldBeCopied()
        {
            using (var package1 = new ExcelPackage())
            {
                package1.Workbook.FormulaParserManager.LoadFunctionModule(new MyModule());
                using (var package2 = new ExcelPackage())
                {
                    var origNumberOfFuncs = package2.Workbook.FormulaParserManager.GetImplementedFunctionNames().Count();

                    // replace functions including the custom functions from package 1
                    package2.Workbook.FormulaParserManager.CopyFunctionsFrom(package1.Workbook);

                    // Assertions: number of functions are increased with 1, and the list of function names contains the custom function.
                    Assert.AreEqual(origNumberOfFuncs + 1, package2.Workbook.FormulaParserManager.GetImplementedFunctionNames().Count());
                    Assert.IsTrue(package2.Workbook.FormulaParserManager.GetImplementedFunctionNames().Contains("myfunction"));
                }
            }
        }

        [TestMethod]
        public void ShouldParse()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A3"].Value = 2;
                var res = package.Workbook.FormulaParser.Parse("1+A3", "test!A2");
                Assert.AreEqual(3d, res);
                Assert.AreEqual(2, sheet.Cells["A3"].Value);

            }
        }

        [TestMethod]
        public void ShouldReturnCalcChain()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "SUM(A2:A3)";
                sheet.Cells["A2"].Formula = "1+2";
                sheet.Cells["A3"].Formula = "MIN(1,2)";
                var dc = package.Workbook.FormulaParserManager.GetCalculationChain(sheet.Cells["A1"]);
                Assert.AreEqual(3, dc.Count());
                Assert.AreEqual("A1", dc.Last().Address);

            }
        }
        [TestMethod]
        public void ValidateCalcChainCrossWorkSheet()
        {
            using (var package = new ExcelPackage())
            {
                var ws1 = package.Workbook.Worksheets.Add("sheet1");
                var ws2 = package.Workbook.Worksheets.Add("sheet2");
                ws1.Cells["A1"].Formula = "sheet2!A1+A2";
                ws1.Cells["A2"].Formula = "1+2";
                ws2.Cells["A1"].Formula = "1+1";
                var dc=package.Workbook.FormulaParserManager.GetCalculationChain(ws1.Cells["A1"]);
                Assert.AreEqual(3, dc.Count());

                Assert.AreEqual("sheet2", dc.ElementAt(0).Worksheet);
                Assert.AreEqual("A1", dc.ElementAt(0).Address);

                Assert.AreEqual("sheet1", dc.ElementAt(1).Worksheet);
                Assert.AreEqual("A2", dc.ElementAt(1).Address);

                Assert.AreEqual("sheet1", dc.ElementAt(2).Worksheet);
                Assert.AreEqual("A1", dc.ElementAt(2).Address);

            }
        }
        [TestMethod]
        public void ValidateCalcChainCrossWorkSheet2()
        {
            using (var package = new ExcelPackage())
            {
                var ws1 = package.Workbook.Worksheets.Add("sheet1");
                var ws2 = package.Workbook.Worksheets.Add("sheet2");
                var ws3 = package.Workbook.Worksheets.Add("sheet3");
                ws1.Cells["A1"].Formula = "1+C3";
                ws1.SetFormula(3,3, "1+1");
                ws2.Cells["A2"].Formula = "1+2";
                ws2.Cells["A1"].Formula = "1+A2";
                ws3.Cells["A1"].Formula = "sheet1!A1-A2+sheet2!A1";
                ws3.SetValue("A2", 1);
                var dc = package.Workbook.FormulaParserManager.GetCalculationChain(ws3.Cells["A1"]);
                Assert.AreEqual(5, dc.Count());

                Assert.AreEqual("sheet1", dc.ElementAt(0).Worksheet);
                Assert.AreEqual("C3", dc.ElementAt(0).Address);

                Assert.AreEqual("sheet1", dc.ElementAt(1).Worksheet);
                Assert.AreEqual("A1", dc.ElementAt(1).Address);

                Assert.AreEqual("sheet2", dc.ElementAt(2).Worksheet);
                Assert.AreEqual("A2", dc.ElementAt(2).Address);

                Assert.AreEqual("sheet2", dc.ElementAt(3).Worksheet);
                Assert.AreEqual("A1", dc.ElementAt(3).Address);

                Assert.AreEqual("sheet3", dc.ElementAt(4).Worksheet);
                Assert.AreEqual("A1", dc.ElementAt(4).Address);
            }
        }
        //[TestMethod]
        //public void ShouldFindAndParseCondFormat()
        //{
        //    var file = new FileInfo("c:\\Temp\\cf.xlsx");
        //    using (var package = new ExcelPackage(file))
        //    {
        //        var sheet = package.Workbook.Worksheets.First();

        //        // conditional formatting can only be read on sheet level
        //        // read them into a dictionary with full cell address as key
        //        var cfCells = new Dictionary<string, List<IExcelConditionalFormattingRule>>();
        //        foreach(var cf in sheet.ConditionalFormatting)
        //        {
        //            // if address is a single cell
        //            if(cf.Address.IsSingleCell)
        //            {
        //                if (!cfCells.ContainsKey(cf.Address.FullAddress)) cfCells[cf.Address.FullAddress] = new List<IExcelConditionalFormattingRule>();
        //                cfCells[cf.Address.FullAddress].Add(cf);
        //            }
        //            else
        //            {
        //                // if the address is a range with muliple cells
        //                for(var col = cf.Address.Start.Column; col <= cf.Address.End.Column; col++)
        //                {
        //                    for(var row = cf.Address.Start.Row; row <= cf.Address.End.Row; row++)
        //                    {
        //                        var adr = new ExcelAddress(row, col, row, col);
        //                        if (!cfCells.ContainsKey(adr.FullAddress)) cfCells[adr.FullAddress] = new List<IExcelConditionalFormattingRule>();
        //                        cfCells[adr.FullAddress].Add(cf);
        //                    }
        //                }
        //            }
        //        }

        //        // now check conditional formats for B2
        //        var parser = new FormulaParser(new EpplusExcelDataProvider(package));
        //        var cellAddress = new ExcelAddress(sheet.Name, "B2");
        //        var cellFormats = cfCells[cellAddress.Address];
        //        foreach(var cf in cellFormats)
        //        {
        //            var expression = cf as IExcelConditionalFormattingExpression;
        //            if(expression != null)
        //            {
        //                // evaluate the result of the conditional formats formula
        //                var result = parser.Parse(expression.Formula, cellAddress.FullAddress);
        //            }
        //        }
        //    }
        //}
    }
}

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
using OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;

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

                CustomCompilers = new Dictionary<Type, FunctionCompiler>();
            }
            public IDictionary<string, ExcelFunction> Functions { get; }
            public IDictionary<Type, FunctionCompiler> CustomCompilers { get; }
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
    }
}

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
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;
using EPPlusTest.FormulaParsing.TestHelpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml;

namespace EPPlusTest.Excel.Functions
{
    [TestClass]
    public class LogicalFunctionsTests
    {
        private ParsingContext _parsingContext = ParsingContext.Create();

        [TestMethod]
        public void IfShouldReturnCorrectResult()
        {
            var func = new If();
            var args = FunctionsHelper.CreateArgs(true, "A", "B");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual("A", result.Result);
        }

        [TestMethod, Ignore]
        public void IfShouldIgnoreCase()
        {
            using (var pck = new ExcelPackage(new FileInfo(@"c:\temp\book1.xlsx")))
            {
                pck.Workbook.Calculate();
                Assert.AreEqual("Sant", pck.Workbook.Worksheets.First().Cells["C3"].Value);
            }
        }

        [TestMethod]
        public void NotShouldReturnFalseIfArgumentIsTrue()
        {
            var func = new Not();
            var args = FunctionsHelper.CreateArgs(true);
            var result = func.Execute(args, _parsingContext);
            Assert.IsFalse((bool)result.Result);
        }

        [TestMethod]
        public void NotShouldReturnTrueIfArgumentIs0()
        {
            var func = new Not();
            var args = FunctionsHelper.CreateArgs(0);
            var result = func.Execute(args, _parsingContext);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void NotShouldReturnFalseIfArgumentIs1()
        {
            var func = new Not();
            var args = FunctionsHelper.CreateArgs(1);
            var result = func.Execute(args, _parsingContext);
            Assert.IsFalse((bool)result.Result);
        }

        [TestMethod]
        public void NotShouldHandleExcelReference()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Value = false;
                sheet.Cells["A2"].Formula = "NOT(A1)";
                sheet.Calculate();
                Assert.IsTrue((bool)sheet.Cells["A2"].Value);
            }
        }

        [TestMethod]
        public void NotShouldHandleExcelReferenceToStringFalse()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Value = "false";
                sheet.Cells["A2"].Formula = "NOT(A1)";
                sheet.Calculate();
                Assert.IsTrue((bool)sheet.Cells["A2"].Value);
            }
        }

        [TestMethod]
        public void NotShouldHandleExcelReferenceToStringTrue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Value = "TRUE";
                sheet.Cells["A2"].Formula = "NOT(A1)";
                sheet.Calculate();
                Assert.IsFalse((bool)sheet.Cells["A2"].Value);
            }
        }

        [TestMethod]
        public void AndShouldHandleStringLiteralTrue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Value = "tRuE";
                sheet.Cells["A2"].Formula = "AND(\"TRUE\", A1)";
                sheet.Calculate();
                Assert.IsTrue((bool)sheet.Cells["A2"].Value);
            }
        }

        [TestMethod]
        public void AndShouldReturnTrueIfAllArgumentsAreTrue()
        {
            var func = new And();
            var args = FunctionsHelper.CreateArgs(true, true, true);
            var result = func.Execute(args, _parsingContext);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void AndShouldReturnTrueIfAllArgumentsAreTrueOr1()
        {
            var func = new And();
            var args = FunctionsHelper.CreateArgs(true, true, 1, true, 1);
            var result = func.Execute(args, _parsingContext);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void AndShouldReturnFalseIfOneArgumentIsFalse()
        {
            var func = new And();
            var args = FunctionsHelper.CreateArgs(true, false, true);
            var result = func.Execute(args, _parsingContext);
            Assert.IsFalse((bool)result.Result);
        }

        [TestMethod]
        public void AndShouldReturnFalseIfOneArgumentIs0()
        {
            var func = new And();
            var args = FunctionsHelper.CreateArgs(true, 0, true);
            var result = func.Execute(args, _parsingContext);
            Assert.IsFalse((bool)result.Result);
        }

        [TestMethod]
        public void AndShouldHandleRange_1()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = true;
                sheet.Cells["A3"].Formula = "AND(A1:A2)";
                sheet.Calculate();
                Assert.IsTrue((bool)sheet.Cells["A3"].Value);
            }
        }

        [TestMethod]
        public void AndShouldHandleRange_2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = false;
                sheet.Cells["A3"].Formula = "AND(A1:A2)";
                sheet.Calculate();
                Assert.IsFalse((bool)sheet.Cells["A3"].Value);
            }
        }

        [TestMethod]
        public void OrShouldReturnTrueIfOneArgumentIsTrue()
        {
            var func = new Or();
            var args = FunctionsHelper.CreateArgs(true, false, false);
            var result = func.Execute(args, _parsingContext);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void OrShouldReturnTrueIfOneArgumentIsTrueString()
        {
            var func = new Or();
            var args = FunctionsHelper.CreateArgs("true", "FALSE", false);
            var result = func.Execute(args, _parsingContext);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void OrShouldHandleRange_1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 0;
                sheet.Cells["A3"].Formula = "OR(A1:A2)";
                sheet.Calculate();
                Assert.IsTrue((bool)sheet.Cells["A3"].Value);
            }
        }

        [TestMethod]
        public void OrShouldHandleRange_2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Value = 0;
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["A3"].Formula = "OR(A1:A2)";
                sheet.Calculate();
                Assert.IsTrue((bool)sheet.Cells["A3"].Value);
            }
        }

        [TestMethod]
        public void IfErrorShouldReturnSecondArgIfCriteriaEvaluatesAsAnError()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("test");
                s1.Cells["A1"].Formula = "IFERROR(0/0, \"hello\")";
                s1.Calculate();
                Assert.AreEqual("hello", s1.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void IfErrorShouldReturnSecondArgIfCriteriaEvaluatesAsAnError2()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("test");
                s1.Cells["A1"].Formula = "IFERROR(A2, \"hello\")";
                s1.Cells["A2"].Formula = "23/0";
                s1.Calculate();
                Assert.AreEqual("hello", s1.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void IfErrorShouldReturnResultOfFormulaIfNoError()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("test");
                s1.Cells["A1"].Formula = "IFERROR(A2, \"hello\")";
                s1.Cells["A2"].Value = "hi there";
                s1.Calculate();
                Assert.AreEqual("hi there", s1.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void IfNaShouldReturnSecondArgIfCriteriaEvaluatesAsAnError2()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("test");
                s1.Cells["A1"].Formula = "IFERROR(A2, \"hello\")";
                s1.Cells["A2"].Value = ExcelErrorValue.Create(eErrorType.NA);
                s1.Calculate();
                Assert.AreEqual("hello", s1.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void IfNaShouldReturnResultOfFormulaIfNoError()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("test");
                s1.Cells["A1"].Formula = "IFNA(A2, \"hello\")";
                s1.Cells["A2"].Value = "hi there";
                s1.Calculate();
                Assert.AreEqual("hi there", s1.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void SwitchShouldReturnFirstMatchingArg()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("test");
                s1.Cells["A1"].Formula = "SWITCH(A2, 1, 2)";
                s1.Cells["A2"].Value = 1;
                s1.Calculate();
                Assert.AreEqual(2d, s1.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void SwitchShouldIgnoreNonMatchingArg()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("test");
                s1.Cells["A1"].Formula = "SWITCH(A2, 1, 2, B2, 3)";
                s1.Cells["A2"].Value = 2;
                s1.Cells["B2"].Value = 2d;
                s1.Calculate();
                Assert.AreEqual(3d, s1.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void SwitchShouldReturnLastArgIfNoMatch()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("test");
                s1.Cells["A1"].Formula = "SWITCH(A2, 1, 2, B2, 3, 5)";
                s1.Cells["A2"].Value = -1;
                s1.Cells["B2"].Value = 2d;
                s1.Calculate();
                Assert.AreEqual(5d, s1.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void XorShouldReturnCorrectResult()
        {
            var func = new Xor();
            var args = FunctionsHelper.CreateArgs(true, false);
            var result = func.Execute(args, ParsingContext.Create());
            Assert.IsTrue((bool)result.Result);

            args = FunctionsHelper.CreateArgs(false, false);
            result = func.Execute(args, ParsingContext.Create());
            Assert.IsFalse((bool)result.Result);

            args = FunctionsHelper.CreateArgs(true, true);
            result = func.Execute(args, ParsingContext.Create());
            Assert.IsFalse((bool)result.Result);

            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = true;
                sheet.Cells["A2"].Value = 0;
                sheet.Cells["A3"].Formula = "XOR(A1:A2,DATE(2020,12,10))";
                sheet.Calculate();
                Assert.IsFalse((bool)sheet.Cells["A3"].Value);
            }
        }
    }
}

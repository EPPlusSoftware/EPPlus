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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using EPPlusTest.FormulaParsing.TestHelpers;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml;
using System.Globalization;

namespace EPPlusTest.Excel.Functions.Text
{
    [TestClass]
    public class TextFunctionsTests
    {
        private ParsingContext _parsingContext = ParsingContext.Create();

        [TestMethod]
        public void CStrShouldConvertNumberToString()
        {
            var func = new CStr();
            var result = func.Execute(FunctionsHelper.CreateArgs(1), _parsingContext);
            Assert.AreEqual(DataType.String, result.DataType);
            Assert.AreEqual("1", result.Result);
        }

        [TestMethod]
        public void LenShouldReturnStringsLength()
        {
            var func = new Len();
            var result = func.Execute(FunctionsHelper.CreateArgs("abc"), _parsingContext);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod]
        public void LowerShouldReturnLowerCaseString()
        {
            var func = new Lower();
            var result = func.Execute(FunctionsHelper.CreateArgs("ABC"), _parsingContext);
            Assert.AreEqual("abc", result.Result);
        }

        [TestMethod]
        public void UpperShouldReturnUpperCaseString()
        {
            var func = new Upper();
            var result = func.Execute(FunctionsHelper.CreateArgs("abc"), _parsingContext);
            Assert.AreEqual("ABC", result.Result);
        }

        [TestMethod]
        public void LeftShouldReturnSubstringFromLeft()
        {
            var func = new Left();
            var result = func.Execute(FunctionsHelper.CreateArgs("abcd", 2), _parsingContext);
            Assert.AreEqual("ab", result.Result);
        }

        [TestMethod]
        public void RightShouldReturnSubstringFromRight()
        {
            var func = new Right();
            var result = func.Execute(FunctionsHelper.CreateArgs("abcd", 2), _parsingContext);
            Assert.AreEqual("cd", result.Result);
        }

        [TestMethod]
        public void MidShouldReturnSubstringAccordingToParams()
        {
            var func = new Mid();
            var result = func.Execute(FunctionsHelper.CreateArgs("abcd", 1, 2), _parsingContext);
            Assert.AreEqual("ab", result.Result);
        }

        [TestMethod]
        public void ReplaceShouldReturnAReplacedStringAccordingToParamsWhenStartIxIs1()
        {
            var func = new Replace();
            var result = func.Execute(FunctionsHelper.CreateArgs("testar", 1, 2, "hej"), _parsingContext);
            Assert.AreEqual("hejstar", result.Result);
        }

        [TestMethod]
        public void ReplaceShouldReturnAReplacedStringAccordingToParamsWhenStartIxIs3()
        {
            var func = new Replace();
            var result = func.Execute(FunctionsHelper.CreateArgs("testar", 3, 3, "hej"), _parsingContext);
            Assert.AreEqual("tehejr", result.Result);
        }

        [TestMethod]
        public void SubstituteShouldReturnAReplacedStringAccordingToParamsWhen()
        {
            var func = new Substitute();
            var result = func.Execute(FunctionsHelper.CreateArgs("testar testar", "es", "xx"), _parsingContext);
            Assert.AreEqual("txxtar txxtar", result.Result);
        }

        [TestMethod]
        public void ConcatenateShouldConcatenateThreeStrings()
        {
            var func = new Concatenate();
            var result = func.Execute(FunctionsHelper.CreateArgs("One", "Two", "Three"), _parsingContext);
            Assert.AreEqual("OneTwoThree", result.Result);
        }

        [TestMethod]
        public void ConcatenateShouldConcatenateStringWithInt()
        {
            var func = new Concatenate();
            var result = func.Execute(FunctionsHelper.CreateArgs(1, "Two"), _parsingContext);
            Assert.AreEqual("1Two", result.Result);
        }

        [TestMethod]
        public void ConcatShouldReturnValErrorIfMoreThan254Args()
        {
            var func = new Concat();
            List<object> args = new List<object>();
            for(var i = 0; i < 255;  i++)
            {
                args.Add("arg " + i);
            }
            var result = func.Execute(FunctionsHelper.CreateArgs(args.ToArray()), _parsingContext);
            Assert.AreEqual("#VALUE!", result.Result.ToString());
        }

        [TestMethod]
        public void ExactShouldReturnTrueWhenTwoEqualStrings()
        {
            var func = new Exact();
            var result = func.Execute(FunctionsHelper.CreateArgs("abc", "abc"), _parsingContext);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void ExactShouldReturnTrueWhenEqualStringAndDouble()
        {
            var func = new Exact();
            var result = func.Execute(FunctionsHelper.CreateArgs("1", 1d), _parsingContext);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void ExactShouldReturnFalseWhenStringAndNull()
        {
            var func = new Exact();
            var result = func.Execute(FunctionsHelper.CreateArgs("1", null), _parsingContext);
            Assert.IsFalse((bool)result.Result);
        }

        [TestMethod]
        public void ExactShouldReturnFalseWhenTwoEqualStringsWithDifferentCase()
        {
            var func = new Exact();
            var result = func.Execute(FunctionsHelper.CreateArgs("abc", "Abc"), _parsingContext);
            Assert.IsFalse((bool)result.Result);
        }

        [TestMethod]
        public void FindShouldReturnIndexOfFoundPhrase()
        {
            var func = new Find();
            var result = func.Execute(FunctionsHelper.CreateArgs("hopp", "hej hopp"), _parsingContext);
            Assert.AreEqual(5, result.Result);
        }

        [TestMethod]
        public void FindShouldReturnIndexOfFoundPhraseBasedOnStartIndex()
        {
            var func = new Find();
            var result = func.Execute(FunctionsHelper.CreateArgs("hopp", "hopp hopp", 2), _parsingContext);
            Assert.AreEqual(6, result.Result);
        }

        [TestMethod]
        public void ProperShouldSetFirstLetterToUpperCase()
        {
            var func = new Proper();
            var result = func.Execute(FunctionsHelper.CreateArgs("this IS A tEst.wi3th SOME w0rds östEr"), _parsingContext);
            Assert.AreEqual("This Is A Test.Wi3Th Some W0Rds Öster", result.Result);
        }

        [TestMethod]
        public void HyperLinkShouldReturnArgIfOneArgIsSupplied()
        {
            var func = new Hyperlink();
            var result = func.Execute(FunctionsHelper.CreateArgs("http://epplus.codeplex.com"), _parsingContext);
            Assert.AreEqual("http://epplus.codeplex.com", result.Result);
        }

        [TestMethod]
        public void HyperLinkShouldReturnLastArgIfTwoArgsAreSupplied()
        {
            var func = new Hyperlink();
            var result = func.Execute(FunctionsHelper.CreateArgs("http://epplus.codeplex.com", "EPPlus"), _parsingContext);
            Assert.AreEqual("EPPlus", result.Result);
        }

        [TestMethod]
        public void TrimShouldReturnDataTypeString()
        {
            var func = new Trim();
            var result = func.Execute(FunctionsHelper.CreateArgs(" epplus "), _parsingContext);
            Assert.AreEqual(DataType.String, result.DataType);
        }

        [TestMethod]
        public void TrimShouldTrimFromBothEnds()
        {
            var func = new Trim();
            var result = func.Execute(FunctionsHelper.CreateArgs(" epplus "), _parsingContext);
            Assert.AreEqual("epplus", result.Result);
        }

        [TestMethod]
        public void TrimShouldTrimMultipleSpaces()
        {
            var func = new Trim();
            var result = func.Execute(FunctionsHelper.CreateArgs(" epplus    5 "), _parsingContext);
            Assert.AreEqual("epplus 5", result.Result);
        }

        [TestMethod]
        public void CleanShouldReturnDataTypeString()
        {
            var func = new Clean();
            var result = func.Execute(FunctionsHelper.CreateArgs("epplus"), _parsingContext);
            Assert.AreEqual(DataType.String, result.DataType);
        }

        [TestMethod]
        public void CleanShouldRemoveNonPrintableChars()
        {
            var input = new StringBuilder();
            for(var x = 1; x < 32; x++)
            {
                input.Append((char)x);
            }
            input.Append("epplus");
            var func = new Clean();
            var result = func.Execute(FunctionsHelper.CreateArgs(input), _parsingContext);
            Assert.AreEqual("epplus", result.Result);
        }

        [TestMethod]
        public void UnicodeShouldReturnCorrectCode()
        {
            var func = new Unicode();
            
            var result = func.Execute(FunctionsHelper.CreateArgs("B"), _parsingContext);
            Assert.AreEqual(66, result.Result);

            result = func.Execute(FunctionsHelper.CreateArgs("a"), _parsingContext);
            Assert.AreEqual(97, result.Result);
        }

        [TestMethod]
        public void UnicharShouldReturnCorrectChar()
        {
            var func = new Unichar();

            var result = func.Execute(FunctionsHelper.CreateArgs(66), _parsingContext);
            Assert.AreEqual("B", result.Result);

            result = func.Execute(FunctionsHelper.CreateArgs(97), _parsingContext);
            Assert.AreEqual("a", result.Result);
        }

        [TestMethod]
        public void NumberValueShouldCastIntegerValue()
        {
            var func = new NumberValue();
            var result = func.Execute(FunctionsHelper.CreateArgs("1000"), _parsingContext);
            Assert.AreEqual(1000d, result.Result);
        }

        [TestMethod]
        public void NumberValueShouldCastDecinalValueWithCurrentCulture()
        {
            var input = $"1{CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator}000{CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator}15";
            var func = new NumberValue();
            var result = func.Execute(FunctionsHelper.CreateArgs(input), _parsingContext);
            Assert.AreEqual(1000.15d, result.Result);
        }

        [TestMethod]
        public void NumberValueShouldCastDecinalValueWithSeparators()
        {
            var input = $"1,000.15";
            var func = new NumberValue();
            var result = func.Execute(FunctionsHelper.CreateArgs(input, ".", ","), _parsingContext);
            Assert.AreEqual(1000.15d, result.Result);
        }

        [TestMethod]
        public void NumberValueShouldHandlePercentage()
        {
            var input = $"1,000.15%";
            var func = new NumberValue();
            var result = func.Execute(FunctionsHelper.CreateArgs(input, ".", ","), _parsingContext);
            Assert.AreEqual(10.0015d, result.Result);
        }

        [TestMethod]
        public void NumberValueShouldHandleMultiplePercentage()
        {
            var input = $"1,000.15%%";
            var func = new NumberValue();
            var result = func.Execute(FunctionsHelper.CreateArgs(input, ".", ","), _parsingContext);
            var r = System.Math.Round((double)result.Result, 15);
            Assert.AreEqual(0.100015d, r);
        }

        [TestMethod]
        public void TextjoinShouldReturnCorrectResult_IgnoreEmpty()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "Hello";
                sheet.Cells["A2"].Value = "world";
                sheet.Cells["A3"].Value = "";
                sheet.Cells["A4"].Value = "!";
                sheet.Cells["A5"].Formula = "TEXTJOIN(\" \", TRUE, A1:A4)";
                sheet.Calculate();
                Assert.AreEqual("Hello world !", sheet.Cells["A5"].Value);
            }
        }

        [TestMethod]
        public void TextjoinShouldReturnCorrectResult_AllowEmpty()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "Hello";
                sheet.Cells["A2"].Value = "world";
                sheet.Cells["A3"].Value = "";
                sheet.Cells["A4"].Value = "!";
                sheet.Cells["A5"].Formula = "TEXTJOIN(\".\", False, A1:A4, \"how are you?\")";
                sheet.Calculate();
                Assert.AreEqual("Hello.world..!.how are you?", sheet.Cells["A5"].Value);
            }
        }

        [TestMethod]
        public void DollarShouldReturnCorrectResult()
        {
            var expected = 123.46.ToString("C", CultureInfo.CurrentCulture);
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 123.456;
                sheet.Cells["A2"].Formula = "DOLLAR(A1)";
                sheet.Calculate();
                Assert.AreEqual(expected, sheet.Cells["A2"].Value);

                expected = 123.5.ToString("C1", CultureInfo.CurrentCulture);
                sheet.Cells["A2"].Formula = "DOLLAR(A1, 1)";
                sheet.Calculate();
                Assert.AreEqual(expected, sheet.Cells["A2"].Value);

                expected = 123.ToString("C0", CultureInfo.CurrentCulture);
                sheet.Cells["A2"].Formula = "DOLLAR(A1, 0)";
                sheet.Calculate();
                Assert.AreEqual(expected, sheet.Cells["A2"].Value);

                expected = 120.ToString("C0", CultureInfo.CurrentCulture);
                sheet.Cells["A2"].Formula = "DOLLAR(A1, -1)";
                sheet.Calculate();
                Assert.AreEqual(expected, sheet.Cells["A2"].Value);
            }
        }

        [TestMethod]
        public void ValueShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "1,234,567.89";
                sheet.Cells["A2"].Formula = "VALUE(A1)";
                sheet.Calculate();
                Assert.AreEqual(1234567.89, sheet.Cells["A2"].Value);
            }
        }
    }
}

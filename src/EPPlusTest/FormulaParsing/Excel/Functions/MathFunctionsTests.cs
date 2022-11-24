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
using System.Threading;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using EPPlusTest.FormulaParsing.TestHelpers;
using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace EPPlusTest.Excel.Functions
{
    [TestClass]
    public class MathFunctionsTests
    {
        private ParsingContext _parsingContext;

        [TestInitialize]
        public void Initialize()
        {
            _parsingContext = ParsingContext.Create();
            _parsingContext.Scopes.NewScope(FormulaRangeAddress.Empty);
        }

        [TestMethod]
        public void PiShouldReturnPIConstant()
        {
            var expectedValue = (double)Math.Round(Math.PI, 14);
            var func = new Pi();
            var args = FunctionsHelper.CreateArgs(0);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void AbsShouldReturnCorrectResult()
        {
            var expectedValue = 3d;
            var func = new Abs();
            var args = FunctionsHelper.CreateArgs(-3d);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void AsinShouldReturnCorrectResult()
        {
            const double expectedValue = 1.5708;
            var func = new Asin();
            var args = FunctionsHelper.CreateArgs(1d);
            var result = func.Execute(args, _parsingContext);
            var rounded = Math.Round((double)result.Result, 4);
            Assert.AreEqual(expectedValue, rounded);
        }

        [TestMethod]
        public void AsinhShouldReturnCorrectResult()
        {
            const double expectedValue = 0.0998;
            var func = new Asinh();
            var args = FunctionsHelper.CreateArgs(0.1d);
            var result = func.Execute(args, _parsingContext);
            var rounded = Math.Round((double)result.Result, 4);
            Assert.AreEqual(expectedValue, rounded);
        }

        [TestMethod]
        public void CombinShouldReturnCorrectResult_6_2()
        {
            var func = new Combin();

            var args = FunctionsHelper.CreateArgs(6, 2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(15d, result.Result);
        }

        [TestMethod]
        public void CombinShouldReturnCorrectResult_decimal()
        {
            var func = new Combin();

            var args = FunctionsHelper.CreateArgs(10.456, 3);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(120d, result.Result);
        }

        [TestMethod]
        public void CombinShouldReturnCorrectResult_6_1()
        {
            var func = new Combin();

            var args = FunctionsHelper.CreateArgs(6, 1);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(6d, result.Result);
        }

        [TestMethod]
        public void CombinaShouldReturnCorrectResult_6_2()
        {
            var func = new Combina();

            var args = FunctionsHelper.CreateArgs(6, 2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(21d, result.Result);
        }

        [TestMethod]
        public void CombinaShouldReturnCorrectResult_6_5()
        {
            var func = new Combina();

            var args = FunctionsHelper.CreateArgs(6, 5);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(252d, result.Result);
        }

        [TestMethod]
        public void PermutationaShouldReturnCorrectResult()
        {
            var func = new Permutationa();

            var args = FunctionsHelper.CreateArgs(6, 6);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(46656d, result.Result);

            args = FunctionsHelper.CreateArgs(10, 6);
            result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1000000d, result.Result);
        }

        [TestMethod]
        public void PermutShouldReturnCorrectResult()
        {
            var func = new Permut();

            var args = FunctionsHelper.CreateArgs(6, 6);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(720d, result.Result);

            args = FunctionsHelper.CreateArgs(10, 6);
            result = func.Execute(args, _parsingContext);
            Assert.AreEqual(151200d, result.Result);
        }

        [TestMethod]
        public void SecShouldReturnCorrectResult_MinusPi()
        {
            var func = new Sec();
            var args = FunctionsHelper.CreateArgs(-3.14159265358979);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(-1d, result.Result);
        }

        [TestMethod]
        public void SecShouldReturnCorrectResult_Zero()
        {
            var func = new Sec();
            var args = FunctionsHelper.CreateArgs(0d);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void SechShouldReturnCorrectResult_PiDividedBy4()
        {
            var func = new SecH();
            var args = FunctionsHelper.CreateArgs(Math.PI / 4);
            var result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(0.7549, Math.Round((double)result, 4));
        }

        [TestMethod]
        public void SechShouldReturnCorrectResult_MinusPi()
        {
            var func = new SecH();
            var args = FunctionsHelper.CreateArgs(-3.14159265358979);
            var result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(0.08627, Math.Round((double)result, 5));
        }

        [TestMethod]
        public void SechShouldReturnCorrectResult_Zero()
        {
            var func = new SecH();
            var args = FunctionsHelper.CreateArgs(0d);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void SecShouldReturnCorrectResult_PiDividedBy4()
        {
            var func = new Sec();
            var args = FunctionsHelper.CreateArgs(Math.PI / 4);
            var result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(1.4142, Math.Round((double)result, 4));
        }

        [TestMethod]
        public void CscShouldReturnCorrectResult_Minus6()
        {
            var func = new Csc();
            var args = FunctionsHelper.CreateArgs(-6);
            var result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(3.5789, Math.Round((double)result, 4));
        }

        [TestMethod]
        public void CotShouldReturnCorrectResult_2()
        {
            var func = new Cot();
            var args = FunctionsHelper.CreateArgs(2);
            var result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(-0.4577, Math.Round((double)result, 4));
        }

        [TestMethod]
        public void CothShouldReturnCorrectResult_MinusPi()
        {
            var func = new Coth();
            var args = FunctionsHelper.CreateArgs(Math.PI * -1);
            var result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(-1.0037, Math.Round((double)result, 4));
        }

        [TestMethod]
        public void AcothShouldReturnCorrectResult_MinusPi()
        {
            var func = new Acoth();
            var args = FunctionsHelper.CreateArgs(-5);
            var result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(-0.2027, Math.Round((double)result, 4));
        }

        [TestMethod]
        public void RadiansShouldReturnCorrectResult_50()
        {
            var func = new Radians();
            var args = FunctionsHelper.CreateArgs(50);
            var result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(0.8727, Math.Round((double)result, 4));
        }

        [TestMethod]
        public void RadiansShouldReturnCorrectResult_360()
        {
            var func = new Radians();
            var args = FunctionsHelper.CreateArgs(360);
            var result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(6.2832, Math.Round((double)result, 4));
        }

        [TestMethod]
        public void AcotShouldReturnCorrectResult_1()
        {
            var func = new Acot();
            var args = FunctionsHelper.CreateArgs(1);
            var result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(0.7854, Math.Round((double)result, 4));
        }

        [TestMethod]
        public void CschShouldReturnCorrectResult_Pi()
        {
            var func = new Csch();
            var args = FunctionsHelper.CreateArgs(Math.PI * -1);
            var result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(-0.0866, Math.Round((double)result, 4));
        }

        [TestMethod]
        public void RomanShouldReturnCorrectResult()
        {
            var func = new Roman();

            var args = FunctionsHelper.CreateArgs(2);
            var result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("II", result, "2 was not II");

            args = FunctionsHelper.CreateArgs(4);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("IV", result, "4 was not IV");

            args = FunctionsHelper.CreateArgs(14);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("XIV", result, "14 was not XIV");

            args = FunctionsHelper.CreateArgs(23);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("XXIII", result, "23 was not XXIII");

            args = FunctionsHelper.CreateArgs(59);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("LIX", result, "59 was not LIX");

            args = FunctionsHelper.CreateArgs(99);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("XCIX", result, "99 was not XCIX");

            args = FunctionsHelper.CreateArgs(412);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("CDXII", result, "412 was not CDXII");

            args = FunctionsHelper.CreateArgs(1214);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("MCCXIV", result, "1214 was not MCCXIV");

            args = FunctionsHelper.CreateArgs(3295);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("MMMCCXCV", result, "3295 was not MMMCCXCV");
        }

        [TestMethod]
        public void RomanType1ShouldReturnCorrectResult()
        {
            var func = new Roman();

            var args = FunctionsHelper.CreateArgs(495, 1);
            var result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("LDVL", result, "495 was not LDVL");

            args = FunctionsHelper.CreateArgs(45, 1);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("VL", result, "45 was not VL");

            args = FunctionsHelper.CreateArgs(49, 1);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("VLIV", result, "59 was not VLIV");

            args = FunctionsHelper.CreateArgs(99, 1);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("VCIV", result, "99 was not VCIV");

            args = FunctionsHelper.CreateArgs(395, 1);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("CCCVC", result, "395 was not CCCVC");

            args = FunctionsHelper.CreateArgs(949, 1);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("CMVLIV", result, "949 was not CMVLIV");

            args = FunctionsHelper.CreateArgs(3295, 1);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("MMMCCVC", result, "3295 was not MMMCCVC");
        }

        [TestMethod]
        public void RomanType2ShouldReturnCorrectResult()
        {
            var func = new Roman();

            var args = FunctionsHelper.CreateArgs(495, 2);
            var result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("XDV", result, "495 was not XDV");

            args = FunctionsHelper.CreateArgs(45, 2);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("VL", result, "45 was not VL");

            args = FunctionsHelper.CreateArgs(59, 2);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("LIX", result, "59 was not LIX");

            args = FunctionsHelper.CreateArgs(99, 2);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("IC", result, "99 was not IC");

            args = FunctionsHelper.CreateArgs(490, 2);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("XD", result, "490 was not XD");

            args = FunctionsHelper.CreateArgs(949, 2);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("CMIL", result, "949 was not CMIL");

            args = FunctionsHelper.CreateArgs(2999, 2);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("MMXMIX", result, "2999 was not MMXMIX");
        }

        [TestMethod]
        public void RomanType3ShouldReturnCorrectResult()
        {
            var func = new Roman();

            var args = FunctionsHelper.CreateArgs(495, 3);
            var result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("VD", result, "495 was not VD");

            args = FunctionsHelper.CreateArgs(499, 3);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("VDIV", result, "499 was not VDIV");

            args = FunctionsHelper.CreateArgs(995, 3);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("VM", result, "995 was not VM");

            args = FunctionsHelper.CreateArgs(999, 3);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("VMIV", result, "999 was not VMIV");

            args = FunctionsHelper.CreateArgs(1999, 3);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("MVMIV", result, "490 was not MVMIV");
        }

        [TestMethod]
        public void RomanType4ShouldReturnCorrectResult()
        {
            var func = new Roman();

            var args = FunctionsHelper.CreateArgs(499, 4);
            var result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("ID", result, "499 was not ID");

            args = FunctionsHelper.CreateArgs(999, 4);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual("IM", result, "999 was not IM");
        }

        [TestMethod]
        public void GcdShouldReturnCorrectResult()
        {
            var func = new Gcd();

            var args = FunctionsHelper.CreateArgs(15, 10, 25);
            var result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(5, result);

            args = FunctionsHelper.CreateArgs(0, 8, 12);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(4, result);
        }

        [TestMethod]
        public void LcmShouldReturnCorrectResult()
        {
            var func = new Lcm();

            var args = FunctionsHelper.CreateArgs(15, 10, 25);
            var result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(150, result);

            args = FunctionsHelper.CreateArgs(1, 8, 12);
            result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(24, result);
        }

        [TestMethod]
        public void SumShouldCalculate2Plus3AndReturn5()
        {
            var func = new Sum();
            var args = FunctionsHelper.CreateArgs(2, 3);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void SumShouldCalculateEnumerableOf2Plus5Plus3AndReturn10()
        {
            var func = new Sum();
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(2, 5), 3);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(10d, result.Result);
        }

        [TestMethod]
        public void SumShouldIgnoreHiddenValuesWhenIgnoreHiddenValuesIsSet()
        {
            var func = new Sum();
            func.IgnoreHiddenValues = true;
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(2, 5), 3, 4);
            args.Last().SetExcelStateFlag(ExcelCellState.HiddenCell);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(10d, result.Result);
        }

        [TestMethod]
        public void SumSqShouldCalculateArray()
        {
            var func = new Sumsq();
            var args = FunctionsHelper.CreateArgs(2, 4);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(20d, result.Result);
        }

        [TestMethod]
        public void SumSqShouldIncludeTrueAsOne()
        {
            var func = new Sumsq();
            var args = FunctionsHelper.CreateArgs(2, 4, true);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(21d, result.Result);
        }

        [TestMethod]
        public void SumSqShouldNoCountTrueTrueInArray()
        {
            var func = new Sumsq();
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(2, 4, true));
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(20d, result.Result);
        }

        [TestMethod]
        public void StdevShouldCalculateCorrectResult()
        {
            var func = new Stdev();
            var args = FunctionsHelper.CreateArgs(1, 3, 5);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        [TestMethod]
        public void StdevaShouldCalculateCorrectResult()
        {
            var func = new Stdeva();
            var args = FunctionsHelper.CreateArgs(1, 3, 5, 2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1.7078d, System.Math.Round((double)result.Result, 4));

            args = FunctionsHelper.CreateArgs(1, 3, 5, 2, true, "text");
            result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1.7889d, System.Math.Round((double)result.Result, 4));
        }

        [TestMethod]
        public void StdevpaShouldCalculateCorrectResult()
        {
            var func = new Stdevpa();
            var args = FunctionsHelper.CreateArgs(1, 3, 5, 2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1.479d, System.Math.Round((double)result.Result, 4));

            args = FunctionsHelper.CreateArgs(1, 3, 5, 2, true, "text");
            result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1.633d, System.Math.Round((double)result.Result, 4));
        }

        [TestMethod]
        public void StdevShouldIgnoreHiddenValuesWhenIgnoreHiddenValuesIsSet()
        {
            var func = new Stdev();
            func.IgnoreHiddenValues = true;
            var args = FunctionsHelper.CreateArgs(1, 3, 5, 6);
            args.Last().SetExcelStateFlag(ExcelCellState.HiddenCell);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        [TestMethod]
        public void StdevPShouldCalculateCorrectResult()
        {
            var func = new StdevP();
            var args = FunctionsHelper.CreateArgs(2, 3, 4);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(0.8165d, Math.Round((double)result.Result, 5));
        }

        [TestMethod]
        public void StdevPShouldIgnoreHiddenValuesWhenIgnoreHiddenValuesIsSet()
        {
            var func = new StdevP();
            func.IgnoreHiddenValues = true;
            var args = FunctionsHelper.CreateArgs(2, 3, 4, 165);
            args.Last().SetExcelStateFlag(ExcelCellState.HiddenCell);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(0.8165d, Math.Round((double)result.Result, 5));
        }

        [TestMethod]
        public void ExpShouldCalculateCorrectResult()
        {
            var func = new Exp();
            var args = FunctionsHelper.CreateArgs(4);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(54.59815003d, System.Math.Round((double)result.Result, 8));
        }

        [TestMethod]
        public void MaxShouldCalculateCorrectResult()
        {
            var func = new Max();
            var args = FunctionsHelper.CreateArgs(4, 2, 5, 2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void MaxShouldIgnoreHiddenValuesIfIgnoreHiddenValuesIsTrue()
        {
            var func = new Max();
            func.IgnoreHiddenValues = true;
            var args = FunctionsHelper.CreateArgs(4, 2, 5, 2);
            args.ElementAt(2).SetExcelStateFlag(ExcelCellState.HiddenCell);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(4d, result.Result);
        }

        [TestMethod]
        public void MaxShouldHandleEmptyRange()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A5"].Formula = "MAX(A1:A4)";
                sheet.Calculate();
                var value = sheet.Cells["A5"].Value;
                Assert.AreEqual(0d, value);
            }
        }

        [TestMethod]
        public void MaxaShouldCalculateCorrectResult()
        {
            var func = new Maxa();
            var args = FunctionsHelper.CreateArgs(-1, 0, 1);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void MaxaShouldCalculateCorrectResultUsingBool()
        {
            var func = new Maxa();
            var args = FunctionsHelper.CreateArgs(-1, 0, true);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void MaxaShouldCalculateCorrectResultUsingString()
        {
            var func = new Maxa();
            var args = FunctionsHelper.CreateArgs(-1, "test");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(0d, result.Result);
        }

        [TestMethod]
        public void MinShouldCalculateCorrectResult()
        {
            var func = new Min();
            var args = FunctionsHelper.CreateArgs(4, 2, 5, 2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        [TestMethod]
        public void MinShouldIgnoreHiddenValuesIfIgnoreHiddenValuesIsTrue()
        {
            var func = new Min();
            func.IgnoreHiddenValues = true;
            var args = FunctionsHelper.CreateArgs(4, 2, 5, 3);
            args.ElementAt(1).SetExcelStateFlag(ExcelCellState.HiddenCell);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod]
        public void MinShouldHandleEmptyRange()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A5"].Formula = "MIN(A1:A4)";
                sheet.Calculate();
                var value = sheet.Cells["A5"].Value;
                Assert.AreEqual(0d, value);
            }
        }

        [TestMethod]
        public void AverageShouldCalculateCorrectResult()
        {
            var expectedResult = (4d + 2d + 5d + 2d) / 4d;
            var func = new Average();
            var args = FunctionsHelper.CreateArgs(4d, 2d, 5d, 2d);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedResult, result.Result);
        }

        [TestMethod]
        public void AverageShouldCalculateCorrectResultWithEnumerableAndBoolMembers()
        {
            var expectedResult = (4d + 2d + 5d + 2d + 1d) / 5d;
            var func = new Average();
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(4d, 2d), 5d, 2d, true);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedResult, result.Result);
        }

        [TestMethod]
        public void AverageShouldIgnoreHiddenFieldsIfIgnoreHiddenValuesIsTrue()
        {
            var expectedResult = (4d + 2d + 2d + 1d) / 4d;
            var func = new Average();
            func.IgnoreHiddenValues = true;
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(4d, 2d), 5d, 2d, true);
            args.ElementAt(1).SetExcelStateFlag(ExcelCellState.HiddenCell);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedResult, result.Result);
        }

        [TestMethod]
        public void AverageShouldThrowDivByZeroExcelErrorValueIfEmptyArgs()
        {
            eErrorType errorType = eErrorType.Value;

            var func = new Average();
            var args = new FunctionArgument[0];
            try
            {
                func.Execute(args, _parsingContext);
            }
            catch (ExcelErrorValueException e)
            {
                errorType = e.ErrorValue.Type;
            }
            Assert.AreEqual(eErrorType.Div0, errorType);
        }

        [TestMethod]
        public void AverageAShouldCalculateCorrectResult()
        {
            var expectedResult = (4d + 2d + 5d + 2d) / 4d;
            var func = new AverageA();
            var args = FunctionsHelper.CreateArgs(4d, 2d, 5d, 2d);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedResult, result.Result);
        }

        [TestMethod]
        public void AverageAShouldIncludeTrueAs1()
        {
            var expectedResult = (4d + 2d + 5d + 2d + 1d) / 5d;
            var func = new AverageA();
            var args = FunctionsHelper.CreateArgs(4d, 2d, 5d, 2d, true);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedResult, result.Result);
        }

        [TestMethod, ExpectedException(typeof(ExcelErrorValueException))]
        public void AverageAShouldThrowValueExceptionIfNonNumericTextIsSupplied()
        {
            var func = new AverageA();
            var args = FunctionsHelper.CreateArgs(4d, 2d, 5d, 2d, "ABC");
            var result = func.Execute(args, _parsingContext);
        }

        [TestMethod]
        public void AverageAShouldCountValueAs0IfNonNumericTextIsSuppliedInArray()
        {
            var func = new AverageA();
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(1d, 2d, 3d, "ABC"));
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1.5d, result.Result);
        }

        [TestMethod]
        public void AverageAShouldCountNumericStringWithValue()
        {
            var func = new AverageA();
            var args = FunctionsHelper.CreateArgs(4d, 2d, "9");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void RoundShouldReturnCorrectResult()
        {
            var func = new Round();
            var args = FunctionsHelper.CreateArgs(2.3433, 3);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2.343d, result.Result);
        }

        [TestMethod]
        public void RoundShouldReturnCorrectResultWhenNbrOfDecimalsIsNegative()
        {
            var func = new Round();
            var args = FunctionsHelper.CreateArgs(9333, -3);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(9000d, result.Result);
        }

        [TestMethod]
        public void RandShouldReturnAValueBetween0and1()
        {
            var func = new Rand();
            var args = new FunctionArgument[0];
            var result1 = func.Execute(args, _parsingContext);
            Assert.IsTrue(((double)result1.Result) > 0 && ((double)result1.Result) < 1);
            var result2 = func.Execute(args, _parsingContext);
            Assert.AreNotEqual(result1.Result, result2.Result, "The two numbers were the same");
            Assert.IsTrue(((double)result2.Result) > 0 && ((double)result2.Result) < 1);
        }

        [TestMethod]
        public void RandBetweenShouldReturnAnIntegerValueBetweenSuppliedValues()
        {
            var func = new RandBetween();
            var args = FunctionsHelper.CreateArgs(1, 5);
            var result = func.Execute(args, _parsingContext);
            CollectionAssert.Contains(new List<double> { 1d, 2d, 3d, 4d, 5d }, result.Result);
        }

        [TestMethod]
        public void RandBetweenShouldReturnAnIntegerValueBetweenSuppliedValuesWhenLowIsNegative()
        {
            var func = new RandBetween();
            var args = FunctionsHelper.CreateArgs(-5, 0);
            var result = func.Execute(args, _parsingContext);
            CollectionAssert.Contains(new List<double> { 0d, -1d, -2d, -3d, -4d, -5d }, result.Result);
        }

        [TestMethod]
        public void CountShouldReturnNumberOfNumericItems()
        {
            var func = new Count();
            var args = FunctionsHelper.CreateArgs(1d, 2m, 3, new DateTime(2012, 4, 1), "4");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void CountShouldIncludeNumericStringsAndDatesInArray()
        {
            var func = new Count();
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(1d, 2m, 3, new DateTime(2012, 4, 1), "4"));
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }


        [TestMethod]
        public void CountShouldIncludeEnumerableMembers()
        {
            var func = new Count();
            var args = FunctionsHelper.CreateArgs(1d, FunctionsHelper.CreateArgs(12, 13));
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod, Ignore]
        public void CountShouldIgnoreHiddenValuesIfIgnoreHiddenValuesIsTrue()
        {
            var func = new Count();
            func.IgnoreHiddenValues = true;
            var args = FunctionsHelper.CreateArgs(1d, FunctionsHelper.CreateArgs(12, 13));
            args.ElementAt(0).SetExcelStateFlag(ExcelCellState.HiddenCell);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        [TestMethod]
        public void CountAShouldCountEmptyString()
        {
            var func = new CountA();
            var args = FunctionsHelper.CreateArgs(1d, 2m, 3, new DateTime(2012, 4, 1), "4", null, string.Empty);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(6d, result.Result);
        }

        [TestMethod]
        public void CountAShouldIncludeEnumerableMembers()
        {
            var func = new CountA();
            var args = FunctionsHelper.CreateArgs(1d, FunctionsHelper.CreateArgs(12, 13));
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod]
        public void CountAShouldIgnoreHiddenValuesIfIgnoreHiddenValuesIsTrue()
        {
            var func = new CountA();
            func.IgnoreHiddenValues = true;
            var args = FunctionsHelper.CreateArgs(1d, FunctionsHelper.CreateArgs(12, 13));
            args.ElementAt(0).SetExcelStateFlag(ExcelCellState.HiddenCell);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        [TestMethod]
        public void ProductShouldMultiplyArguments()
        {
            var func = new Product();
            var args = FunctionsHelper.CreateArgs(2d, 2d, 4d);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(16d, result.Result);
        }

        [TestMethod]
        public void ProductShouldHandleEnumerable()
        {
            var func = new Product();
            var args = FunctionsHelper.CreateArgs(2d, 2d, FunctionsHelper.CreateArgs(4d, 2d));
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(32d, result.Result);
        }

        [TestMethod]
        public void ProductShouldIgnoreHiddenValuesIfIgnoreHiddenIsTrue()
        {
            var func = new Product();
            func.IgnoreHiddenValues = true;
            var args = FunctionsHelper.CreateArgs(2d, 2d, FunctionsHelper.CreateArgs(4d, 2d));
            args.ElementAt(1).SetExcelStateFlag(ExcelCellState.HiddenCell);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(16d, result.Result);
        }

        [TestMethod]
        public void ProductShouldHandleFirstItemIsEnumerable()
        {
            var func = new Product();
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(4d, 2d), 2d, 2d);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(32d, result.Result);
        }

        [TestMethod]
        public void VarShouldReturnCorrectResult()
        {
            var func = new Var();
            var args = FunctionsHelper.CreateArgs(1, 2, 3, 4);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1.6667d, System.Math.Round((double)result.Result, 4));
        }

        [TestMethod]
        public void VaraShouldReturnCorrectResult()
        {
            var func = new Vara();
            var args = FunctionsHelper.CreateArgs(1, 3, 5, 2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2.9167d, System.Math.Round((double)result.Result, 4));

            args = FunctionsHelper.CreateArgs(1, 3, 5, 2, true, "text");
            result = func.Execute(args, _parsingContext);
            Assert.AreEqual(3.2d, System.Math.Round((double)result.Result, 4));
        }

        [TestMethod]
        public void VarpaShouldReturnCorrectResult()
        {
            var func = new Varpa();
            var args = FunctionsHelper.CreateArgs(1, 3, 5, 2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2.1875, System.Math.Round((double)result.Result, 4));

            args = FunctionsHelper.CreateArgs(1, 3, 5, 2, true, "text");
            result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2.6667d, System.Math.Round((double)result.Result, 4));
        }

        [TestMethod]
        public void VarDotSShouldReturnCorrectResult()
        {
            var func = new VarDotS();
            var args = FunctionsHelper.CreateArgs(1, 2, 3, 4);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1.6667d, System.Math.Round((double)result.Result, 4));
        }

        [TestMethod]
        public void VarShouldIgnoreHiddenValuesIfIgnoreHiddenIsTrue()
        {
            var func = new Var();
            func.IgnoreHiddenValues = true;
            var args = FunctionsHelper.CreateArgs(1, 2, 3, 4, 9);
            args.Last().SetExcelStateFlag(ExcelCellState.HiddenCell);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1.6667d, System.Math.Round((double)result.Result, 4));
        }

        [TestMethod]
        public void VarPShouldReturnCorrectResult()
        {
            var func = new VarP();
            var args = FunctionsHelper.CreateArgs(1, 2, 3, 4);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1.25d, result.Result);
        }

        [TestMethod]
        public void VarDotPShouldReturnCorrectResult()
        {
            var func = new VarDotP();
            var args = FunctionsHelper.CreateArgs(1, 2, 3, 4);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1.25d, result.Result);
        }

        [TestMethod]
        public void VarPShouldIgnoreHiddenValuesIfIgnoreHiddenIsTrue()
        {
            var func = new VarP();
            func.IgnoreHiddenValues = true;
            var args = FunctionsHelper.CreateArgs(1, 2, 3, 4, 9);
            args.Last().SetExcelStateFlag(ExcelCellState.HiddenCell);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1.25d, result.Result);
        }

        [TestMethod]
        public void ModShouldReturnCorrectResult()
        {
            var func = new Mod();
            var args = FunctionsHelper.CreateArgs(5, 2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void CosShouldReturnCorrectResult()
        {
            var func = new Cos();
            var args = FunctionsHelper.CreateArgs(2);
            var result = func.Execute(args, _parsingContext);
            var roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(-0.416146837d, roundedResult);
        }

        [TestMethod]
        public void CosHShouldReturnCorrectResult()
        {
            var func = new Cosh();
            var args = FunctionsHelper.CreateArgs(2);
            var result = func.Execute(args, _parsingContext);
            var roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(3.762195691, roundedResult);
        }

        [TestMethod]
        public void AcosShouldReturnCorrectResult()
        {
            var func = new Acos();
            var args = FunctionsHelper.CreateArgs(0.1);
            var result = func.Execute(args, _parsingContext);
            var roundedResult = Math.Round((double)result.Result, 4);
            Assert.AreEqual(1.4706, roundedResult);
        }

        [TestMethod]
        public void ACosHShouldReturnCorrectResult()
        {
            var func = new Acosh();
            var args = FunctionsHelper.CreateArgs(2);
            var result = func.Execute(args, _parsingContext);
            var roundedResult = Math.Round((double)result.Result, 3);
            Assert.AreEqual(1.317, roundedResult);
        }

        [TestMethod]
        public void SinShouldReturnCorrectResult()
        {
            var func = new Sin();
            var args = FunctionsHelper.CreateArgs(2);
            var result = func.Execute(args, _parsingContext);
            var roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(0.909297427, roundedResult);
        }

        [TestMethod]
        public void SinhShouldReturnCorrectResult()
        {
            var func = new Sinh();
            var args = FunctionsHelper.CreateArgs(2);
            var result = func.Execute(args, _parsingContext);
            var roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(3.626860408d, roundedResult);
        }

        [TestMethod]
        public void TanShouldReturnCorrectResult()
        {
            var func = new Tan();
            var args = FunctionsHelper.CreateArgs(2);
            var result = func.Execute(args, _parsingContext);
            var roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(-2.185039863d, roundedResult);
        }

        [TestMethod]
        public void TanhShouldReturnCorrectResult()
        {
            var func = new Tanh();
            var args = FunctionsHelper.CreateArgs(2);
            var result = func.Execute(args, _parsingContext);
            var roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(0.96402758d, roundedResult);
        }

        [TestMethod]
        public void AtanShouldReturnCorrectResult()
        {
            var func = new Atan();
            var args = FunctionsHelper.CreateArgs(10);
            var result = func.Execute(args, _parsingContext);
            var roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(1.471127674d, roundedResult);
        }

        [TestMethod]
        public void Atan2ShouldReturnCorrectResult()
        {
            var func = new Atan2();
            var args = FunctionsHelper.CreateArgs(1, 2);
            var result = func.Execute(args, _parsingContext);
            var roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(1.107148718d, roundedResult);
        }

        [TestMethod]
        public void AtanhShouldReturnCorrectResult()
        {
            var func = new Atanh();
            var args = FunctionsHelper.CreateArgs(0.1);
            var result = func.Execute(args, _parsingContext);
            var roundedResult = Math.Round((double)result.Result, 4);
            Assert.AreEqual(0.1003d, roundedResult);
        }

        [TestMethod]
        public void LogShouldReturnCorrectResult()
        {
            var func = new Log();
            var args = FunctionsHelper.CreateArgs(2);
            var result = func.Execute(args, _parsingContext);
            var roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(0.301029996d, roundedResult);
        }

        [TestMethod]
        public void LogShouldReturnCorrectResultWithBase()
        {
            var func = new Log();
            var args = FunctionsHelper.CreateArgs(2, 2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void Log10ShouldReturnCorrectResult()
        {
            var func = new Log10();
            var args = FunctionsHelper.CreateArgs(2);
            var result = func.Execute(args, _parsingContext);
            var roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(0.301029996d, roundedResult);
        }

        [TestMethod]
        public void LnShouldReturnCorrectResult()
        {
            var func = new Ln();
            var args = FunctionsHelper.CreateArgs(5);
            var result = func.Execute(args, _parsingContext);
            var roundedResult = Math.Round((double)result.Result, 5);
            Assert.AreEqual(1.60944d, roundedResult);
        }

        [TestMethod]
        public void SqrtPiShouldReturnCorrectResult()
        {
            var func = new SqrtPi();
            var args = FunctionsHelper.CreateArgs(2);
            var result = func.Execute(args, _parsingContext);
            var roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(2.506628275d, roundedResult);
        }

        [TestMethod]
        public void SignShouldReturnMinus1IfArgIsNegative()
        {
            var func = new Sign();
            var args = FunctionsHelper.CreateArgs(-2);
            var result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(-1d, result);
        }

        [TestMethod]
        public void SignShouldReturn1IfArgIsPositive()
        {
            var func = new Sign();
            var args = FunctionsHelper.CreateArgs(2);
            var result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(1d, result);
        }

        [TestMethod]
        public void RounddownShouldReturnCorrectResultWithPositiveNumber()
        {
            var func = new Rounddown();
            var args = FunctionsHelper.CreateArgs(9.999, 2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(9.99, result.Result);
        }

        [TestMethod]
        public void RounddownShouldHandleNegativeNumber()
        {
            var func = new Rounddown();
            var args = FunctionsHelper.CreateArgs(-9.999, 2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(-9.99, result.Result);
        }

        [TestMethod]
        public void RounddownShouldHandleNegativeNumDigits()
        {
            var func = new Rounddown();
            var args = FunctionsHelper.CreateArgs(999.999, -2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(900d, result.Result);
        }

        [TestMethod]
        public void RounddownShouldReturn0IfNegativeNumDigitsIsTooLarge()
        {
            var func = new Rounddown();
            var args = FunctionsHelper.CreateArgs(999.999, -4);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(0d, result.Result);
        }

        [TestMethod]
        public void RounddownShouldHandleZeroNumDigits()
        {
            var func = new Rounddown();
            var args = FunctionsHelper.CreateArgs(999.999, 0);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(999d, result.Result);
        }

        [TestMethod]
        public void RoundupShouldReturnCorrectResultWithPositiveNumber()
        {
            var func = new Roundup();
            var args = FunctionsHelper.CreateArgs(9.9911, 3);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(9.992, result.Result);
        }

        [TestMethod]
        public void RoundupShouldHandleNegativeNumDigits()
        {
            var func = new Roundup();
            var args = FunctionsHelper.CreateArgs(99123, -2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(99200d, result.Result);
        }

        [TestMethod]
        public void RoundupShouldHandleZeroNumDigits()
        {
            var func = new Roundup();
            var args = FunctionsHelper.CreateArgs(999.999, 0);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1000d, result.Result);
        }

        [TestMethod]
        public void TruncShouldReturnCorrectResult()
        {
            var func = new Trunc();
            var args = FunctionsHelper.CreateArgs(99.99);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(99d, result.Result);
        }

        [TestMethod]
        public void FactShouldRoundDownAndReturnCorrectResult()
        {
            var func = new Fact();
            var args = FunctionsHelper.CreateArgs(5.99);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(120d, result.Result);
        }

        [TestMethod]
        public void FactShouldReturnErrorNegativeNumber()
        {
            var func = new Fact();
            var args = FunctionsHelper.CreateArgs(-1);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(DataType.ExcelError, result.DataType);
        }

        [TestMethod]
        public void FactDoubleShouldReturnCorrectResult5()
        {
            var func = new FactDouble();
            var arg = FunctionsHelper.CreateArgs(5);
            var result = func.Execute(arg, _parsingContext);
            Assert.AreEqual(15d, result.Result);
        }

        [TestMethod]
        public void FactDoubleShouldReturnCorrectResult8()
        {
            var func = new FactDouble();
            var arg = FunctionsHelper.CreateArgs(8);
            var result = func.Execute(arg, _parsingContext);
            Assert.AreEqual(384d, result.Result);
        }

        [TestMethod]
        public void QuotientShouldReturnCorrectResult()
        {
            var func = new Quotient();
            var args = FunctionsHelper.CreateArgs(5, 2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2, result.Result);
        }

        [TestMethod]
        public void QuotientShouldReturnErrorDenomIs0()
        {
            var func = new Quotient();
            var args = FunctionsHelper.CreateArgs(1, 0);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(DataType.ExcelError, result.DataType);
        }

        [TestMethod]
        public void LargeShouldReturnTheLargestNumberIf1()
        {
            var func = new Large();
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(1, 2, 3), 1);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod]
        public void LargeShouldReturnTheSecondLargestNumberIf2()
        {
            var func = new Large();
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(4, 1, 2, 3), 2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod]
        public void LargeShouldReturnErrorIfIndexOutOfBounds()
        {
            var func = new Large();
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(4, 1, 2, 3), 6);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(DataType.ExcelError, result.DataType);
        }

        [TestMethod]
        public void SmallShouldReturnTheSmallestNumberIf1()
        {
            var func = new Small();
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(1, 2, 3), 1);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void SmallShouldReturnTheSecondSmallestNumberIf2()
        {
            var func = new Small();
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(4, 1, 2, 3), 2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        [TestMethod]
        public void SmallShouldThrowIfIndexOutOfBounds()
        {
            var func = new Small();
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(4, 1, 2, 3), 6);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(DataType.ExcelError, result.DataType);
        }

        [TestMethod]
        public void MedianShouldReturnErrorIfNoArgs()
        {
            var func = new Median();
            var args = FunctionsHelper.Empty();
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(DataType.ExcelError, result.DataType);
        }

        [TestMethod]
        public void MedianShouldCalculateCorrectlyWithOneMember()
        {
            var func = new Median();
            var args = FunctionsHelper.CreateArgs(1);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void MedianShouldCalculateCorrectlyWithOddMembers()
        {
            var func = new Median();
            var args = FunctionsHelper.CreateArgs(3, 5, 1, 4, 2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod]
        public void MedianShouldCalculateCorrectlyWithEvenMembers()
        {
            var func = new Median();
            var args = FunctionsHelper.CreateArgs(1, 2, 3, 4);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2.5d, result.Result);
        }

        [TestMethod]
        public void CountIfShouldHandleNegativeCriteria()
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("test");
                sheet1.Cells["A1"].Value = -1;
                sheet1.Cells["A2"].Value = -2;
                sheet1.Cells["A3"].Formula = "CountIf(A1:A2,\"-1\")";
                sheet1.Calculate();
                Assert.AreEqual(1d, sheet1.Cells["A3"].Value);
            }
        }

        [TestMethod]
        public void OddShouldRound0To1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("test");
                sheet1.Cells["A1"].Value = 0;
                sheet1.Cells["A3"].Formula = "ODD(A1)";
                sheet1.Calculate();
                Assert.AreEqual(1, sheet1.Cells["A3"].Value);
            }
        }

        [TestMethod]
        public void OddShouldRound1To1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("test");
                sheet1.Cells["A1"].Value = 1;
                sheet1.Cells["A3"].Formula = "ODD(A1)";
                sheet1.Calculate();
                Assert.AreEqual(1, sheet1.Cells["A3"].Value);
            }
        }

        [TestMethod]
        public void OddShouldRound2To3()
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("test");
                sheet1.Cells["A1"].Value = 2;
                sheet1.Cells["A3"].Formula = "ODD(A1)";
                sheet1.Calculate();
                Assert.AreEqual(3, sheet1.Cells["A3"].Value);
            }
        }

        [TestMethod]
        public void OddShouldRoundMinus1point3ToMinus3()
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("test");
                sheet1.Cells["A1"].Value = -1.3;
                sheet1.Cells["A3"].Formula = "ODD(A1)";
                sheet1.Calculate();
                Assert.AreEqual(-3, sheet1.Cells["A3"].Value);
            }
        }

        [TestMethod]
        public void EvenShouldRound0To0()
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("test");
                sheet1.Cells["A1"].Value = 0;
                sheet1.Cells["A3"].Formula = "EVEN(A1)";
                sheet1.Calculate();
                Assert.AreEqual(0, sheet1.Cells["A3"].Value);
            }
        }

        [TestMethod]
        public void EvenShouldRound1To2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("test");
                sheet1.Cells["A1"].Value = 1;
                sheet1.Cells["A3"].Formula = "EVEN(A1)";
                sheet1.Calculate();
                Assert.AreEqual(2, sheet1.Cells["A3"].Value);
            }
        }

        [TestMethod]
        public void EvenShouldRound2To2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("test");
                sheet1.Cells["A1"].Value = 2;
                sheet1.Cells["A3"].Formula = "EVEN(A1)";
                sheet1.Calculate();
                Assert.AreEqual(2, sheet1.Cells["A3"].Value);
            }
        }

        [TestMethod]
        public void EvenShouldRoundMinus1point3ToMinus2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("test");
                sheet1.Cells["A1"].Value = -1.3;
                sheet1.Cells["A3"].Formula = "EVEN(A1)";
                sheet1.Calculate();
                Assert.AreEqual(-2, sheet1.Cells["A3"].Value);
            }
        }

        [TestMethod]
        public void Rank()
        {
            using (var p = new ExcelPackage())
            {
                var w = p.Workbook.Worksheets.Add("testsheet");
                w.SetValue(1, 1, 1);
                w.SetValue(2, 1, 1);
                w.SetValue(3, 1, 2);
                w.SetValue(4, 1, 2);
                w.SetValue(5, 1, 4);
                w.SetValue(6, 1, 4);

                w.SetFormula(1, 2, "RANK(1,A1:A5)");
                w.SetFormula(1, 3, "RANK(1,A1:A5,1)");
                w.SetFormula(1, 4, "RANK.AVG(1,A1:A5)");
                w.SetFormula(1, 5, "RANK.AVG(1,A1:A5,1)");

                w.SetFormula(2, 2, "RANK.EQ(2,A1:A5)");
                w.SetFormula(2, 3, "RANK.EQ(2,A1:A5,1)");
                w.SetFormula(2, 4, "RANK.AVG(2,A1:A5,1)");
                w.SetFormula(2, 5, "RANK.AVG(2,A1:A5,0)");

                w.SetFormula(3, 2, "RANK(3,A1:A5)");
                w.SetFormula(3, 3, "RANK(3,A1:A5,1)");
                w.SetFormula(3, 4, "RANK.AVG(3,A1:A5,1)");
                w.SetFormula(3, 5, "RANK.AVG(3,A1:A5,0)");

                w.SetFormula(4, 2, "RANK.EQ(4,A1:A5)");
                w.SetFormula(4, 3, "RANK.EQ(4,A1:A5,1)");
                w.SetFormula(4, 4, "RANK.AVG(4,A1:A5,1)");
                w.SetFormula(4, 5, "RANK.AVG(4,A1:A5)");


                w.SetFormula(5, 4, "RANK.AVG(4,A1:A6,1)");
                w.SetFormula(5, 5, "RANK.AVG(4,A1:A6)");

                w.Calculate();

                Assert.AreEqual(w.GetValue(1, 2), 4D);
                Assert.AreEqual(w.GetValue(1, 3), 1D);
                Assert.AreEqual(w.GetValue(1, 4), 4.5D);
                Assert.AreEqual(w.GetValue(1, 5), 1.5D);

                Assert.AreEqual(w.GetValue(2, 2), 2D);
                Assert.AreEqual(w.GetValue(2, 3), 3D);
                Assert.AreEqual(w.GetValue(2, 4), 3.5D);
                Assert.AreEqual(w.GetValue(2, 5), 2.5D);

                Assert.IsInstanceOfType(w.GetValue(3, 2), typeof(ExcelErrorValue));
                Assert.IsInstanceOfType(w.GetValue(3, 3), typeof(ExcelErrorValue));
                Assert.IsInstanceOfType(w.GetValue(3, 4), typeof(ExcelErrorValue));
                Assert.IsInstanceOfType(w.GetValue(3, 5), typeof(ExcelErrorValue));

                Assert.AreEqual(w.GetValue(4, 2), 1D);
                Assert.AreEqual(w.GetValue(4, 3), 5D);
                Assert.AreEqual(w.GetValue(4, 4), 5D);
                Assert.AreEqual(w.GetValue(4, 5), 1D);

                Assert.AreEqual(w.GetValue(5, 4), 5.5D);
                Assert.AreEqual(w.GetValue(5, 5), 1.5D);
            }
        }

        [TestMethod]
        public void PercentrankInc_Test1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["A5"].Formula = "PERCENTRANK.INC(A1:A3,2)";
                sheet.Calculate();
                var result = sheet.Cells["A5"].Value;
                Assert.AreEqual(0.5, result);

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 6;
                sheet.Cells["A5"].Formula = "PERCENTRANK.INC(A1:A3,3)";
                sheet.Calculate();
                result = sheet.Cells["A5"].Value;
                Assert.AreEqual(0.625, result);
            }
        }

        [TestMethod]
        public void PercentrankInc_Test2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 4;
                sheet.Cells["A4"].Value = 6.5;
                sheet.Cells["A5"].Value = 8;
                sheet.Cells["A6"].Value = 9;
                sheet.Cells["A7"].Value = 10;
                sheet.Cells["A8"].Value = 12;
                sheet.Cells["A9"].Value = 14;

                sheet.Cells["A10"].Formula = "PERCENTRANK.INC(A1:A9,6.5)";
                sheet.Calculate();
                var result = sheet.Cells["A10"].Value;
                Assert.AreEqual(0.375, result);

                sheet.Cells["A10"].Formula = "PERCENTRANK.INC(A1:A9,7,5)";
                sheet.Calculate();
                result = sheet.Cells["A10"].Value;
                Assert.AreEqual(0.41666, result);
            }
        }

        [TestMethod] 
        public void PercentrankExc_Test1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 4;
                sheet.Cells["A4"].Value = 6.5;
                sheet.Cells["A5"].Value = 8;
                sheet.Cells["A6"].Value = 9;
                sheet.Cells["A7"].Value = 10;
                sheet.Cells["A8"].Value = 12;
                sheet.Cells["A9"].Value = 14;
                sheet.Cells["B1"].Formula = "PERCENTRANK.EXC(A1:A9,6.5)";
                sheet.Calculate();
                var result = sheet.Cells["B1"].Value;
                Assert.AreEqual(0.4, result);

                sheet.Cells["B1"].Formula = "PERCENTRANK.EXC(A1:A9,7,5)";
                sheet.Calculate();
                result = sheet.Cells["B1"].Value;
                Assert.AreEqual(0.43333, result);

                sheet.Cells["B1"].Formula = "PERCENTRANK.EXC(A1:A9,18)";
                sheet.Calculate();
                result = sheet.Cells["B1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), result);
            }
        }

        [TestMethod]
        public void Percentile_Test1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["A1"].Value = 2;
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["A3"].Value = 6;
                sheet.Cells["A4"].Value = 4;
                sheet.Cells["A5"].Value = 3;
                sheet.Cells["A6"].Value = 5;

                sheet.Cells["A10"].Formula = "PERCENTILE(A1:A6,0.2)";
                sheet.Calculate();
                var result = sheet.Cells["A10"].Value;
                Assert.AreEqual(2d, result);

                sheet.Cells["A10"].Formula = "PERCENTILE(A1:A6,60%)";
                sheet.Calculate();
                result = sheet.Cells["A10"].Value;
                Assert.AreEqual(4d, result);

                sheet.Cells["A10"].Formula = "PERCENTILE(A1:A6,50%)";
                sheet.Calculate();
                result = sheet.Cells["A10"].Value;
                Assert.AreEqual(3.5d, result);

                sheet.Cells["A10"].Formula = "PERCENTILE(A1:A6,95%)";
                sheet.Calculate();
                result = sheet.Cells["A10"].Value;
                Assert.AreEqual(5.75d, result);
            }
        }

        [TestMethod]
        public void PercentileInc_Test1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["A1"].Value = 0;
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["A3"].Value = 2;
                sheet.Cells["A4"].Value = 3;
                sheet.Cells["A5"].Value = 4;
                sheet.Cells["A6"].Value = 5;

                sheet.Cells["A10"].Formula = "PERCENTILE.INC(A1:A6,0.2)";
                sheet.Calculate();
                var result = sheet.Cells["A10"].Value;
                Assert.AreEqual(1d, result);

                sheet.Cells["A10"].Formula = "PERCENTILE.INC(A1:A6,60%)";
                sheet.Calculate();
                result = sheet.Cells["A10"].Value;
                Assert.AreEqual(3d, result);

                sheet.Cells["A10"].Formula = "PERCENTILE.INC(A1:A6,50%)";
                sheet.Calculate();
                result = sheet.Cells["A10"].Value;
                Assert.AreEqual(2.5d, result);

                sheet.Cells["A10"].Formula = "PERCENTILE.INC(A1:A6,95%)";
                sheet.Calculate();
                result = sheet.Cells["A10"].Value;
                Assert.AreEqual(4.75d, result);
            }
        }

        [TestMethod]
        public void PercentileExc_Test1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["A4"].Value = 4;

                sheet.Cells["A10"].Formula = "PERCENTILE.EXC(A1:A4,0.2)";
                sheet.Calculate();
                var result = sheet.Cells["A10"].Value;
                Assert.AreEqual(1d, result);

                sheet.Cells["A10"].Formula = "PERCENTILE.EXC(A1:A4,60%)";
                sheet.Calculate();
                result = sheet.Cells["A10"].Value;
                Assert.AreEqual(3d, result);

                sheet.Cells["A10"].Formula = "PERCENTILE.EXC(A1:A4,50%)";
                sheet.Calculate();
                result = sheet.Cells["A10"].Value;
                Assert.AreEqual(2.5d, result);

                sheet.Cells["A10"].Formula = "PERCENTILE.EXC(A1:A4,95%)";
                sheet.Calculate();
                result = sheet.Cells["A10"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void Quartile_Test1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["A1"].Value = 2;
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["A3"].Value = 6;
                sheet.Cells["A4"].Value = 4;
                sheet.Cells["A5"].Value = 3;
                sheet.Cells["A6"].Value = 5;
                sheet.Cells["A7"].Value = 0;

                sheet.Cells["A10"].Formula = "QUARTILE(A1:A7,0)";
                sheet.Calculate();
                var result = sheet.Cells["A10"].Value;
                Assert.AreEqual(0d, result);

                sheet.Cells["A10"].Formula = "QUARTILE(A1:A7,1)";
                sheet.Calculate();
                result = sheet.Cells["A10"].Value;
                Assert.AreEqual(1.5d, result);

                sheet.Cells["A10"].Formula = "QUARTILE(A1:A7, 2)";
                sheet.Calculate();
                result = sheet.Cells["A10"].Value;
                Assert.AreEqual(3d, result);

                sheet.Cells["A10"].Formula = "QUARTILE(A1:A7,3)";
                sheet.Calculate();
                result = sheet.Cells["A10"].Value;
                Assert.AreEqual(4.5d, result);

                sheet.Cells["A10"].Formula = "QUARTILE(A1:A7,4)";
                sheet.Calculate();
                result = sheet.Cells["A10"].Value;
                Assert.AreEqual(6d, result);
            }
        }

        [TestMethod]
        public void QuartileInc_Test1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["A1"].Value = 2;
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["A3"].Value = 6;
                sheet.Cells["A4"].Value = 4;
                sheet.Cells["A5"].Value = 3;
                sheet.Cells["A6"].Value = 5;
                sheet.Cells["A7"].Value = 0;

                sheet.Cells["A10"].Formula = "QUARTILE.INC(A1:A7,0)";
                sheet.Calculate();
                var result = sheet.Cells["A10"].Value;
                Assert.AreEqual(0d, result);

                sheet.Cells["A10"].Formula = "QUARTILE.INC(A1:A7,1)";
                sheet.Calculate();
                result = sheet.Cells["A10"].Value;
                Assert.AreEqual(1.5d, result);

                sheet.Cells["A10"].Formula = "QUARTILE.INC(A1:A7, 2)";
                sheet.Calculate();
                result = sheet.Cells["A10"].Value;
                Assert.AreEqual(3d, result);

                sheet.Cells["A10"].Formula = "QUARTILE.INC(A1:A7,3)";
                sheet.Calculate();
                result = sheet.Cells["A10"].Value;
                Assert.AreEqual(4.5d, result);

                sheet.Cells["A10"].Formula = "QUARTILE.INC(A1:A7,4)";
                sheet.Calculate();
                result = sheet.Cells["A10"].Value;
                Assert.AreEqual(6d, result);
            }
        }

        [TestMethod]
        public void ModeShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["A3"].Value = 2;
                sheet.Cells["A4"].Value = 2;
                sheet.Cells["A5"].Value = 2;
                sheet.Cells["A6"].Value = 3;
                sheet.Cells["B1"].Formula = "MODE(A1:A6)";
                sheet.Calculate();
                Assert.AreEqual(2d, sheet.Cells["B1"].Value);
            }
        }

        [TestMethod]
        public void ModeShouldReturnLowestIfMultipleResults()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["A1"].Value = 2;
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["A3"].Value = 2;
                sheet.Cells["A4"].Value = 1;
                sheet.Cells["A5"].Value = 3;
                sheet.Cells["A6"].Value = 3;
                sheet.Cells["B1"].Formula = "MODE.SNGL(A1:A6)";
                sheet.Calculate();
                Assert.AreEqual(1d, sheet.Cells["B1"].Value);
            }
        }

        [TestMethod]
        public void MultinomialShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["A1"].Value = 3;
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["A3"].Value = 2;
                sheet.Cells["A4"].Value = 5;
                sheet.Cells["B1"].Formula = "MULTINOMIAL(A1:A4)";
                sheet.Calculate();
                Assert.AreEqual(27720d, sheet.Cells["B1"].Value);
            }
        }

        [TestMethod]
        public void CovarShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["A1"].Value = 3;
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["A3"].Value = 2;
                sheet.Cells["A4"].Value = 5;
                sheet.Cells["B1"].Value = 2;
                sheet.Cells["B2"].Value = 6;
                sheet.Cells["B3"].Value = 2;
                sheet.Cells["B4"].Value = 8;

                sheet.Cells["C1"].Formula = "COVAR(A1:A4, B1:B4)";
                sheet.Calculate();
                Assert.AreEqual(1.625d, sheet.Cells["C1"].Value);

                sheet.Cells["C1"].Formula = "COVARIANCE.P(A1:A4, B1:B4)";
                sheet.Calculate();
                Assert.AreEqual(1.625d, sheet.Cells["C1"].Value);
            }
        }

        [TestMethod]
        public void CovarianceSshouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["A1"].Value = 3;
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["A3"].Value = 2;
                sheet.Cells["A4"].Value = 5;
                sheet.Cells["B1"].Value = 2;
                sheet.Cells["B2"].Value = 6;
                sheet.Cells["B3"].Value = 2;
                sheet.Cells["B4"].Value = 8;

                sheet.Cells["C1"].Formula = "COVARIANCE.S(A1:A4, B1:B4)";
                sheet.Calculate();
                Assert.AreEqual(2.16667d, System.Math.Round((double)sheet.Cells["C1"].Value, 5));
            }
        }
    }
}

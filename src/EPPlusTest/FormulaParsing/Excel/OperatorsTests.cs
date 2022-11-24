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
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest.Excel
{
    [TestClass]
    public class OperatorsTests
    {
        private ParsingContext _ctx = ParsingContext.Create();

        [TestMethod]
        public void OperatorPlusShouldThrowExceptionIfNonNumericOperand()
        {
            var result = Operator.Plus.Apply(new CompileResult(1, DataType.Integer), new CompileResult("a", DataType.String), _ctx);
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), result.Result);
        }

        [TestMethod]
        public void OperatorPlusShouldAddNumericStringAndNumber()
        {
            var result = Operator.Plus.Apply(new CompileResult(1, DataType.Integer), new CompileResult("2", DataType.String), _ctx);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod]
        public void OperatorMinusShouldThrowExceptionIfNonNumericOperand()
        {
            var result = Operator.Minus.Apply(new CompileResult(1, DataType.Integer), new CompileResult("a", DataType.String), _ctx);
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), result.Result);
        }

        [TestMethod]
        public void OperatorMinusShouldSubtractNumericStringAndNumber()
        {
            var result = Operator.Minus.Apply(new CompileResult(5, DataType.Integer), new CompileResult("2", DataType.String), _ctx);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod]
        public void OperatorDivideShouldReturnDivideByZeroIfRightOperandIsZero()
        {
            var result = Operator.Divide.Apply(new CompileResult(1d, DataType.Decimal), new CompileResult(0d, DataType.Decimal), _ctx);
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Div0), result.Result);
        }

        [TestMethod]
        public void OperatorDivideShouldDivideCorrectly()
        {
            var result = Operator.Divide.Apply(new CompileResult(9d, DataType.Decimal), new CompileResult(3d, DataType.Decimal), _ctx);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod]
        public void OperatorDivideShouldReturnValueErrorIfNonNumericOperand()
        {
            var result = Operator.Divide.Apply(new CompileResult(1, DataType.Integer), new CompileResult("a", DataType.String), _ctx);
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), result.Result);
        }

        [TestMethod]
        public void OperatorDivideShouldDivideNumericStringAndNumber()
        {
            var result = Operator.Divide.Apply(new CompileResult(9, DataType.Integer), new CompileResult("3", DataType.String), _ctx);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod]
        public void OperatorMultiplyShouldThrowExceptionIfNonNumericOperand()
        {
            Operator.Multiply.Apply(new CompileResult(1, DataType.Integer), new CompileResult("a", DataType.String), _ctx);
        }

        [TestMethod]
        public void OperatoMultiplyShouldMultiplyNumericStringAndNumber()
        {
            var result = Operator.Multiply.Apply(new CompileResult(1, DataType.Integer), new CompileResult("3", DataType.String), _ctx);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod]
        public void OperatorConcatShouldConcatTwoStrings()
        {
            var result = Operator.Concat.Apply(new CompileResult("a", DataType.String), new CompileResult("b", DataType.String), _ctx);
            Assert.AreEqual("ab", result.Result);
        }

        [TestMethod]
        public void OperatorConcatShouldConcatANumberAndAString()
        {
            var result = Operator.Concat.Apply(new CompileResult(12, DataType.Integer), new CompileResult("b", DataType.String), _ctx);
            Assert.AreEqual("12b", result.Result);
        }

        [TestMethod]
        public void OperatorEqShouldReturnTruefSuppliedValuesAreEqual()
        {
            var result = Operator.Eq.Apply(new CompileResult(12, DataType.Integer), new CompileResult(12, DataType.Integer), _ctx);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void OperatorEqShouldReturnFalsefSuppliedValuesDiffer()
        {
            var result = Operator.Eq.Apply(new CompileResult(11, DataType.Integer), new CompileResult(12, DataType.Integer), _ctx);
            Assert.IsFalse((bool)result.Result);
        }

        [TestMethod]
        public void OperatorNotEqualToShouldReturnTruefSuppliedValuesDiffer()
        {
            var result = Operator.NotEqualsTo.Apply(new CompileResult(11, DataType.Integer), new CompileResult(12, DataType.Integer), _ctx);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void OperatorNotEqualToShouldReturnFalsefSuppliedValuesAreEqual()
        {
            var result = Operator.NotEqualsTo.Apply(new CompileResult(11, DataType.Integer), new CompileResult(11, DataType.Integer), _ctx);
            Assert.IsFalse((bool)result.Result);
        }

        [TestMethod]
        public void OperatorGreaterThanToShouldReturnTrueIfLeftIsSetAndRightIsNull()
        {
            var result = Operator.GreaterThan.Apply(new CompileResult(11, DataType.Integer), new CompileResult(null, DataType.Empty), _ctx);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void OperatorGreaterThanToShouldReturnTrueIfLeftIs11AndRightIs10()
        {
            var result = Operator.GreaterThan.Apply(new CompileResult(11, DataType.Integer), new CompileResult(10, DataType.Integer), _ctx);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void OperatorExpShouldReturnCorrectResult()
        {
            var result = Operator.Exp.Apply(new CompileResult(2, DataType.Integer), new CompileResult(3, DataType.Integer), _ctx);
            Assert.AreEqual(8d, result.Result);
        }

        [TestMethod]
		public void OperatorsActingOnNumericStrings()
		{
            var ctx = ParsingContext.Create();
			double number1 = 42.0;
			double number2 = -143.75;
			CompileResult result1 = new CompileResult(number1.ToString("n"), DataType.String);
			CompileResult result2 = new CompileResult(number2.ToString("n"), DataType.String);
			var operatorResult = Operator.Concat.Apply(result1, result2, ctx);
			Assert.AreEqual($"{number1.ToString("n")}{number2.ToString("n")}", operatorResult.Result);
			operatorResult = Operator.Divide.Apply(result1, result2, ctx);
			Assert.AreEqual(number1 / number2, operatorResult.Result);
			operatorResult = Operator.Exp.Apply(result1, result2, ctx);
			Assert.AreEqual(Math.Pow(number1, number2), operatorResult.Result);
			operatorResult = Operator.Minus.Apply(result1, result2, ctx);
			Assert.AreEqual(number1 - number2, operatorResult.Result);
			operatorResult = Operator.Multiply.Apply(result1, result2, ctx);
			Assert.AreEqual(number1 * number2, operatorResult.Result);
			operatorResult = Operator.Percent.Apply(result1, result2, ctx);
			Assert.AreEqual(number1 * number2, operatorResult.Result);
			operatorResult = Operator.Plus.Apply(result1, result2, ctx);
			Assert.AreEqual(number1 + number2, operatorResult.Result);
			// Comparison operators always compare string-wise and don't parse out the actual numbers.
			operatorResult = Operator.NotEqualsTo.Apply(result1, new CompileResult(number1.ToString("n0"), DataType.String), ctx);
			Assert.IsTrue((bool)operatorResult.Result);
			operatorResult = Operator.Eq.Apply(result1, new CompileResult(number1.ToString("n0"), DataType.String), ctx);
			Assert.IsFalse((bool)operatorResult.Result);
			operatorResult = Operator.GreaterThan.Apply(result1, result2, ctx);
			Assert.IsTrue((bool)operatorResult.Result);
			operatorResult = Operator.GreaterThanOrEqual.Apply(result1, result2, ctx);
			Assert.IsTrue((bool)operatorResult.Result);
			operatorResult = Operator.LessThan.Apply(result1, result2, ctx);
			Assert.IsFalse((bool)operatorResult.Result);
			operatorResult = Operator.LessThanOrEqual.Apply(result1, result2, ctx);
			Assert.IsFalse((bool)operatorResult.Result);
		}

		[TestMethod]
		public void OperatorsActingOnDateStrings()
		{
            var ctx = ParsingContext.Create();
            const string dateFormat = "M-dd-yyyy";
            DateTime date1 = new DateTime(2015, 2, 20);
            DateTime date2 = new DateTime(2015, 12, 1);
            var numericDate1 = date1.ToOADate();
            var numericDate2 = date2.ToOADate();
            CompileResult result1 = new CompileResult(date1.ToString(dateFormat), DataType.String); // 2/20/2015
            CompileResult result2 = new CompileResult(date2.ToString(dateFormat), DataType.String); // 12/1/2015
            var operatorResult = Operator.Concat.Apply(result1, result2, ctx);
            Assert.AreEqual($"{date1.ToString(dateFormat)}{date2.ToString(dateFormat)}", operatorResult.Result);
            operatorResult = Operator.Divide.Apply(result1, result2, ctx);
            Assert.AreEqual(numericDate1 / numericDate2, operatorResult.Result);
            operatorResult = Operator.Exp.Apply(result1, result2, ctx);
            Assert.AreEqual(Math.Pow(numericDate1, numericDate2), operatorResult.Result);
            operatorResult = Operator.Minus.Apply(result1, result2, ctx);
			Assert.AreEqual(numericDate1 - numericDate2, operatorResult.Result);
			operatorResult = Operator.Multiply.Apply(result1, result2, ctx);
			Assert.AreEqual(numericDate1 * numericDate2, operatorResult.Result);
			operatorResult = Operator.Percent.Apply(result1, result2, ctx);
			Assert.AreEqual(numericDate1 * numericDate2, operatorResult.Result);
			operatorResult = Operator.Plus.Apply(result1, result2, ctx);
			Assert.AreEqual(numericDate1 + numericDate2, operatorResult.Result);
			// Comparison operators always compare string-wise and don't parse out the actual numbers.
			operatorResult = Operator.Eq.Apply(result1, new CompileResult(date1.ToString("f"), DataType.String), ctx);
			Assert.IsFalse((bool)operatorResult.Result);
			operatorResult = Operator.NotEqualsTo.Apply(result1, new CompileResult(date1.ToString("f"), DataType.String), ctx);
			Assert.IsTrue((bool)operatorResult.Result);
			operatorResult = Operator.GreaterThan.Apply(result1, result2, ctx);
			Assert.IsTrue((bool)operatorResult.Result);
			operatorResult = Operator.GreaterThanOrEqual.Apply(result1, result2, ctx);
			Assert.IsTrue((bool)operatorResult.Result);
			operatorResult = Operator.LessThan.Apply(result1, result2, ctx);
			Assert.IsFalse((bool)operatorResult.Result);
			operatorResult = Operator.LessThanOrEqual.Apply(result1, result2, ctx);
			Assert.IsFalse((bool)operatorResult.Result);
		}
	}
}

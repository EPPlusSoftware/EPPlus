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
//using System.Diagnostics.Design;
using System.Globalization;
using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest
{
    [TestClass]
    public class ExpressionEvaluatorTests
    {
        private ExpressionEvaluator _evaluator;
        private ParsingContext _context;

        [TestInitialize]
        public void Setup()
        {
            _context = ParsingContext.Create();
            _evaluator = new ExpressionEvaluator(_context);
        }

        #region Numeric Expression Tests
        [TestMethod]
        public void EvaluateShouldReturnTrueIfOperandsAreEqual()
        {
            var result = _evaluator.Evaluate("1", "1");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateShouldReturnTrueIfOperandsAreMatchingButDifferentTypes()
        {
            var result = _evaluator.Evaluate(1d, "1");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateShouldEvaluateOperator()
        {
            var result = _evaluator.Evaluate(1d, "<2");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateShouldEvaluateNumericString()
        {
            var result = _evaluator.Evaluate("1", ">0");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateShouldHandleBooleanArg()
        {
            var result = _evaluator.Evaluate(true, "TRUE");
            Assert.IsTrue(result);
        }

        [TestMethod, ExpectedException(typeof(ArgumentException))]
        public void EvaluateShouldThrowIfOperatorIsNotBoolean()
        {
            var result = _evaluator.Evaluate(1d, "+1");
        }
        [TestMethod]
        public void EvaluateShouldEvaluateToGreaterThanMinusOne ()
        {
            var result = _evaluator.Evaluate(1d, "<>-1");
            Assert.IsTrue(result);
        }
        #endregion

        #region Date tests
        [TestMethod]
        public void EvaluateShouldHandleDateArg()
        {
#if (!Core)
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
#endif
            var result = _evaluator.Evaluate(new DateTime(2016,6,28), "2016-06-28");
            Assert.IsTrue(result);
#if (!Core)
            Thread.CurrentThread.CurrentCulture = ci;
#endif

        }

        [TestMethod]
        public void EvaluateShouldHandleDateArgWithOperator()
        {
#if (!Core)
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
#endif
            var result = _evaluator.Evaluate(new DateTime(2016, 6, 28), ">2016-06-27");
            Assert.IsTrue(result);
#if (!Core)
            Thread.CurrentThread.CurrentCulture = ci;
#endif
        }
        #endregion

        #region Blank Expression Tests
        [TestMethod]
        public void EvaluateBlankExpressionEqualsNull()
        {
            var result = _evaluator.Evaluate(null, "");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateBlankExpressionEqualsEmptyString()
        {
            var result = _evaluator.Evaluate(string.Empty, "");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateBlankExpressionEqualsZero()
        {
            var result = _evaluator.Evaluate(0d, "");
            Assert.IsFalse(result);
        }
#endregion

#region Quotes Expression Tests
        [TestMethod]
        public void EvaluateQuotesExpressionEqualsNull()
        {
            var result = _evaluator.Evaluate(null, "\"\"");
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void EvaluateQuotesExpressionEqualsZero()
        {
            var result = _evaluator.Evaluate(0d, "\"\"");
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void EvaluateQuotesExpressionEqualsCharacter()
        {
            var result = _evaluator.Evaluate("a", "\"\"");
            Assert.IsFalse(result);
        }
#endregion

#region NotEqualToZero Expression Tests
        [TestMethod]
        public void EvaluateNotEqualToZeroExpressionEqualsNull()
        {
            var result = _evaluator.Evaluate(null, "<>0");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateNotEqualToZeroExpressionEqualsEmptyString()
        {
            var result = _evaluator.Evaluate(string.Empty, "<>0");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateNotEqualToZeroExpressionEqualsCharacter()
        {
            var result = _evaluator.Evaluate("a", "<>0");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateNotEqualToZeroExpressionEqualsNonZero()
        {
            var result = _evaluator.Evaluate(1d, "<>0");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateNotEqualToZeroExpressionEqualsZero()
        {
            var result = _evaluator.Evaluate(0d, "<>0");
            Assert.IsFalse(result);
        }
#endregion

#region NotEqualToBlank Expression Tests
        [TestMethod]
        public void EvaluateNotEqualToBlankExpressionEqualsNull()
        {
            var result = _evaluator.Evaluate(null, "<>");
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void EvaluateNotEqualToBlankExpressionEqualsEmptyString()
        {
            var result = _evaluator.Evaluate(string.Empty, "<>");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateNotEqualToBlankExpressionEqualsCharacter()
        {
            var result = _evaluator.Evaluate("a", "<>");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateNotEqualToBlankExpressionEqualsNonZero()
        {
            var result = _evaluator.Evaluate(1d, "<>");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateNotEqualToBlankExpressionEqualsZero()
        {
            var result = _evaluator.Evaluate(0d, "<>");
            Assert.IsTrue(result);
        }
#endregion

#region Character Expression Tests
        [TestMethod]
        public void EvaluateCharacterExpressionEqualNull()
        {
            var result = _evaluator.Evaluate(null, "a");
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void EvaluateCharacterExpressionEqualsEmptyString()
        {
            var result = _evaluator.Evaluate(string.Empty, "a");
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void EvaluateCharacterExpressionEqualsNumeral()
        {
            var result = _evaluator.Evaluate(1d, "a");
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void EvaluateCharacterExpressionEqualsSameCharacter()
        {
            var result = _evaluator.Evaluate("a", "a");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateCharacterExpressionEqualsDifferentCharacter()
        {
            var result = _evaluator.Evaluate("b", "a");
            Assert.IsFalse(result);
        }
#endregion

#region CharacterWithOperator Expression Tests
        [TestMethod]
        public void EvaluateCharacterWithOperatorExpressionEqualNull()
        {
            var result = _evaluator.Evaluate(null, ">a");
            Assert.IsFalse(result);
            result = _evaluator.Evaluate(null, "<a");
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void EvaluateCharacterWithOperatorExpressionEqualsEmptyString()
        {
            var result = _evaluator.Evaluate(string.Empty, ">a");
            Assert.IsFalse(result);
            result = _evaluator.Evaluate(string.Empty, "<a");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateCharacterWithOperatorExpressionEqualsNumeral()
        {
            var result = _evaluator.Evaluate(1d, ">a");
            Assert.IsFalse(result);
            result = _evaluator.Evaluate(1d, "<a");
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void EvaluateShouldHandleLeadingEqualOperatorAndWildCard()
        {
            var result = _evaluator.Evaluate("TEST", "=*EST*");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateCharacterWithOperatorExpressionEqualsSameCharacter()
        {
            var result = _evaluator.Evaluate("a", ">a");
            Assert.IsFalse(result);
            result = _evaluator.Evaluate("a", ">=a");
            Assert.IsTrue(result);
            result = _evaluator.Evaluate("a", "<a");
            Assert.IsFalse(result);
            result = _evaluator.Evaluate("a", ">=a");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateCharacterWithOperatorExpressionEqualsDifferentCharacter()
        {
            var result = _evaluator.Evaluate("b", ">a");
            Assert.IsTrue(result);
            result = _evaluator.Evaluate("b", "<a");
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void EvaluateCharacterWithSpaceBetweenOperatorAndCharacter()
        {
            var result = _evaluator.Evaluate("b", "> a");
            Assert.IsTrue(result);
            result = _evaluator.Evaluate("b", "< a");
            Assert.IsFalse(result);
        }
#endregion
    }
}

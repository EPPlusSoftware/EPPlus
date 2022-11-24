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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
    [TestClass]
    public class IntegerExpressionTests
    {
        private ParsingContext _context = ParsingContext.Create();

        [TestMethod]
        public void MergeWithNextWithPlusOperatorShouldCalulateSumCorrectly()
        {
            var exp1 = new IntegerExpression("1", _context);
            exp1.Operator = Operator.Plus;
            var exp2 = new IntegerExpression("2", _context);

            var result = exp1.MergeWithNext(new List<Expression>() { exp1, exp2}, 0);

            Assert.AreEqual(3d, result.Compile().ResultValue);
        }

        [TestMethod]
        public void MergeWithNextWithPlusOperatorShouldSetNextPointer()
        {
            var exp1 = new IntegerExpression("1", _context);
            exp1.Operator = Operator.Plus;
            var exp2 = new IntegerExpression("2", _context);

            var l = new List<Expression>() { exp1, exp2 };
            var result = exp1.MergeWithNext(l, 0);

            Assert.AreEqual(1, l.Count);
            Assert.AreEqual(0, l.IndexOf(result));
        }

        //[TestMethod]
        //public void CompileShouldHandlePercent()
        //{
        //    var exp1 = new IntegerExpression("1");
        //    exp1.Operator = Operator.Percent;
        //    exp1.Next = ConstantExpressions.Percent;
        //    var result = exp1.Compile();
        //    Assert.AreEqual(0.01, result.Result);
        //    Assert.AreEqual(DataType.Decimal, result.DataType);
        //}
    }
}

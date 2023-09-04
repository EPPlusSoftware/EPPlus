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
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using EPPlusTest.FormulaParsing.TestHelpers;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Ranges;

namespace EPPlusTest.Excel.Functions
{
    [TestClass]
    public class ExcelFunctionTests
    {
        private class ExcelFunctionTester : ExcelFunction
        {
            public override int ArgumentMinLength => 0;
            public IEnumerable<ExcelDoubleCellValue> ArgsToDoubleEnumerableImpl(IEnumerable<FunctionArgument> args)
            {
                return ArgsToDoubleEnumerable(args, ParsingContext.Create());
            }
            #region Other members
            public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
            {
                throw new NotImplementedException();
            }
            #endregion
        }

        [TestMethod]
        public void ArgsToDoubleEnumerableShouldHandleInnerEnumerables()
        {
            var args = FunctionsHelper.CreateArgs(1, 2, InMemoryRange.GetFromArray(3, 4));
            var tester = new ExcelFunctionTester();
            var result = tester.ArgsToDoubleEnumerableImpl(args);
            Assert.AreEqual(4, result.Count());
        }
    }
}

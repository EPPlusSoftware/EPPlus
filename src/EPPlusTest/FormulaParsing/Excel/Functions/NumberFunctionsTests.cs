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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Numeric;
using EPPlusTest.FormulaParsing.TestHelpers;

namespace EPPlusTest.Excel.Functions
{
    [TestClass]
    public class NumberFunctionsTests
    {
        private ParsingContext _parsingContext = ParsingContext.Create();

        [TestMethod]
        public void CIntShouldConvertTextToInteger()
        {
            var func = new CInt();
            var args = FunctionsHelper.CreateArgs("2");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2, result.Result);
        }

        [TestMethod]
        public void IntShouldConvertDecimalToInteger()
        {
            var func = new CInt();
            var args = FunctionsHelper.CreateArgs(2.88m);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2, result.Result);
        }

        [TestMethod]
        public void IntShouldConvertNegativeDecimalToInteger()
        {
            var func = new CInt();
            var args = FunctionsHelper.CreateArgs(-2.88m);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(-3, result.Result);
        }

        [TestMethod]
        public void IntShouldConvertStringToInteger()
        {
            var func = new CInt();
            var args = FunctionsHelper.CreateArgs("-2.88");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(-3, result.Result);
        }
    }
}

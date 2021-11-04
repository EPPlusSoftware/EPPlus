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
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace EPPlusTest.FormulaParsing.IntegrationTests
{
    [TestClass]
    public class OperatorsTests
    {
        private ExcelPackage _package;
        private ExcelWorksheet _ws;
        private readonly ExcelErrorValue DivByZero = ExcelErrorValue.Create(eErrorType.Div0);

        [TestInitialize]
        public void Initialize()
        {
            _package = new ExcelPackage();
            _ws = _package.Workbook.Worksheets.Add("test");
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }

        [TestMethod]
        public void DivByZeroShouldReturnError()
        {
            var result = _ws.Calculate("10/0 + 3");
            Assert.AreEqual(DivByZero, result);
        }

        [TestMethod]
        public void ConcatShouldUseFormatG15()
        {
            var result = _ws.Calculate("14.000000000000002 & \"%\"");
            Assert.AreEqual("14%", result);

            result = _ws.Calculate("\"%\" & 14.000000000000002");
            Assert.AreEqual("%14", result);
        }
    }
}

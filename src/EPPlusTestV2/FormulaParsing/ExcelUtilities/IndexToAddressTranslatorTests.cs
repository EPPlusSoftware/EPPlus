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
using FakeItEasy;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace EPPlusTest.ExcelUtilities
{
    [TestClass]
    public class IndexToAddressTranslatorTests
    {
        private ExcelDataProvider _excelDataProvider;
        private IndexToAddressTranslator _indexToAddressTranslator;

        [TestInitialize]
        public void Setup()
        {
            SetupTranslator(12345, ExcelReferenceType.RelativeRowAndColumn);
        }

        private void SetupTranslator(int maxRows, ExcelReferenceType refType)
        {
            _excelDataProvider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => _excelDataProvider.ExcelMaxRows).Returns(maxRows);
            _indexToAddressTranslator = new IndexToAddressTranslator(_excelDataProvider, refType);
        }

        [TestMethod, ExpectedException(typeof(ArgumentNullException))]
        public void ShouldThrowIfExcelDataProviderIsNull()
        {
            new IndexToAddressTranslator(null);
        }

        [TestMethod]
        public void ShouldTranslate1And1ToA1()
        {
            var result = _indexToAddressTranslator.ToAddress(1, 1);
            Assert.AreEqual("A1", result);
        }

        [TestMethod]
        public void ShouldTranslate27And1ToAA1()
        {
            var result = _indexToAddressTranslator.ToAddress(27, 1);
            Assert.AreEqual("AA1", result);
        }

        [TestMethod]
        public void ShouldTranslate53And1ToBA1()
        {
            var result = _indexToAddressTranslator.ToAddress(53, 1);
            Assert.AreEqual("BA1", result);
        }

        [TestMethod]
        public void ShouldTranslate702And1ToZZ1()
        {
            var result = _indexToAddressTranslator.ToAddress(702, 1);
            Assert.AreEqual("ZZ1", result);
        }

        [TestMethod]
        public void ShouldTranslate703ToAAA4()
        {
            var result = _indexToAddressTranslator.ToAddress(703, 4);
            Assert.AreEqual("AAA4", result);
        }

        [TestMethod]
        public void ShouldTranslateToEntireColumnWhenRowIsEqualToMaxRows()
        {
            A.CallTo(() => _excelDataProvider.ExcelMaxRows).Returns(123456);
            var result = _indexToAddressTranslator.ToAddress(1, 123456);
            Assert.AreEqual("A", result);
        }

        [TestMethod]
        public void ShouldTranslateToAbsoluteAddress()
        {
            SetupTranslator(123456, ExcelReferenceType.AbsoluteRowAndColumn);
            var result = _indexToAddressTranslator.ToAddress(1, 1);
            Assert.AreEqual("$A$1", result);
        }

        [TestMethod]
        public void ShouldTranslateToAbsoluteRowAndRelativeCol()
        {
            SetupTranslator(123456, ExcelReferenceType.AbsoluteRowRelativeColumn);
            var result = _indexToAddressTranslator.ToAddress(1, 1);
            Assert.AreEqual("A$1", result);
        }

        [TestMethod]
        public void ShouldTranslateToRelativeRowAndAbsoluteCol()
        {
            SetupTranslator(123456, ExcelReferenceType.RelativeRowAbsoluteColumn);
            var result = _indexToAddressTranslator.ToAddress(1, 1);
            Assert.AreEqual("$A1", result);
        }

        [TestMethod]
        public void ShouldTranslateToRelativeRowAndCol()
        {
            SetupTranslator(123456, ExcelReferenceType.RelativeRowAndColumn);
            var result = _indexToAddressTranslator.ToAddress(1, 1);
            Assert.AreEqual("A1", result);
        }
    }
}

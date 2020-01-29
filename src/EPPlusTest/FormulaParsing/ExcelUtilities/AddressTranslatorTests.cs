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
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing;
using FakeItEasy;

namespace EPPlusTest.ExcelUtilities
{
    [TestClass]
    public class AddressTranslatorTests
    {
        private AddressTranslator _addressTranslator;
        private ExcelDataProvider _excelDataProvider;
        private const int ExcelMaxRows = 1356;

        [TestInitialize]
        public void Setup()
        {
            _excelDataProvider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => _excelDataProvider.ExcelMaxRows).Returns(ExcelMaxRows);
            _addressTranslator = new AddressTranslator(_excelDataProvider);
        }

        [TestMethod, ExpectedException(typeof(ArgumentNullException))]
        public void ConstructorShouldThrowIfProviderIsNull()
        {
            new AddressTranslator(null);
        }

        [TestMethod]
        public void ShouldTranslateRowNumber()
        {
            int col, row;
            _addressTranslator.ToColAndRow("A2", out col, out row);
            Assert.AreEqual(2, row);
        }

        [TestMethod]
        public void ShouldTranslateLettersToColumnIndex()
        {
            int col, row;
            _addressTranslator.ToColAndRow("C1", out col, out row);
            Assert.AreEqual(3, col);
            _addressTranslator.ToColAndRow("AA2", out col, out row);
            Assert.AreEqual(27, col);
            _addressTranslator.ToColAndRow("BC1", out col, out row);
            Assert.AreEqual(55, col);
        }

        [TestMethod]
        public void ShouldTranslateLetterAddressUsingMaxRowsFromProviderLower()
        {
            int col, row;
            _addressTranslator.ToColAndRow("A", out col, out row);
            Assert.AreEqual(1, row);
        }

        [TestMethod]
        public void ShouldTranslateLetterAddressUsingMaxRowsFromProviderUpper()
        {
            int col, row;
            _addressTranslator.ToColAndRow("A", out col, out row, AddressTranslator.RangeCalculationBehaviour.LastPart);
            Assert.AreEqual(ExcelMaxRows, row);
        }
    }
}

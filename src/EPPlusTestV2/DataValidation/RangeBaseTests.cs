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

namespace EPPlusTest.DataValidation
{
    [TestClass]
    public class RangeBaseTests : ValidationTestBase
    {
        [TestInitialize]
        public void Setup()
        {
            SetupTestData();
        }

        [TestCleanup]
        public void Cleanup()
        {
            CleanupTestData();
        }

        [TestMethod]
        public void RangeBase_AddIntegerValidation_ValidationIsAdded()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddIntegerDataValidation();

            // Assert
            Assert.AreEqual(1, _sheet.DataValidations.Count);
        }

        [TestMethod]
        public void RangeBase_AddIntegerValidation_AddressIsCorrect()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddIntegerDataValidation();

            // Assert
            Assert.AreEqual("A1:A2", _sheet.DataValidations[0].Address.Address);
        }

        [TestMethod]
        public void RangeBase_AddDecimalValidation_ValidationIsAdded()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddDecimalDataValidation();

            // Assert
            Assert.AreEqual(1, _sheet.DataValidations.Count);
        }

        [TestMethod]
        public void RangeBase_AddDecimalValidation_AddressIsCorrect()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddDecimalDataValidation();

            // Assert
            Assert.AreEqual("A1:A2", _sheet.DataValidations[0].Address.Address);
        }

        [TestMethod]
        public void RangeBase_AddTextLengthValidation_ValidationIsAdded()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddTextLengthDataValidation();

            // Assert
            Assert.AreEqual(1, _sheet.DataValidations.Count);
        }

        [TestMethod]
        public void RangeBase_AddTextLengthValidation_AddressIsCorrect()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddTextLengthDataValidation();

            // Assert
            Assert.AreEqual("A1:A2", _sheet.DataValidations[0].Address.Address);
        }

        [TestMethod]
        public void RangeBase_AddDateTimeValidation_ValidationIsAdded()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddDateTimeDataValidation();

            // Assert
            Assert.AreEqual(1, _sheet.DataValidations.Count);
        }

        [TestMethod]
        public void RangeBase_AddDateTimeValidation_AddressIsCorrect()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddDateTimeDataValidation();

            // Assert
            Assert.AreEqual("A1:A2", _sheet.DataValidations[0].Address.Address);
        }

        [TestMethod]
        public void RangeBase_AddListValidation_ValidationIsAdded()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddListDataValidation();

            // Assert
            Assert.AreEqual(1, _sheet.DataValidations.Count);
        }

        [TestMethod]
        public void RangeBase_AddListValidation_AddressIsCorrect()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddListDataValidation();

            // Assert
            Assert.AreEqual("A1:A2", _sheet.DataValidations[0].Address.Address);
        }

        [TestMethod]
        public void RangeBase_AdTimeValidation_ValidationIsAdded()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddTimeDataValidation();

            // Assert
            Assert.AreEqual(1, _sheet.DataValidations.Count);
        }

        [TestMethod]
        public void RangeBase_AddTimeValidation_AddressIsCorrect()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddTimeDataValidation();

            // Assert
            Assert.AreEqual("A1:A2", _sheet.DataValidations[0].Address.Address);
        }
    }
}

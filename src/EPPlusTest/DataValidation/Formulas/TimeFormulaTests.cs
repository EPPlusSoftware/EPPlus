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
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.DataValidation;

namespace EPPlusTest.DataValidation.Formulas
{
    [TestClass]
    public class TimeFormulaTests : ValidationTestBase
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
        public void TimeFormula_ValueIsSetFromConstructorValidateHour()
        {
            // Arrange
            var time = new ExcelTime(0.675M);
            LoadXmlTestData("A1", "time", "0.675");

            // Act
            var formula = new ExcelDataValidationTime(ExcelDataValidation.NewId(), "A1");

            // Assert
            Assert.AreEqual(time.Hour, formula.Formula.Value.Hour);
        }

        [TestMethod]
        public void TimeFormula_ValueIsSetFromConstructorValidateMinute()
        {
            // Arrange
            var time = new ExcelTime(0.395M);
            LoadXmlTestData("A1", "time", "0.395");

            // Act
            var formula = new ExcelDataValidationTime(ExcelDataValidation.NewId(), "A1");

            // Assert
            Assert.AreEqual(time.Minute, formula.Formula.Value.Minute);
        }

        [TestMethod]
        public void TimeFormula_ValueIsSetFromConstructorValidateSecond()
        {
            // Arrange
            var time = new ExcelTime(0.812M);
            LoadXmlTestData("A1", "time", "0.812");

            // Act
            var formula = new ExcelDataValidationTime(ExcelDataValidation.NewId(), "A1");

            // Assert
            Assert.AreEqual(time.Second.Value, formula.Formula.Value.Second.Value);
        }
    }
}

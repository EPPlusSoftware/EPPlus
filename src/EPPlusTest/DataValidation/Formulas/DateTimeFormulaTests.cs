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
using System;

namespace EPPlusTest.DataValidation.Formulas
{
    [TestClass]
    public class DateTimeFormulaTests : ValidationTestBase
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
            _dataValidationNode = null;
        }

        [TestMethod]
        public void DateTimeFormula_FormulaValueIsSetFromXmlNodeInConstructor()
        {
            // Arrange
            var date = DateTime.Parse("2011-01-08");
            var dateAsString = date.ToOADate().ToString(_cultureInfo);
            LoadXmlTestData("A1", "decimal", dateAsString);
            // Act
            var validation = new ExcelDataValidationDateTime(ExcelDataValidation.NewId(), "A1");
            // Assert
            Assert.AreEqual(date, validation.Formula.Value);
        }

        [TestMethod]
        public void DateTimeFormula_FormulasFormulaIsSetFromXmlNodeInConstructor()
        {
            // Arrange
            var date = DateTime.Parse("2011-01-08");
            LoadXmlTestData("A1", "decimal", "A1");

            // Act
            var validation = new ExcelDataValidationDateTime(ExcelDataValidation.NewId(), "A1");

            // Assert
            Assert.AreEqual("A1", validation.Formula.ExcelFormula);
        }
    }
}

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

namespace EPPlusTest.FormulaParsing.IntegrationTests.ErrorHandling
{
    /// <summary>
    /// Summary description for SumTests
    /// </summary>
    [TestClass, Ignore]
    public class SumTests : FormulaErrorHandlingTestBase
    {
        [TestInitialize]
        public void ClassInitialize()
        {
            BaseInitialize();
        }

        [TestCleanup]
        public void ClassCleanup()
        {
            BaseCleanup();
        }

        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }
        [TestMethod]
        public void SingleCell()
        {
            Assert.AreEqual(3d, Worksheet.Cells["B9"].Value);
        }

        [TestMethod]
        public void MultiCell()
        {
            Assert.AreEqual(40d, Worksheet.Cells["C9"].Value);
        }

        [TestMethod]
        public void Name()
        {
            Assert.AreEqual(10d, Worksheet.Cells["E9"].Value);
        }

        [TestMethod]
        public void ReferenceError()
        {
            Assert.AreEqual("#REF!", Worksheet.Cells["H9"].Value.ToString());
        }

        [TestMethod]
        public void NameOnOtherSheet()
        {
            Assert.AreEqual(130d, Worksheet.Cells["I9"].Value);
        }

        [TestMethod]
        public void ArrayInclText()
        {
            Assert.AreEqual(7d, Worksheet.Cells["J9"].Value);
        }

        //[TestMethod]
        //public void NameError()
        //{
        //    Assert.AreEqual("#NAME?", Worksheet.Cells["L9"].Value);
        //}

        //[TestMethod]
        //public void DivByZeroError()
        //{
        //    Assert.AreEqual("#DIV/0!", Worksheet.Cells["M9"].Value);
        //}
    }
}

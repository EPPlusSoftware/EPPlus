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
using OfficeOpenXml;
using OfficeOpenXml.Core;
using System;
using System.Collections.Generic;
using System.Text;

namespace EPPlusTest.Core
{
    /// <summary>
    /// All of these tests relate from cell B3
    /// </summary>
    [TestClass]
    public class R1C1TranslatorTest
    {
        [TestMethod]
        public void R()
        {
            var r1c1 = "R";
            var expectedAddress = "3:3";
            AssertAddresses(r1c1, expectedAddress);
        }

        [TestMethod]
        public void C()
        {
            var r1c1 = "C";
            var expectedAddress = "B:B";
            AssertAddresses(r1c1, expectedAddress);
        }
        [TestMethod]
        public void C3()
        {
            var r1c1 = "C3";
            var expectedAddress = "$C:$C";
            AssertAddresses(r1c1, expectedAddress);
        }
        [TestMethod]
        public void RC()
        {
            var r1c1 = "RC";
            var expectedAddress = "B3";
            AssertAddresses(r1c1, expectedAddress);
        }
        [TestMethod]
        public void R1C()
        {
            var r1c1 = "R1C";
            var expectedAddress = "B$1";
            AssertAddresses(r1c1, expectedAddress);
        }
        [TestMethod]
        public void R_1_C()
        {
            var r1c1 = "R[1]C";
            var expectedAddress = "B4";
            var address = R1C1Translator.FromR1C1(r1c1, 3, 3);
            AssertAddresses(r1c1, expectedAddress);
        }
        [TestMethod]
        public void R_Minus_1_C()
        {
            var r1c1 = "R[-1]C";
            var expectedAddress = "B2";
            AssertAddresses(r1c1, expectedAddress);
        }
        [TestMethod]
        public void RC_1()
        {
            var r1c1 = "RC[1]";
            var expectedAddress = "C3";
            AssertAddresses(r1c1, expectedAddress);
        }
        [TestMethod]
        public void RC_Minus_1()
        {
            var r1c1 = "RC[-1]";
            var expectedAddress = "A3";
            AssertAddresses(r1c1, expectedAddress);
        }
        [TestMethod]
        public void C_10_C_12()
        {
            var r1c1 = "C9:C10";
            var expectedAddress = "$I:$J";
            AssertAddresses(r1c1, expectedAddress);
        }
        [TestMethod]
        public void C1_C__1()
        {
            var r1c1 = "C1:C[-1]";
            var expectedAddress = "$A:A";
            AssertAddresses(r1c1, expectedAddress);
        }
        [TestMethod]
        public void C_1_C10()
        {
            var r1c1 = "C[1]:C10";
            var expectedAddress = "C:$J";
            AssertAddresses(r1c1, expectedAddress);
        }
        [TestMethod]
        public void RC_Minus_1_R_2_C_3()
        {
            var r1c1 = "RC[-1]:R[2]C[3]";
            var expectedAddress = "A3:E5";
            AssertAddresses(r1c1, expectedAddress);
        }
        [TestMethod]
        public void TranslateC1FullColumnWithSheet()
        {
            const string formula = "SUM(Sheet1!A:A)";
            var formulaR1C1 = ExcelCellBase.TranslateToR1C1(formula, 1, 2);
            Assert.AreEqual("SUM(Sheet1!C[-1])", formulaR1C1); // fails: formulaR1C1 == "Sum(C[-1])"
        }
        [TestMethod]
        public void TranslateRCFullColumnWithSheet()
        {
            const string formulaR1C1 = "SUM(Sheet1!C[-1])";
            var formula = ExcelCellBase.TranslateFromR1C1(formulaR1C1, 1, 2);
            Assert.AreEqual("SUM(Sheet1!A:A)", formula); // also fails: formula == "Sum(A:A)"
        }
        private static void AssertAddresses(string r1c1, string expectedAddress)
        {
            var address = R1C1Translator.FromR1C1(r1c1, 3, 2);  //From Cell B2
            Assert.AreEqual(expectedAddress, address);
            address = R1C1Translator.ToR1C1(new ExcelAddressBase(address), 3, 2); //From Cell B2
            Assert.AreEqual(r1c1, address);
        }
    }
}

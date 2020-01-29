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

namespace EPPlusTest.ExcelUtilities
{
    [TestClass]
    public class ExcelAddressInfoTests
    {
        [TestMethod, ExpectedException(typeof(ArgumentException))]
        public void ParseShouldThrowIfAddressIsNull()
        {
            ExcelAddressInfo.Parse(null);
        }

        [TestMethod]
        public void ParseShouldSetWorksheet()
        {
            var info = ExcelAddressInfo.Parse("Worksheet!A1");
            Assert.AreEqual("Worksheet", info.Worksheet);
        }

        [TestMethod]
        public void WorksheetIsSpecifiedShouldBeTrueWhenWorksheetIsSupplied()
        {
            var info = ExcelAddressInfo.Parse("Worksheet!A1");
            Assert.IsTrue(info.WorksheetIsSpecified);
        }

        [TestMethod]
        public void ShouldIndicateMultipleCellsWhenAddressContainsAColon()
        {
            var info = ExcelAddressInfo.Parse("A1:A2");
            Assert.IsTrue(info.IsMultipleCells);
        }

        [TestMethod]
        public void ShouldSetStartCell()
        {
            var info = ExcelAddressInfo.Parse("A1:A2");
            Assert.AreEqual("A1", info.StartCell);
        }

        [TestMethod]
        public void ShouldSetEndCell()
        {
            var info = ExcelAddressInfo.Parse("A1:A2");
            Assert.AreEqual("A2", info.EndCell);
        }

        [TestMethod]
        public void ParseShouldSetAddressOnSheet()
        {
            var info = ExcelAddressInfo.Parse("Worksheet!A1:A2");
            Assert.AreEqual("A1:A2", info.AddressOnSheet);
        }

        [TestMethod]
        public void AddressOnSheetShouldBeSameAsAddressIfNoWorksheetIsSpecified()
        {
            var info = ExcelAddressInfo.Parse("A1:A2");
            Assert.AreEqual("A1:A2", info.AddressOnSheet);
        }
    }
}

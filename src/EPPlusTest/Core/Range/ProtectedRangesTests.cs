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
using EPPlusTest;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace OfficeOpenXml.Core.Range
{
    [TestClass]
    public class ProtectedRangesTests : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("ProtectedRanges.xlsx", true);
        }
        [ClassCleanup]
        public static void CleanUp()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void AddProtectedRange()
        {
            var ws = _pck.Workbook.Worksheets.Add("SingleProtectedRange");
            LoadTestdata(ws);
            ws.ProtectedRanges.Add("ProtectedRange", ws.Cells["A1:B2"]);
            //Assert
            Assert.AreEqual(1, ws.ProtectedRanges.Count);
            Assert.AreEqual("ProtectedRange", ws.ProtectedRanges[0].Name);
            Assert.AreEqual("A1:B2", ws.ProtectedRanges[0].Address.Address);
        }
        [TestMethod]
        public void AddThreeProtectedRanges()
        {
            var ws = _pck.Workbook.Worksheets.Add("ThreeProtectedRanges");
            LoadTestdata(ws);
            var pr1 = ws.ProtectedRanges.Add("ProtectedRange1", ws.Cells["A1:B2"]);
            pr1.SetPassword("EPPlus");
            var pr2 = ws.ProtectedRanges.Add("ProtectedRange2", ws.Cells["C1:D2"]);
            pr2.SetPassword("EPPlus2");
            var pr3 = ws.ProtectedRanges.Add("ProtectedRange3", ws.Cells["B1:E8"]);
            //Assert
            pr3.SetPassword("EPPlus3");

            Assert.AreEqual(3, ws.ProtectedRanges.Count);
            Assert.AreEqual("ProtectedRange1", ws.ProtectedRanges[0].Name);
            Assert.AreEqual("ProtectedRange2", ws.ProtectedRanges[1].Name);
            Assert.AreEqual("ProtectedRange3", ws.ProtectedRanges[2].Name);
            Assert.AreEqual("A1:B2", pr1.Address.Address);
            Assert.AreEqual("C1:D2", pr2.Address.Address);
            Assert.AreEqual("B1:E8", pr3.Address.Address);
        }
        [TestMethod]
        public void RemoveProtectedRange()
        {
            using (var pck = new ExcelPackage())
            {
                var ws = pck.Workbook.Worksheets.Add("RemoveProtectedRange");
                ws.ProtectedRanges.Add("ProtectedRange", ws.Cells["A1:B2"]);
                Assert.AreEqual(1, ws.ProtectedRanges.Count);
                ws.ProtectedRanges.Remove(ws.ProtectedRanges[0]);
                //Assert
                Assert.AreEqual(0, ws.ProtectedRanges.Count);
                Assert.IsFalse(ws.ProtectedRanges.ExistNode("d:protectedRanges"));
            }
        }
        [TestMethod]
        public void RemoveAtProtectedRange()
        {
            using (var pck = new ExcelPackage())
            {
                var ws = pck.Workbook.Worksheets.Add("RemoveAtProtectedRange");
                ws.ProtectedRanges.Add("ProtectedRange", ws.Cells["A1:B2"]);
                Assert.AreEqual(1, ws.ProtectedRanges.Count);
                ws.ProtectedRanges.RemoveAt(0);
                //Assert
                Assert.AreEqual(0, ws.ProtectedRanges.Count);
                Assert.IsFalse(ws.ProtectedRanges.ExistNode("d:protectedRanges"));
            }
        }
        [TestMethod]
        public void ClearProtectedRange()
        {
            using (var pck = new ExcelPackage())
            {
                var ws = pck.Workbook.Worksheets.Add("ClearProtectedRange");
                ws.ProtectedRanges.Add("ProtectedRange1", ws.Cells["A1:B2"]);
                ws.ProtectedRanges.Add("ProtectedRange2", ws.Cells["A2:B3"]);
                ws.ProtectedRanges.Add("ProtectedRange3", ws.Cells["A3:B4"]);
                Assert.AreEqual(3, ws.ProtectedRanges.Count);
                ws.ProtectedRanges.Clear();
                //Assert
                Assert.AreEqual(0, ws.ProtectedRanges.Count);
                Assert.IsFalse(ws.ProtectedRanges.ExistNode("d:protectedRanges"));
            }
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void DuplicateNameShouldThrowInvalidOperationException()
        {
            using (var pck = new ExcelPackage())
            {
                var ws = pck.Workbook.Worksheets.Add("ThreeProtectedRanges");
                ws.ProtectedRanges.Add("Range", ws.Cells["A1:B2"]);
                ws.ProtectedRanges.Add("range", ws.Cells["A4:B5"]);
            }
        }

    }
}

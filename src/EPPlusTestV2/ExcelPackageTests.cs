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
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlusTest
{
    [TestClass]
    public class ExcelPackageTests
    {
        [TestMethod, Ignore]
        public void ConstructorWithStringPath()
        {
            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Test.xlsx");
            using(var package = new ExcelPackage(path))
            {

            }
        }

        [TestMethod, Ignore]
        public void ConstructorWithStringPathAndPassword()
        {
            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Test.xlsx");
            using (var package = new ExcelPackage(path, "pwd123"))
            {

            }
        }

        [DataTestMethod]
        [DataRow(EncryptionAlgorithm.AES128)]
        [DataRow(EncryptionAlgorithm.AES192)]
        [DataRow(EncryptionAlgorithm.AES256)]
        public void ShouldEncryptAndDecryptPackage(EncryptionAlgorithm algorithm)
        {
            byte[] bytes;
            var pwd = "pwd123";
            using (var ms = new MemoryStream())
            { 
                using (var encryptedPackage = new ExcelPackage())
                {
                    encryptedPackage.Encryption.Algorithm = algorithm;
                    var sheet = encryptedPackage.Workbook.Worksheets.Add("Sheet1");
                    sheet.Cells["A1"].Value = 1;
                    encryptedPackage.SaveAs(ms, pwd);
                    bytes = ms.ToArray();
                }
            }
            using(var ms2 = new MemoryStream(bytes))
            {
                using (var decryptedPackage = new ExcelPackage(ms2, pwd))
                {
                    var sheet = decryptedPackage.Workbook.Worksheets.First();
                    Assert.AreEqual(1d, sheet.Cells["A1"].Value);
                }
            }
        }
    }
}

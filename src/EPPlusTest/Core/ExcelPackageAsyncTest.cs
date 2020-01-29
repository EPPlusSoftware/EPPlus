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
using System.IO;
using System.Threading.Tasks;

namespace EPPlusTest.Core
{
    /************************************************************************************************************************
     * Note that some of these tests will fail the first time they are runned as they read the files created by other tests.
     ************************************************************************************************************************/
    [TestClass]
    public class ExcelPackageAsyncTest : TestBase
    {
        private const int noRows= 10000;

        public ExcelPackageAsyncTest()
        {
        }

        public static void CopyRead(FileInfo file)
        {
            var dirName = file.DirectoryName;
            var fileName = file.FullName;

            File.Copy(fileName, dirName + $"\\{file.Name.Substring(0, file.Name.Length-file.Extension.Length)}Read.xlsx", true);
        }

        [TestMethod]
        public async Task SaveAsyncTest()
        {
            using (var pck = OpenPackage("Async.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("AsyncSave");
                LoadTestdata(ws, noRows);
                await pck.SaveAsync().ConfigureAwait(false);
                CopyRead(pck.File);
            }
        }
        [TestMethod]
        public async Task SaveAsyncEncryptedTest()
        {
            using (var pck = OpenPackage("AsyncEnc.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("AsyncEncryptedSave");
                LoadTestdata(ws, noRows);
                await pck.SaveAsync("EPPlus").ConfigureAwait(false);
                CopyRead(pck.File);
            }
        }
        [TestMethod]
        public async Task LoadAsyncTest()
        {
            AssertIfNotExists("AsyncRead.xlsx");
            using (var pck = await OpenPackageAsync("AsyncRead.xlsx").ConfigureAwait(false))
            {
                var ws = TryGetWorksheet(pck, "AsyncSave");
                Assert.AreEqual($"A1:D{noRows}", ws.Dimension.Address);
            }
        }
        [TestMethod]
        public async Task LoadAsyncEncryptedTest()
        {
            AssertIfNotExists("AsyncEncRead.xlsx");
            using (var pck = await OpenPackageAsync("AsyncEncRead.xlsx", false, "EPPlus").ConfigureAwait(false))
            {
                var ws = TryGetWorksheet(pck, "AsyncEncryptedSave");
                Assert.AreEqual($"A1:D{noRows}", ws.Dimension.Address);
            }
        }
        [TestMethod]
        public async Task GetAsByteArrayLoadStreamTest()
        {
            AssertIfNotExists("AsyncRead.xlsx");
            using (var pck = await OpenPackageAsync("AsyncRead.xlsx").ConfigureAwait(false))
            {
                var ws = TryGetWorksheet(pck, "AsyncSave");

                var b = await pck.GetAsByteArrayAsync();
                var ms = new MemoryStream(b);

                var pck2 = new ExcelPackage();
                await pck2.LoadAsync(ms);
                ws = TryGetWorksheet(pck2, "AsyncSave");
                Assert.AreEqual($"A1:D{noRows}", ws.Dimension.Address);
            }
        }
        [TestMethod]
        public async Task GetAsByteArrayEncryptedLoadStreamEncryptedTest()
        {
            AssertIfNotExists("AsyncEncRead.xlsx");
            var password = "EPPlus";
            using (var pck = await OpenPackageAsync("AsyncEncRead.xlsx", false, password).ConfigureAwait(false))
            {
                var ws = TryGetWorksheet(pck, "AsyncEncryptedSave");
                Assert.AreEqual($"A1:D{noRows}", ws.Dimension.Address);

                var b = await pck.GetAsByteArrayAsync(password);

                var ms = new MemoryStream(b);
                var pck2 = new ExcelPackage();

                await pck2.LoadAsync(ms, password);
                ws = TryGetWorksheet(pck2, "AsyncEncryptedSave");
                Assert.AreEqual($"A1:D{noRows}", ws.Dimension.Address);
            }
        }
    }
}

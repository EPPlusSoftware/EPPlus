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
using EPPlusTest.Drawing.Chart.Styling;
using FakeItEasy.Configuration;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Constants;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest
{
    [TestClass]
    public class ExcelPackageTests : TestBase
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

        [TestMethod]
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

        [TestMethod]
        public void SaveAsTemplate_WithVba_SetsTemplateMacroEnabledContentType()
        {
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = "Test";

                package.Workbook.CreateVBAProject();
                package.SaveAsTemplate = true;
                SaveWorkbook("SaveTemplate_Vba.xltm", package);
            }

            AssertSavedWorkbookContentType("SaveTemplate_Vba.xltm",
                ContentTypes.contentTypeTemplateMacroEnabled);
        }

        [TestMethod]
        public void ReadTemplate_ResaveWithoutOptions_ConvertsToWorkbook()
        {
            using (var package = OpenTemplatePackage("ExcelTemplateTest.xltx"))
            {
                SaveWorkbook("ReadTemplate_Resaved.xlsx", package);
            }

            AssertSavedWorkbookContentType("ReadTemplate_Resaved.xlsx",
                ContentTypes.contentTypeWorkbookDefault);
        }

        [TestMethod]
        public void SaveAs_FileInfo_WithOptions_WritesTemplateContentType()
        {
            var file = new FileInfo(Path.Combine(_worksheetPath, "SaveTemplate_FileInfo.xltx"));

            using (var package = new ExcelPackage())
            {
                package.Workbook.Worksheets.Add("Sheet1");
                package.SaveAs(file, o => o.SaveAsTemplate = true);
            }

            using (var reopened = new ExcelPackage(file))
            {
                AssertWorkbookContentType(ContentTypes.contentTypeTemplateDefault, reopened);
            }
        }

        [TestMethod]
        public void SaveAs_Stream_WithOptions_WritesTemplateContentType()
        {
            using (var ms = new MemoryStream())
            {
                using (var package = new ExcelPackage())
                {
                    package.Workbook.Worksheets.Add("Sheet1");
                    package.SaveAs(ms, o => o.SaveAsTemplate = true);
                }

                ms.Position = 0;
                using (var reopened = new ExcelPackage(ms))
                {
                    AssertWorkbookContentType(ContentTypes.contentTypeTemplateDefault, reopened);
                }
            }
        }

        [TestMethod]
        public void SaveAs_FilePath_WithOptions_WritesTemplateContentType()
        {
            var path = Path.Combine(_worksheetPath, "SaveTemplate_Path.xltx");

            using (var package = new ExcelPackage())
            {
                package.Workbook.Worksheets.Add("Sheet1");
                package.SaveAs(path, o => o.SaveAsTemplate = true);
            }

            using (var reopened = new ExcelPackage(new FileInfo(path)))
            {
                AssertWorkbookContentType(ContentTypes.contentTypeTemplateDefault, reopened);
            }
        }

        [TestMethod]
        public void Save_WithOptions_WritesTemplateContentType()
        {
            var file = new FileInfo(Path.Combine(_worksheetPath, "SaveTemplate_Save.xltx"));
            if (file.Exists) file.Delete();

            using (var package = new ExcelPackage(file))
            {
                package.Workbook.Worksheets.Add("TestSheet");
                package.Save(o => o.SaveAsTemplate = true);
            }

            using (var reopened = new ExcelPackage(file))
            {
                AssertWorkbookContentType(ContentTypes.contentTypeTemplateDefault, reopened);
            }
        }

        [TestMethod]
        public void SaveAs_WithOptions_DefaultsToWorkbookContentType()
        {
            using (var ms = new MemoryStream())
            {
                using (var package = new ExcelPackage())
                {
                    package.Workbook.Worksheets.Add("Sheet1");
                    package.SaveAs(ms, o => { });   
                }

                ms.Position = 0;
                using (var reopened = new ExcelPackage(ms))
                {
                    AssertWorkbookContentType(ContentTypes.contentTypeWorkbookDefault, reopened);
                }
            }
        }

        [TestMethod]
        public void ResaveLoadedTemplate_ConvertsToWorkbook()
        {
            using (var ms = new MemoryStream())
            {
                using (var package = OpenTemplatePackage("ExcelTemplateTest.xltx"))
                {
                    package.SaveAs(ms);
                }

                ms.Position = 0;
                using (var reopened = new ExcelPackage(ms))
                {
                    AssertWorkbookContentType(ContentTypes.contentTypeWorkbookDefault, reopened);
                }
            }
        }

        [TestMethod]
        public async Task SaveAsync_WithOptions_WritesTemplateContentType()
        {
            var file = new FileInfo(Path.Combine(_worksheetPath, "SaveTemplate_SaveAsync.xltx"));
            if (file.Exists) file.Delete();

            using (var package = new ExcelPackage(file))
            {
                package.Workbook.Worksheets.Add("Sheet1");
                await package.SaveAsync(o => o.SaveAsTemplate = true);
            }

            using (var reopened = new ExcelPackage(file))
            {
                AssertWorkbookContentType(ContentTypes.contentTypeTemplateDefault, reopened);
            }
        }

        [TestMethod]
        public async Task SaveAsAsync_FileInfo_WithOptions_WritesTemplateContentType()
        {
            var file = new FileInfo(Path.Combine(_worksheetPath, "SaveTemplate_AsyncFileInfo.xltx"));
            if (file.Exists) file.Delete();

            using (var package = new ExcelPackage())
            {
                package.Workbook.Worksheets.Add("Sheet1");
                await package.SaveAsAsync(file, o => o.SaveAsTemplate = true);
            }

            using (var reopened = new ExcelPackage(file))
            {
                AssertWorkbookContentType(ContentTypes.contentTypeTemplateDefault, reopened);
            }
        }

        [TestMethod]
        public async Task SaveAsAsync_FilePath_WithOptions_WritesTemplateContentType()
        {
            var path = Path.Combine(_worksheetPath, "SaveTemplate_AsyncPath.xltx");

            using (var package = new ExcelPackage())
            {
                package.Workbook.Worksheets.Add("Sheet1");
                await package.SaveAsAsync(path, o => o.SaveAsTemplate = true);
            }

            using (var reopened = new ExcelPackage(new FileInfo(path)))
            {
                AssertWorkbookContentType(ContentTypes.contentTypeTemplateDefault, reopened);
            }
        }

        [TestMethod]
        public async Task SaveAsAsync_Stream_WithOptions_WritesTemplateContentType()
        {
            using (var ms = new MemoryStream())
            {
                using (var package = new ExcelPackage())
                {
                    package.Workbook.Worksheets.Add("Sheet1");
                    await package.SaveAsAsync(ms, o => o.SaveAsTemplate = true);
                }

                ms.Position = 0;
                using (var reopened = new ExcelPackage(ms))
                {
                    AssertWorkbookContentType(ContentTypes.contentTypeTemplateDefault, reopened);
                }
            }
        }
        private void AssertSavedWorkbookContentType(string fileName, string expectedContentType)
        {
            using (var reopened = OpenPackage(fileName))
            {
                AssertWorkbookContentType(expectedContentType, reopened);
            }
        }

        private static void AssertWorkbookContentType(string expected, ExcelPackage package)
        {
            Assert.AreEqual(expected, package.Workbook.Part.ContentType);
        }
    }
}

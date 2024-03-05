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
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.
    Table;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Threading;

namespace EPPlusTest.Style
{
    [TestClass]
    public class RichTextTest : TestBase
    {
        static ExcelPackage _pck;

        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("RichText.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            var dirName = _pck.File.DirectoryName;
            var fileName = _pck.File.FullName;

            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void RichTextTest1()
        {
            var ws = _pck.Workbook.Worksheets.Add("RichTextSheet1");
            ws.Cells["A1"].Value = "Fint";
            var rt = ws.Cells["A1"].RichText.Add(" fint");
            var rt1 = ws.Cells["A1"].RichText.Add(" fint2");
            rt.Color = Color.Red;
            rt1.Color = Color.Green;
        }

        [TestMethod]
        public void RichTextReadTest()
        {
            using (var p = OpenTemplatePackage("RichTextNew.xlsx"))
            { 
                var ws = p.Workbook.Worksheets[0];
                Assert.AreEqual("Svart Röd Grön FET SNEMellan  Rum", ws.Cells["A1"].RichText.Text);
                Assert.AreEqual("Vad Som helst här", ws.Cells["A2"].Text);
                Assert.AreEqual("Vad Som helst här", ws.Cells["A2"].RichText.Text);
                ws.Cells["A1"].RichText[3].ColorSettings.Theme = eThemeSchemeColor.Accent5;
                SaveAndCleanup(p);
            }
        }
    }
}
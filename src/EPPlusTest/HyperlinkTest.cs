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
using System.Drawing;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;

namespace EPPlusTest
{
    [TestClass]
    public class HyperLinkTest : TestBase
    {
        static ExcelPackage _pck;
        static ExcelWorksheet _ws;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("Hyperlinks.xlsx", true);
            _ws = _pck.Workbook.Worksheets.Add("Hyperlink");
            var hl = _pck.Workbook.Styles.CreateNamedStyle("Hyperlink");
            hl.BuildInId = 8; //Hyperlink
            hl.Style.Font.UnderLine = true;
            hl.Style.Font.Color.Theme = eThemeSchemeColor.Hyperlink;
            _ws.Cells["A1:A3"].StyleName = "Hyperlink";
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void InsertNormalUri()
        {
            _ws.Cells["A1"].Hyperlink = new Uri("https://epplussoftware.com");
        }
        [TestMethod]
        public void InsertWorksheetLocalLink()
        {
            var hl = new ExcelHyperLink("A2", "A2");
            hl.Display = "Link to A2";
            _ws.Cells["A2"].Hyperlink = hl;
        }
        [TestMethod]
        public void InsertUriWithLocation()
        {
            var hl = new ExcelHyperLink("https://epplussoftware.com");
            hl.Display = "www.epplussoftware.com";
            hl.ReferenceAddress = "aa,bb=cc"; //Will set uri https://epplussoftware.com/#
            _ws.Cells["A3"].Hyperlink = hl;
        }
        [TestMethod]
        public void ReadUriWithLocation()
        {
            using (var p1 = new ExcelPackage())
            {
                var ws = p1.Workbook.Worksheets.Add("Sheet1");
                var hl = new ExcelHyperLink("https://epplussoftware.com");
                hl.Display = "www.epplussoftware.com";
                hl.ReferenceAddress = "aa,bb=cc"; //Will set uri https://epplussoftware.com/#
                ws.Cells["A1"].Hyperlink = hl;

                p1.Save();

                using (var p2 = new ExcelPackage(p1.Stream))
                {
                    var ws2 = p2.Workbook.Worksheets[0];

                    var hl2 = (ExcelHyperLink)ws2.Cells["A1"].Hyperlink;
                    Assert.AreEqual("https://epplussoftware.com", hl2.OriginalString);
                    Assert.AreEqual("www.epplussoftware.com", hl2.Display);
                    Assert.AreEqual("aa,bb=cc", hl2.ReferenceAddress);
                }
            }
        }

        [TestMethod]
        public void DrawingHyperlinkUpdate_ShouldNotThrowException()
        {
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Sheet1");
                
                // Add a drawing (shape)
                var shape = ws.Drawings.AddShape("MyShape", eShapeStyle.Rect);
                
                // 1. Assign initial external URL hyperlink
                var initialUrl = new ExcelHyperLink("https://epplussoftware.com") { ToolTip = "Initial ToolTip" };
                shape.Hyperlink = initialUrl;
                
                Assert.IsNotNull(shape.Hyperlink);
                
                // 2. Re-assign to a new external URL hyperlink (this would crash without the fix)
                var updatedUrl = new ExcelHyperLink("https://github.com/EPPlusSoftware/EPPlus") { ToolTip = "Updated ToolTip" };
                shape.Hyperlink = updatedUrl;
                
                Assert.AreEqual("https://github.com/EPPlusSoftware/EPPlus", shape.Hyperlink.OriginalString);
                
                // 3. Re-assign to an internal sheet reference (this would also crash without the fix)
                var internalLink = new ExcelHyperLink("Sheet1!A10", "Go to A10") { ToolTip = "Internal ToolTip" };
                shape.Hyperlink = internalLink;
                
                Assert.AreEqual("Sheet1!A10", ((ExcelHyperLink)shape.Hyperlink).ReferenceAddress);

                // 4. Save and read back to verify XML and tooltip are correct
                package.Save();

                using (var readPackage = new ExcelPackage(package.Stream))
                {
                    var readWs = readPackage.Workbook.Worksheets["Sheet1"];
                    var readShape = readWs.Drawings["MyShape"];
                    
                    var readLink = (ExcelHyperLink)readShape.Hyperlink;
                    Assert.IsNotNull(readLink);
                    Assert.AreEqual("Sheet1!A10", readLink.ReferenceAddress);
                    Assert.AreEqual("Internal ToolTip", readLink.ToolTip);
                }
            }
        }

    }
}

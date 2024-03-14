﻿/*******************************************************************************
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
using System.Drawing;

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
        public void RichTextPropertiesReadTest()
        {
            using (var p = OpenTemplatePackage("RichTextTests.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                //Test Normal Text
                Assert.AreEqual("This is just a string of poor text… :(", ws.Cells["A1"].Text);
                Assert.AreEqual("This is just a string of poor text… :(", ws.Cells["A1"].RichText[0].Text);
                //Test Bold
                Assert.AreEqual(true, ws.Cells["A2"].RichText[1].Bold);
                //Test Italic
                Assert.AreEqual(true, ws.Cells["A3"].RichText[1].Italic);
                //Test Strike
                Assert.AreEqual(true, ws.Cells["A4"].RichText[1].Strike);
                Assert.AreEqual(true, ws.Cells["A4"].RichText[3].Strike);
                //Test Vertical alignment
                Assert.AreEqual(ExcelVerticalAlignmentFont.Superscript, ws.Cells["A5"].RichText[1].VerticalAlign);
                Assert.AreEqual(ExcelVerticalAlignmentFont.Subscript, ws.Cells["A5"].RichText[3].VerticalAlign);
                //Test Size
                Assert.AreEqual(26, ws.Cells["A6"].RichText[1].Size);
                //Test Font
                Assert.AreEqual("Arial", ws.Cells["A7"].RichText[1].FontName);
                //Test Color
                Assert.AreEqual(Color.Purple.ToArgb(), ws.Cells["A8"].RichText[0].ColorSettings.Rgb.ToArgb());
                Assert.AreEqual(Color.Red.ToArgb(), ws.Cells["A8"].RichText[2].ColorSettings.Rgb.ToArgb());
                Assert.AreEqual(Color.Green.ToArgb(), ws.Cells["A8"].RichText[4].ColorSettings.Rgb.ToArgb());
                Assert.AreEqual(Color.Yellow.ToArgb(), ws.Cells["A8"].RichText[6].ColorSettings.Rgb.ToArgb());
                Assert.AreEqual(Color.Blue.ToArgb(), ws.Cells["A8"].RichText[7].ColorSettings.Rgb.ToArgb());
                //test Underline
                Assert.AreEqual(ExcelUnderLineType.Single, ws.Cells["A15"].RichText[0].UnderLineType);
                Assert.AreEqual(ExcelUnderLineType.Double, ws.Cells["A15"].RichText[2].UnderLineType);
            }
        }

        [TestMethod]
        public void RichTextPropertiesWriteTest()
        {
            using (var p = OpenTemplatePackage("RichTextTests.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                //Set bold, italic, strike and underline properties.
                ws.Cells["A18"].RichText[0].Bold = true;
                ws.Cells["A18"].RichText[0].Italic = true;
                ws.Cells["A18"].RichText[0].Strike = true;
                ws.Cells["A18"].RichText[0].UnderLineType = ExcelUnderLineType.Single;
                //Set vertical align properties
                ws.Cells["A17"].RichText[0].VerticalAlign = ExcelVerticalAlignmentFont.Subscript;
                ws.Cells["A18"].RichText[0].VerticalAlign = ExcelVerticalAlignmentFont.Superscript;
                //Set font properties
                ws.Cells["A9"].RichText[0].Charset = 161;
                ws.Cells["A18"].RichText[0].Size = 72;
                ws.Cells["A18"].RichText[0].FontName = "Arial";
                //Assign color properties
                ws.Cells["A17"].RichText[0].ColorSettings.Theme = eThemeSchemeColor.Accent5;
                ws.Cells["A18"].RichText[0].Color = Color.LightBlue;
                ws.Cells["A18"].RichText[0].ColorSettings.Indexed = 3;
                p.Save();
                using (var p2 = new ExcelPackage(p.Stream))
                {
                    var ws2 = p2.Workbook.Worksheets[0];
                    //Test reading bold, italic, strike and underline properties.
                    Assert.AreEqual(true, ws2.Cells["A18"].RichText[0].Bold);
                    Assert.AreEqual(true, ws2.Cells["A18"].RichText[0].Italic);
                    Assert.AreEqual(true, ws2.Cells["A18"].RichText[0].Strike);
                    Assert.AreEqual(ExcelUnderLineType.Single, ws2.Cells["A18"].RichText[0].UnderLineType);
                    //Test reading vertical align properties
                    Assert.AreEqual(ExcelVerticalAlignmentFont.Superscript, ws2.Cells["A18"].RichText[0].VerticalAlign);
                    Assert.AreEqual(ExcelVerticalAlignmentFont.Subscript, ws2.Cells["A17"].RichText[0].VerticalAlign);
                    //Test reading font properties
                    Assert.AreEqual(161, ws2.Cells["A9"].RichText[0].Charset);
                    Assert.AreEqual(72, ws2.Cells["A18"].RichText[0].Size);
                    Assert.AreEqual("Arial", ws2.Cells["A18"].RichText[0].FontName);
                    //Test reading color properties
                    Assert.AreEqual(eThemeSchemeColor.Accent5, ws2.Cells["A17"].RichText[0].ColorSettings.Theme);
                    Assert.AreEqual(Color.LightBlue.ToArgb(), ws2.Cells["A18"].RichText[0].Color.ToArgb());
                    Assert.AreEqual(Color.LightBlue.ToArgb(), ws2.Cells["A18"].RichText[0].ColorSettings.Rgb.ToArgb());
                    Assert.AreEqual(3, ws.Cells["A18"].RichText[0].ColorSettings.Indexed);
                }
                //Test Color Tint (tint applies to color so this one is tested separately)
                ws.Cells["A18"].RichText[0].ColorSettings.Tint = 0.5;
                p.Save();
                using (var p2 = new ExcelPackage(p.Stream))
                {
                    var ws2 = p2.Workbook.Worksheets[0];
                    Assert.AreEqual(0.5, ws2.Cells["A18"].RichText[0].ColorSettings.Tint);
                }
                //Test Color Auto property
                ws.Cells["A18"].RichText[0].ColorSettings.Auto = true;
                p.Save();
                using (var p2 = new ExcelPackage(p.Stream))
                {
                    var ws2 = p2.Workbook.Worksheets[0];
                    Assert.AreEqual(true, ws2.Cells["A18"].RichText[0].ColorSettings.Auto);
                }
            }
        }

        [TestMethod]
        public void RichTextPropertiesCopyTest()
        {
            using (var p = OpenTemplatePackage("RichTextTests.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                p.Workbook.Worksheets.Add("Page2");
                var ws2 = p.Workbook.Worksheets[1];
                ws.Cells["A1:A19"].Copy(ws2.Cells["B1:B19"]);
                p.Save();
                Assert.AreEqual(ws.Cells["A1"].Text, ws2.Cells["B1"].Text);
                Assert.AreEqual(ws.Cells["A2"].RichText[1].Bold, ws2.Cells["B2"].RichText[1].Bold);
                Assert.AreEqual(ws.Cells["A3"].RichText[1].Italic, ws2.Cells["B3"].RichText[1].Italic);
                Assert.AreEqual(ws.Cells["A8"].RichText[0].Color, ws2.Cells["B8"].RichText[0].Color);
                Assert.AreEqual(ws.Cells["A19"].Comment.RichText.Text, ws2.Cells["B19"].Comment.RichText.Text);
            }
        }

        [TestMethod]
        public void RichTextPropertiesCopyAndChangeTest()
        {
            using (var p = OpenTemplatePackage("RichTextTests.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                ws.Cells["A2"].Copy(ws.Cells["B2"]);
                ws.Cells["B2"].RichText.Text = "New Text Value";
                ws.Cells["A19"].Copy(ws.Cells["B19"]);
                ws.Cells["B19"].Comment.RichText.Text = "New Comment";
                ws.Cells["B19"].Comment.Author = "Merlin";
                p.Save();
                Assert.AreNotEqual(ws.Cells["B2"].RichText.Text, ws.Cells["A2"].RichText.Text);
                Assert.AreNotEqual(ws.Cells["B19"].Comment.RichText.Text, ws.Cells["A19"].Comment.RichText.Text);
            }
        }

        [TestMethod]
        public void RichTextWorkSheetCopy()
        {
            using (var p = OpenTemplatePackage("RichTextTests.xlsx"))
            {
                var ws = p.Workbook.Worksheets.Add("TargetSheet", p.Workbook.Worksheets[0]);
                ws.Cells["A18"].RichText.Text = "Something else";
                ws.Cells["A18"].RichText[0].Strike = true;
                p.Save();
                Assert.AreNotEqual(p.Workbook.Worksheets[0].Cells["A18"].RichText.Text, ws.Cells["A18"].RichText.Text);
            }
        }
    }
}

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
using System.Drawing;

namespace EPPlusTest.Style
{
    [TestClass]
    public class StylingTest : TestBase
    {
        static ExcelPackage _pck;

        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("Style.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void VerifyColumnStyle()
        {
            var ws=_pck.Workbook.Worksheets.Add("RangeStyle");
            LoadTestdata(ws, 100,2,2);

            ws.Row(3).Style.Fill.SetBackground(ExcelIndexedColor.Indexed5);
            ws.Column(3).Style.Fill.SetBackground(ExcelIndexedColor.Indexed7);
            ws.Column(7).Style.Fill.SetBackground(eThemeSchemeColor.Accent1);
            ws.Row(6).Style.Fill.SetBackground(ExcelIndexedColor.Indexed4);

            ws.Cells["C3,F3"].Style.Fill.SetBackground(Color.Red);
            ws.Cells["F3"].Style.Fill.SetBackground(Color.Red);
            ws.Cells["C2"].Value = 2;
            ws.Cells["A3"].Value = "A3";

            Assert.AreEqual(7, ws.Cells["C2"].Style.Fill.BackgroundColor.Indexed);
            Assert.AreEqual(eThemeSchemeColor.Accent1, ws.Cells["G2"].Style.Fill.BackgroundColor.Theme);

            Assert.AreEqual(5, ws.Cells["A3"].Style.Fill.BackgroundColor.Indexed);
            Assert.AreEqual(Color.Red.ToArgb().ToString("X"), ws.Cells["C3"].Style.Fill.BackgroundColor.Rgb);
            Assert.AreEqual(Color.Red.ToArgb().ToString("X"), ws.Cells["F3"].Style.Fill.BackgroundColor.Rgb);
            Assert.AreEqual(eThemeSchemeColor.Accent1, ws.Cells["G3"].Style.Fill.BackgroundColor.Theme);
            Assert.AreEqual(5, ws.Cells["H3"].Style.Fill.BackgroundColor.Indexed);

            Assert.AreEqual(4, ws.Cells["A6"].Style.Fill.BackgroundColor.Indexed);
            Assert.AreEqual(4, ws.Cells["F6"].Style.Fill.BackgroundColor.Indexed);
            Assert.AreEqual(4, ws.Cells["G6"].Style.Fill.BackgroundColor.Indexed);

            Assert.AreEqual(eThemeSchemeColor.Accent1, ws.Cells["G7"].Style.Fill.BackgroundColor.Theme);

            Assert.AreEqual(7, ws.Cells["C102"].Style.Fill.BackgroundColor.Indexed);

            _pck.Save();
        }
        [TestMethod]
        public void TextRotation255()
        {
            var ws = _pck.Workbook.Worksheets.Add("TextRotation");

            ws.Cells["A1:A182"].Value="RotatedText";
            for(int i=1;i<=180;i++)
            {
                ws.Cells[i,1].Style.TextRotation = i;
            }
            ws.Cells[181, 1].Style.TextRotation = 255;
            ws.Cells[182, 1].Style.SetTextVertical();

            Assert.AreEqual(255, ws.Cells[181, 1].Style.TextRotation);
            Assert.AreEqual(255, ws.Cells[182, 1].Style.TextRotation);
        }
        [TestMethod]
        public void ValidateGradient()
        {
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("List1");
                var gradient = ws.Cells[1, 1].Style.Fill.Gradient;

                //Validate uninititialized values.
                Assert.AreEqual(ExcelFillGradientType.None, gradient.Type);
                Assert.AreEqual(0, gradient.Degree);
                Assert.AreEqual(0, gradient.Top);
                Assert.AreEqual(0, gradient.Bottom);
                Assert.AreEqual(0, gradient.Right);
                Assert.AreEqual(0, gradient.Left);

                Assert.IsNull(gradient.Color1.Rgb);
                Assert.IsNull(gradient.Color1.Theme);
                Assert.AreEqual(-1,gradient.Color1.Indexed);
                Assert.IsFalse(gradient.Color1.Auto);
                
                //Validate Inititialized values.
                gradient.Type = ExcelFillGradientType.Linear;
                Assert.AreEqual(ExcelFillGradientType.Linear, gradient.Type);
                gradient.Top = 0.1;
                Assert.AreEqual(0.1, gradient.Top);
                gradient.Bottom = 0.2;
                Assert.AreEqual(0.2, gradient.Bottom);
                gradient.Right = 0.3;
                Assert.AreEqual(0.3, gradient.Right);
                gradient.Left = 0.4;
                Assert.AreEqual(0.4, gradient.Left);

                gradient.Color2.SetColor(Color.Red);
                Assert.AreEqual("FFFF0000", gradient.Color2.Rgb);
                gradient.Color2.Theme = eThemeSchemeColor.Accent1;
                Assert.AreEqual(eThemeSchemeColor.Accent1, gradient.Color2.Theme);
                gradient.Color2.SetColor(ExcelIndexedColor.Indexed62);
                Assert.AreEqual(62, gradient.Color2.Indexed);
                gradient.Color2.SetAuto();
                Assert.IsTrue(gradient.Color2.Auto);
            }
        }
    }
}

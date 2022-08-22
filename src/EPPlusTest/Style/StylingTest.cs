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
using OfficeOpenXml.SystemDrawing.Text;
using System.Drawing;
using System.Globalization;
using System.Threading;

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
        [TestMethod]
        public void ValidateFontCharsetCondenseExtendAndShadow()
        {
            var ws = _pck.Workbook.Worksheets.Add("Font");
            ws.Cells["A1:C3"].Value = "Font";

            Assert.IsNull(ws.Cells["A1"].Style.Font.Charset);

            ws.Cells["A1"].Style.Font.Charset=2;

            Assert.AreEqual(2, ws.Cells["A1"].Style.Font.Charset);
        }
        [TestMethod]
        public void NormalStyleIssue()
        {
            using (var p = OpenPackage("NormalShouldReflectToEmptyCells.xlsx", true))
            {
                ExcelStyle normal = p.Workbook.Styles.NamedStyles[0].Style;
                normal.Font.Name = "Calibri";
                normal.Font.Size = 10;
                normal.Fill.PatternType = ExcelFillStyle.Solid;
                normal.Fill.BackgroundColor.SetColor(Color.LightGray);
                p.Workbook.Styles.NamedStyles[0].CustomBuildin = true;

                ExcelWorksheet ws = p.Workbook.Worksheets.Add("test");
                Assert.AreEqual("Calibri", normal.Font.Name);
                Assert.AreEqual(10, normal.Font.Size);
                //p.Workbook.Styles.UpdateXml();
                Assert.AreEqual("Calibri", normal.Font.Name);
                Assert.AreEqual(10, normal.Font.Size);
                ws.DefaultRowHeight = 12.75;
                ws.SetValue(1, 1, "test");
                Assert.AreEqual(10, ws.Cells["A1"].Style.Font.Size);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void ChangingTheNormalStyleFontWithAutofitColumns()
        {
            using (var p = new ExcelPackage())
            {
                var CustomFont = new Font("Calibri", 11);
                p.Settings.TextSettings.PrimaryTextMeasurer = new SystemDrawingTextMeasurer();
                p.Workbook.ThemeManager.CreateDefaultTheme();
                var defaultTheme = p.Workbook.ThemeManager.CurrentTheme;
                defaultTheme.FontScheme.MajorFont.SetLatinFont(CustomFont.Name);
                defaultTheme.FontScheme.MinorFont.SetLatinFont(CustomFont.Name);
                ExcelStyle normal = p.Workbook.Styles.NamedStyles[0].Style;
                normal.Font.Name = CustomFont.Name;
                normal.Font.Size = CustomFont.Size;
                ExcelWorkbook workbook = p.Workbook;
                ExcelWorksheet ws = p.Workbook.Worksheets.Add("sheet");
                ExcelStyle style = workbook.Styles.CreateNamedStyle("style").Style;
                //ws.SetValue(1, 1, "very long text very long text very long text");
                ws.Cells["A1"].Value = "番番番番(番番番)番番番番(番番番)番番番番(番番番)番番番番(番番番)番番番番(番番番)";

                ws.Cells[1, 1].StyleName = "style";
                Assert.AreEqual(11, ws.Cells[1, 1].Style.Font.Size);
                ws.Cells.AutoFitColumns(1);
                SaveWorkbook("AutoFitColumnWithStyle.xlsx", p);
            }
        }
        [TestMethod]
        public void SetThemeFontIssue()
        {
            using (var p = OpenPackage("DefaultFont.xlsx", true))
            {
                var DefaultFont = new Font("Corbel", 10);
                p.Workbook.ThemeManager.CreateDefaultTheme();
                var defaultTheme = p.Workbook.ThemeManager.CurrentTheme;
                defaultTheme.FontScheme.MajorFont.SetLatinFont(DefaultFont.Name);
                defaultTheme.FontScheme.MinorFont.SetLatinFont(DefaultFont.Name);
                ExcelStyle normal = p.Workbook.Styles.NamedStyles[0].Style;
                normal.Font.Name = DefaultFont.Name;
                normal.Font.Size = DefaultFont.Size;
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells[1, 1].Value = 1000;

                Assert.AreEqual("Corbel", ws.Cells[1, 1].Style.Font.Name);
                Assert.AreEqual(10, ws.Cells[1, 1].Style.Font.Size);

                ws.Cells[1, 1].Style.Numberformat.Format = "#,##0";
                ws.Cells[1, 1].Style.Border.BorderAround(ExcelBorderStyle.Hair);

                Assert.AreEqual("Corbel", ws.Cells[1, 1].Style.Font.Name);
                Assert.AreEqual(10, ws.Cells[1, 1].Style.Font.Size);

                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void VerifyDateText()
        {
            var ci = CultureInfo.CurrentCulture;
            CultureInfo.CurrentCulture = new CultureInfo("en-US");
            try
            {
                using (var p = new ExcelPackage())
                {
                    var ws = p.Workbook.Worksheets.Add("Sheet1");
                    ws.Cells["A1"].Value = 0;
                    ws.Cells["A2"].Value = 1;
                    ws.Cells["A3"].Value = -1;
                    ws.Cells["A1:A3"].Style.Numberformat.Format = "h:mm:ss tt";
                    Assert.AreEqual("0:00:00 AM", ws.Cells["A1"].Text);
                    Assert.AreEqual("0:00:00 AM", ws.Cells["A2"].Text);
                    Assert.IsNull(ws.Cells["A3"].Text); //Invalid value -1, replace with #####
                }
            }
            finally
            {
                CultureInfo.CurrentCulture = ci;
            }
        }
        [TestMethod]
        public void VerifyStyleXfsCount()
        {
            using (var p = new ExcelPackage())
            {
                var wb = p.Workbook;
                var ws = wb.Worksheets.Add("Sheet1");

                ws.Cells["A1:A3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells["A1:A3"].Style.Fill.BackgroundColor.SetColor(Color.Red);
                ws.Cells["A1:A5"].Style.Border.BorderAround(ExcelBorderStyle.Dotted);
                ws.Cells["B1:B5"].Style.Font.Name = "Arial";
                wb.Styles.UpdateXml();
                var count = wb.StylesXml.SelectSingleNode("//d:styleSheet/d:cellXfs/@count", wb.NameSpaceManager).Value;
                Assert.AreEqual("6", count);
            }
        }
    }
}



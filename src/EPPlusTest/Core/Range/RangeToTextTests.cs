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
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace EPPlusTest.Core.Range
{
    [TestClass]
    public class RangeToTextTests : TestBase
    {
        static ExcelPackage _pck;
        static ExcelWorksheet _ws;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("Range_ToText.xlsx", true);
            _ws = _pck.Workbook.Worksheets.Add("ToTextData");
            var noItems = 100;
            LoadTestdata(_ws, noItems);
            SetDateValues(_ws, noItems);

            _ws.SetValue("C6", "\"" + _ws.GetValue<string>(6, 3) + "\"");
        }
        #region ToText
        [TestMethod]
        public void ToTextDefault()
        {
            //Setup
            var fmt = new ExcelOutputTextFormat();
            var text = _ws.Cells["A1:D5"].ToText(fmt);
            var lines = text.Split(new string[] { fmt.EOL }, StringSplitOptions.None);
            var cols = lines[1].Split(fmt.Delimiter);
            
            //Assert
            Assert.AreEqual(_ws.Cells["A2"].Text, cols[0]);
            Assert.AreEqual(_ws.Cells["B2"].Text, cols[1]);
            Assert.AreEqual(_ws.Cells["C2"].Text, cols[2]);
            Assert.AreEqual("66.00", cols[3]);
        }
        [TestMethod]
        public void ToTextNoCellFormat()
        {
            //Setup
            var fmt = new ExcelOutputTextFormat()
            {
                UseCellFormat = false
            };
            var text = _ws.Cells["A1:D5"].ToText(fmt);
            var lines = text.Split(new string[] { fmt.EOL }, StringSplitOptions.None);
            var cols = lines[1].Split(fmt.Delimiter);

            //Assert
            Assert.AreEqual(_ws.Cells["A2"].GetValue<DateTime>().ToString("G", CultureInfo.InvariantCulture), cols[0]);
            Assert.AreEqual(_ws.Cells["B2"].Text, cols[1]);
            Assert.AreEqual(_ws.Cells["C2"].Text, cols[2]);
            Assert.AreEqual(_ws.Cells["D2"].GetValue<double>().ToString("r", CultureInfo.InvariantCulture), cols[3]);
        }
        [TestMethod]
        public void ToTextSwedishCulture()
        {
            //Setup
            var culture = new CultureInfo("sv-SE");
            var fmt = new ExcelOutputTextFormat()
            {
                Culture = culture,
            };
            var text = _ws.Cells["A1:D5"].ToText(fmt);
            var lines = text.Split(new string[] { fmt.EOL }, StringSplitOptions.None);
            var cols = lines[1].Split(fmt.Delimiter);

            //Assert
            Assert.AreEqual(_ws.Cells["A2"].GetValue<DateTime>().ToString("yyyy-MM-dd", culture), cols[0]);
            Assert.AreEqual(_ws.Cells["B2"].Text, cols[1]);
            Assert.AreEqual(_ws.Cells["C2"].Text, cols[2]);
            Assert.IsTrue(lines[1].EndsWith(_ws.Cells["D2"].GetValue<double>().ToString("0.00", culture)));
        }
        [TestMethod]
        public void ToTextFormatAndTextQualifier()
        {
            //Setup
            var fmt = new ExcelOutputTextFormat()
            {
                TextQualifier='"',
                Formats=new string[] { "yyyy-MM-dd", null, null, "0.00" },
            };
            var text = _ws.Cells["A1:D5"].ToText(fmt);
            var lines = text.Split(new string[] { fmt.EOL }, StringSplitOptions.None);
            var cols = lines[1].Split(fmt.Delimiter);

            //Assert
            Assert.AreEqual(_ws.Cells["A2"].GetValue<DateTime>().ToString("yyyy-MM-dd"), cols[0]);
            Assert.AreEqual(_ws.Cells["B2"].Text, cols[1]);
            Assert.AreEqual(fmt.TextQualifier + _ws.Cells["C2"].Text + fmt.TextQualifier, cols[2]);
            Assert.AreEqual(_ws.Cells["D2"].GetValue<double>().ToString("0.00", CultureInfo.InvariantCulture), cols[3]);
        }
        [TestMethod]
        public void ToTextTextQualifierDouble()
        {
            //Setup
            var fmt = new ExcelOutputTextFormat()
            {
                TextQualifier = '"',
                Formats = new string[] { "yyyy-MM-dd" }
            };
            var text = _ws.Cells["A1:D6"].ToText(fmt);
            var lines = text.Split(new string[] { fmt.EOL }, StringSplitOptions.None);
            var cols = lines[5].Split(fmt.Delimiter);

            //Assert
            Assert.AreEqual(_ws.Cells["A6"].GetValue<DateTime>().ToString("yyyy-MM-dd"), cols[0]);
            Assert.AreEqual(_ws.Cells["B6"].Text, cols[1]);
            Assert.AreEqual(new string (fmt.TextQualifier,2) + _ws.Cells["C6"].Text + new string(fmt.TextQualifier, 2), cols[2]);
            Assert.AreEqual("198.00", cols[3]);
        }
        [TestMethod]
        public void ToTextDelimiterAndCustomDecimalDelimiter()
        {
            //Setup
            var fmt = new ExcelOutputTextFormat()
            {
                Delimiter=';',
                DecimalSeparator=",",
                Formats=new string[] {null,null,null,"0.00"}
            };
            var text = _ws.Cells["A1:D6"].ToText(fmt);
            var lines = text.Split(new string[] { fmt.EOL }, StringSplitOptions.None);
            var cols = lines[1].Split(fmt.Delimiter);

            //Assert
            Assert.AreEqual(_ws.Cells["A2"].Text, cols[0]);
            Assert.AreEqual(_ws.Cells["B2"].Text, cols[1]);
            Assert.AreEqual(_ws.Cells["C2"].Text, cols[2]);
            Assert.AreEqual("66,00", cols[3]);
        }
        [TestMethod]
        public void ToTextCustomThousandSeparator()
        {
            //Setup
            var fmt = new ExcelOutputTextFormat()
            {
                Delimiter = '|',
                DecimalSeparator = ",",
                ThousandsSeparator = " ",
                Formats = new string[] { null, null, null, "#,##0.00" }
            };
            var text = _ws.Cells["A1:D35"].ToText(fmt);
            var lines = text.Split(new string[] { fmt.EOL }, StringSplitOptions.None);
            var cols = lines[34].Split(fmt.Delimiter);

            //Assert
            Assert.AreEqual(_ws.Cells["A35"].Text, cols[0]);
            Assert.AreEqual(_ws.Cells["B35"].Text, cols[1]);
            Assert.AreEqual(_ws.Cells["C35"].Text, cols[2]);
            Assert.AreEqual("1 155,00", cols[3]);
        }
        [TestMethod]
        public void ToTextFormatTextNoCellFormat()
        {
            //Setup
            var fmt = new ExcelOutputTextFormat()
            {
                TextQualifier = '\'',
                UseCellFormat = false,
                Formats = new string[] { "$", "$", "$", "$" }
            };
            var text = _ws.Cells["A1:D5"].ToText(fmt);
            var lines = text.Split(new string[] { fmt.EOL }, StringSplitOptions.None);
            var cols = lines[1].Split(fmt.Delimiter);

            //Assert
            Assert.AreEqual(fmt.TextQualifier.ToString() + _ws.Cells["A2"].Value.ToString() + fmt.TextQualifier.ToString(), cols[0]);
            Assert.AreEqual(fmt.TextQualifier.ToString() + _ws.Cells["B2"].Value.ToString() + fmt.TextQualifier.ToString(), cols[1]);
            Assert.AreEqual(fmt.TextQualifier.ToString() + _ws.Cells["C2"].Value.ToString() + fmt.TextQualifier.ToString(), cols[2]);
            Assert.AreEqual(fmt.TextQualifier.ToString() + _ws.Cells["D2"].Value.ToString() + fmt.TextQualifier.ToString(), cols[3]);
        }
        [TestMethod]
        public void ToTextFormatTextAndCellFormat()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            //Setup
            var fmt = new ExcelOutputTextFormat()
            {
                Delimiter='.',
                TextQualifier = '\'',
                UseCellFormat=true,
                Culture=new CultureInfo("sv-SE"),
                Formats = new string[] { "", "$", "$", null}
            };
            var text = _ws.Cells["A1:D5"].ToText(fmt);
            var lines = text.Split(new string[] { fmt.EOL }, StringSplitOptions.None);
            var cols = lines[1].Split(fmt.Delimiter);

            //Assert
            Assert.AreEqual(_ws.Cells["A2"].Text, cols[0]);
            Assert.AreEqual(fmt.TextQualifier.ToString() + _ws.Cells["B2"].Value.ToString() + fmt.TextQualifier.ToString(), cols[1]);
            Assert.AreEqual(fmt.TextQualifier.ToString() + _ws.Cells["C2"].Value.ToString() + fmt.TextQualifier.ToString(), cols[2]);
            Assert.AreEqual("66,00", cols[3]);

            Thread.CurrentThread.CurrentCulture = ci;
        }
        [TestMethod]
        public void ToTextSkipLines()
        {
            //Setup
            var fmt = new ExcelOutputTextFormat()
            {
                SkipLinesBeginning =1,
                SkipLinesEnd=1
            };

            var text = _ws.Cells["A1:D5"].ToText(fmt);
            var lines = text.Split(new string[] { fmt.EOL }, StringSplitOptions.None);
            var colsHeaders = lines[0].Split(fmt.Delimiter);
            var cols = lines[1].Split(fmt.Delimiter);

            //Assert
            Assert.AreEqual(3, lines.Length);

            Assert.AreEqual(_ws.Cells["A1"].Text, colsHeaders[0]);
            Assert.AreEqual(_ws.Cells["B1"].Text, colsHeaders[1]);
            Assert.AreEqual(_ws.Cells["C1"].Text, colsHeaders[2]);
            Assert.AreEqual(_ws.Cells["D1"].Text, colsHeaders[3]);

            Assert.AreEqual(_ws.Cells["A3"].Text, cols[0]);
            Assert.AreEqual(_ws.Cells["B3"].Text, cols[1]);
            Assert.AreEqual(_ws.Cells["C3"].Text, cols[2]);
            Assert.AreEqual("99.00", cols[3]);
        }
        [TestMethod]
        public void ToTextHeaderAndFooter()
        {
            //Setup
            var fmt = new ExcelOutputTextFormat()
            {
                Header = "Starts With",
                Footer = "Ends With"
            };
            var text = _ws.Cells["A1:D5"].ToText(fmt);

            //Assert
            Assert.IsTrue(text.StartsWith(fmt.Header + fmt.EOL));
            Assert.IsTrue(text.EndsWith(fmt.EOL + fmt.Footer));
        }
        #endregion
        #region ToTextAsync
        [TestMethod]
        public async Task ToTextDefaultAsync()
        {
            //Setup
            var fmt = new ExcelOutputTextFormat();
            var text = await _ws.Cells["A1:D5"].ToTextAsync(fmt).ConfigureAwait(false);
            var lines = text.Split(new string[] { fmt.EOL }, StringSplitOptions.None);
            var cols = lines[1].Split(fmt.Delimiter);

            //Assert
            Assert.AreEqual(_ws.Cells["A2"].Text, cols[0]);
            Assert.AreEqual(_ws.Cells["B2"].Text, cols[1]);
            Assert.AreEqual(_ws.Cells["C2"].Text, cols[2]);
            Assert.AreEqual("66.00", cols[3]);
        }
        [TestMethod]
        public async Task ToTextNoCellFormatAsync()
        {
            //Setup
            var fmt = new ExcelOutputTextFormat()
            {
                UseCellFormat = false
            };
            var text = await _ws.Cells["A1:D5"].ToTextAsync(fmt).ConfigureAwait(false);
            var lines = text.Split(new string[] { fmt.EOL }, StringSplitOptions.None);
            var cols = lines[1].Split(fmt.Delimiter);

            //Assert
            Assert.AreEqual(_ws.Cells["A2"].GetValue<DateTime>().ToString("G", CultureInfo.InvariantCulture), cols[0]);
            Assert.AreEqual(_ws.Cells["B2"].Text, cols[1]);
            Assert.AreEqual(_ws.Cells["C2"].Text, cols[2]);
            Assert.AreEqual(_ws.Cells["D2"].GetValue<double>().ToString("r", CultureInfo.InvariantCulture), cols[3]);
        }
        [TestMethod]
        public async Task ToTextSwedishCultureAsync()
        {
            //Setup
            var culture = new CultureInfo("sv-SE");
            var fmt = new ExcelOutputTextFormat()
            {
                Culture = culture,
            };
            var text = await _ws.Cells["A1:D5"].ToTextAsync(fmt).ConfigureAwait(false);
            var lines = text.Split(new string[] { fmt.EOL }, StringSplitOptions.None);
            var cols = lines[1].Split(fmt.Delimiter);

            //Assert
            Assert.AreEqual(_ws.Cells["A2"].GetValue<DateTime>().ToString("yyyy-MM-dd", culture), cols[0]);
            Assert.AreEqual(_ws.Cells["B2"].Text, cols[1]);
            Assert.AreEqual(_ws.Cells["C2"].Text, cols[2]);
            Assert.IsTrue(lines[1].EndsWith(_ws.Cells["D2"].GetValue<double>().ToString("0.00", culture)));
        }
        [TestMethod]
        public async Task ToTextFormatAndTextQualifierAsync()
        {
            //Setup
            var fmt = new ExcelOutputTextFormat()
            {
                TextQualifier = '"',
                Formats = new string[] { "yyyy-MM-dd", null, null, "0.00" },
            };
            var text = await _ws.Cells["A1:D5"].ToTextAsync(fmt).ConfigureAwait(false);
            var lines = text.Split(new string[] { fmt.EOL }, StringSplitOptions.None);
            var cols = lines[1].Split(fmt.Delimiter);

            //Assert
            Assert.AreEqual(_ws.Cells["A2"].GetValue<DateTime>().ToString("yyyy-MM-dd"), cols[0]);
            Assert.AreEqual(_ws.Cells["B2"].Text, cols[1]);
            Assert.AreEqual(fmt.TextQualifier + _ws.Cells["C2"].Text + fmt.TextQualifier, cols[2]);
            Assert.AreEqual(_ws.Cells["D2"].GetValue<double>().ToString("0.00", CultureInfo.InvariantCulture), cols[3]);
        }
        [TestMethod]
        public async Task ToTextTextQualifierDoubleAsync()
        {
            //Setup
            var fmt = new ExcelOutputTextFormat()
            {
                TextQualifier = '"',
                Formats = new string[] { "yyyy-MM-dd" }
            };
            var text = await _ws.Cells["A1:D6"].ToTextAsync(fmt).ConfigureAwait(false);
            var lines = text.Split(new string[] { fmt.EOL }, StringSplitOptions.None);
            var cols = lines[5].Split(fmt.Delimiter);

            //Assert
            Assert.AreEqual(_ws.Cells["A6"].GetValue<DateTime>().ToString("yyyy-MM-dd"), cols[0]);
            Assert.AreEqual(_ws.Cells["B6"].Text, cols[1]);
            Assert.AreEqual(new string(fmt.TextQualifier, 2) + _ws.Cells["C6"].Text + new string(fmt.TextQualifier, 2), cols[2]);
            Assert.AreEqual("198.00", cols[3]);
        }
        [TestMethod]
        public async Task ToTextDelimiterAndCustomDecimalDelimiterAsync()
        {
            //Setup
            var fmt = new ExcelOutputTextFormat()
            {
                Delimiter = ';',
                DecimalSeparator = ",",
                Formats = new string[] { null, null, null, "0.00" }
            };
            var text = await _ws.Cells["A1:D6"].ToTextAsync(fmt).ConfigureAwait(false);
            var lines = text.Split(new string[] { fmt.EOL }, StringSplitOptions.None);
            var cols = lines[1].Split(fmt.Delimiter);

            //Assert
            Assert.AreEqual(_ws.Cells["A2"].Text, cols[0]);
            Assert.AreEqual(_ws.Cells["B2"].Text, cols[1]);
            Assert.AreEqual(_ws.Cells["C2"].Text, cols[2]);
            Assert.AreEqual("66,00", cols[3]);
        }
        [TestMethod]
        public async Task ToTextCustomThousandSeparatorAsync()
        {
            //Setup
            var fmt = new ExcelOutputTextFormat()
            {
                Delimiter = '|',
                DecimalSeparator = ",",
                ThousandsSeparator = " ",
                Formats = new string[] { null, null, null, "#,##0.00" }
            };
            var text = await _ws.Cells["A1:D35"].ToTextAsync(fmt).ConfigureAwait(false);
            var lines = text.Split(new string[] { fmt.EOL }, StringSplitOptions.None);
            var cols = lines[34].Split(fmt.Delimiter);

            //Assert
            Assert.AreEqual(_ws.Cells["A35"].Text, cols[0]);
            Assert.AreEqual(_ws.Cells["B35"].Text, cols[1]);
            Assert.AreEqual(_ws.Cells["C35"].Text, cols[2]);
            Assert.AreEqual("1 155,00", cols[3]);
        }
        [TestMethod]
        public async Task ToTextFormatTextNoCellFormatAsync()
        {
            //Setup
            var fmt = new ExcelOutputTextFormat()
            {
                TextQualifier = '\'',
                UseCellFormat = false,
                Formats = new string[] { "$", "$", "$", "$" }
            };
            var text = await _ws.Cells["A1:D5"].ToTextAsync(fmt).ConfigureAwait(false);
            var lines = text.Split(new string[] { fmt.EOL }, StringSplitOptions.None);
            var cols = lines[1].Split(fmt.Delimiter);

            //Assert
            Assert.AreEqual(fmt.TextQualifier.ToString() + _ws.Cells["A2"].Value.ToString() + fmt.TextQualifier.ToString(), cols[0]);
            Assert.AreEqual(fmt.TextQualifier.ToString() + _ws.Cells["B2"].Value.ToString() + fmt.TextQualifier.ToString(), cols[1]);
            Assert.AreEqual(fmt.TextQualifier.ToString() + _ws.Cells["C2"].Value.ToString() + fmt.TextQualifier.ToString(), cols[2]);
            Assert.AreEqual(fmt.TextQualifier.ToString() + _ws.Cells["D2"].Value.ToString() + fmt.TextQualifier.ToString(), cols[3]);
        }
        [TestMethod]
        public async Task ToTextFormatTextAndCellFormatAsync()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            //Setup
            var fmt = new ExcelOutputTextFormat()
            {
                Delimiter = '.',
                TextQualifier = '\'',
                UseCellFormat = true,
                Culture = new CultureInfo("sv-SE"),
                Formats = new string[] { "", "$", "$", null }
            };
            var text = await _ws.Cells["A1:D5"].ToTextAsync(fmt).ConfigureAwait(false);
            var lines = text.Split(new string[] { fmt.EOL }, StringSplitOptions.None);
            var cols = lines[1].Split(fmt.Delimiter);

            //Assert
            Assert.AreEqual(_ws.Cells["A2"].Text, cols[0]);
            Assert.AreEqual(fmt.TextQualifier.ToString() + _ws.Cells["B2"].Value.ToString() + fmt.TextQualifier.ToString(), cols[1]);
            Assert.AreEqual(fmt.TextQualifier.ToString() + _ws.Cells["C2"].Value.ToString() + fmt.TextQualifier.ToString(), cols[2]);
            Assert.AreEqual("66,00", cols[3]);
            Thread.CurrentThread.CurrentCulture = ci;
        }
        [TestMethod]
        public async Task ToTextSkipLinesAsync()
        {
            //Setup
            var fmt = new ExcelOutputTextFormat()
            {
                SkipLinesBeginning = 1,
                SkipLinesEnd = 1
            };
            var text = await _ws.Cells["A1:D5"].ToTextAsync(fmt).ConfigureAwait(false);
            var lines = text.Split(new string[] { fmt.EOL }, StringSplitOptions.None);
            var colsHeaders = lines[0].Split(fmt.Delimiter);
            var cols = lines[1].Split(fmt.Delimiter);

            //Assert
            Assert.AreEqual(3, lines.Length);

            Assert.AreEqual(_ws.Cells["A1"].Text, colsHeaders[0]);
            Assert.AreEqual(_ws.Cells["B1"].Text, colsHeaders[1]);
            Assert.AreEqual(_ws.Cells["C1"].Text, colsHeaders[2]);
            Assert.AreEqual(_ws.Cells["D1"].Text, colsHeaders[3]);

            Assert.AreEqual(_ws.Cells["A3"].Text, cols[0]);
            Assert.AreEqual(_ws.Cells["B3"].Text, cols[1]);
            Assert.AreEqual(_ws.Cells["C3"].Text, cols[2]);
            Assert.AreEqual("99.00", cols[3]);
        }
        [TestMethod]
        public async Task ToTextHeaderAndFooterAsync()
        {
            //Setup
            var fmt = new ExcelOutputTextFormat()
            {
                Header = "Starts With",
                Footer = "Ends With"
            };
            var text = await _ws.Cells["A1:D5"].ToTextAsync(fmt).ConfigureAwait(false);

            //Assert
            Assert.IsTrue(text.StartsWith(fmt.Header + fmt.EOL));
            Assert.IsTrue(text.EndsWith(fmt.EOL + fmt.Footer));
        }
        [TestMethod]
        public void ToTextHandleRichTextCells()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("RichText");
                //Setup
                ws.Cells["A1"].RichText.Add("RichText 1");
                var rt = ws.Cells["A2"].RichText;
                rt.Add("Rich");
                var rtPart = rt.Add("Text");
                rtPart.Color = Color.Red;
                rt.Add(" 2");
                var text = ws.Cells["A1:A2"].ToText();

                //Assert
                Assert.AreEqual("RichText 1\r\nRichText 2", text);
                Assert.AreEqual(3, ws.Cells["A2"].RichText.Count);
                Assert.AreEqual(Color.Red.ToArgb(), ws.Cells["A2"].RichText[1].Color.ToArgb());
            }
        }

        #endregion
    }
}

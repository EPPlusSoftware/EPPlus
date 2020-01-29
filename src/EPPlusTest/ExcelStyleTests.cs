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
using System.Xml;

namespace EPPlusTest
{
    [TestClass]
    public class ExcelStyleTests
    {
        [TestMethod]
        public void QuotePrefixStyle()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("QuotePrefixTest");
                var cell = ws.Cells["B2"];
                cell.Style.QuotePrefix = true;
                Assert.IsTrue(cell.Style.QuotePrefix);

                p.Workbook.Styles.UpdateXml();                
                var nodes = p.Workbook.StylesXml.SelectNodes("//d:cellXfs/d:xf", p.Workbook.NameSpaceManager);
                // Since the quotePrefix attribute is not part of the default style,
                // a new one should be created and referenced.
                Assert.AreNotEqual(0, cell.StyleID);
                Assert.IsNull(nodes[0].Attributes["quotePrefix"]);
                Assert.AreEqual("1", nodes[cell.StyleID].Attributes["quotePrefix"].Value);
            }
        }
    }
}

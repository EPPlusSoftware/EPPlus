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
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
namespace OfficeOpenXml.Core.Range
{
    [TestClass]
    public class RangeAddressCopyTests : TestBase
    {
        [TestMethod]
        public void ValidateCopyFormulasRow()
        {
            using (var pck = new ExcelPackage())
            {
                var ws = pck.Workbook.Worksheets.Add("CopyRowWise");
                ws.Cells["A1:C3"].Value = 1;
                ws.Cells["D3"].Formula="A1";
                ws.Cells["E3"].Formula = "B2";
                ws.Cells["F3"].Formula = "C3";
                ws.Cells["G3"].Formula = "A$1";
                ws.Cells["H3"].Formula = "B$2";
                ws.Cells["J3"].Formula = "C$3";

                //Validate that formulas are copied correctly row-wise
                ws.Cells["D3"].Copy(ws.Cells["D2"]);
                Assert.AreEqual("#REF!", ws.Cells["D2"].Formula);
                ws.Cells["E3"].Copy(ws.Cells["E2"]);
                Assert.AreEqual("B1", ws.Cells["E2"].Formula);
                ws.Cells["F3"].Copy(ws.Cells["F2"]);
                Assert.AreEqual("C2", ws.Cells["F2"].Formula);
                ws.Cells["G3"].Copy(ws.Cells["G2"]);
                Assert.AreEqual("A$1", ws.Cells["G2"].Formula);
                ws.Cells["H3"].Copy(ws.Cells["H2"]);
                Assert.AreEqual("B$2", ws.Cells["H2"].Formula);
                ws.Cells["J3"].Copy(ws.Cells["J2"]);
                Assert.AreEqual("C$3", ws.Cells["J2"].Formula);

                ws.Cells["D3"].Copy(ws.Cells["D1"]);
                Assert.AreEqual("#REF!", ws.Cells["D1"].Formula);
                ws.Cells["E3"].Copy(ws.Cells["E1"]);
                Assert.AreEqual("#REF!", ws.Cells["E1"].Formula);
                ws.Cells["F3"].Copy(ws.Cells["F1"]);
                Assert.AreEqual("C1", ws.Cells["F1"].Formula);
                ws.Cells["G3"].Copy(ws.Cells["G1"]);
                Assert.AreEqual("A$1", ws.Cells["G1"].Formula);
                ws.Cells["H3"].Copy(ws.Cells["H1"]);
                Assert.AreEqual("B$2", ws.Cells["H1"].Formula);
                ws.Cells["J3"].Copy(ws.Cells["J1"]);
                Assert.AreEqual("C$3", ws.Cells["J1"].Formula);
            }
        }
        [TestMethod]
        public void ValidateCopyFormulasMultiCellRow()
        {
            using (var pck = new ExcelPackage())
            {
                var ws1 = pck.Workbook.Worksheets.Add("Sheet1");
                var ws2 = pck.Workbook.Worksheets.Add("Sheet2");

                //Validate that formulas are copied correctly row-wise
                ws1.Cells["D3"].Formula = "SUM(A1:B1)";

                ws1.Cells["D3"].Copy(ws1.Cells["D2"]);
                Assert.AreEqual("SUM(#REF!)", ws1.Cells["D2"].Formula);
                ws1.Cells["D3"].Copy(ws1.Cells["C3"]);
                Assert.AreEqual("SUM(#REF!)", ws1.Cells["C3"].Formula);
                ws1.Cells["D3"].Copy(ws1.Cells["E3"]);
                Assert.AreEqual("SUM(B1:C1)", ws1.Cells["E3"].Formula);
                ws1.Cells["D3"].Copy(ws1.Cells["D4"]);
                Assert.AreEqual("SUM(A2:B2)", ws1.Cells["D4"].Formula);

                ws1.Cells["D3"].Copy(ws2.Cells["D2"]);
                Assert.AreEqual("SUM(#REF!)", ws2.Cells["D2"].Formula);
                ws1.Cells["D3"].Copy(ws2.Cells["C3"]);
                Assert.AreEqual("SUM(#REF!)", ws2.Cells["C3"].Formula);
                ws1.Cells["D3"].Copy(ws2.Cells["E3"]);
                Assert.AreEqual("SUM(B1:C1)", ws2.Cells["E3"].Formula);
                ws1.Cells["D3"].Copy(ws2.Cells["D4"]);
                Assert.AreEqual("SUM(A2:B2)", ws2.Cells["D4"].Formula);

            }
        }
        [TestMethod]
        public void ValidateCopyFormulasMultiCellFullColumn()
        {
            using (var pck = new ExcelPackage())
            {
                var ws1 = pck.Workbook.Worksheets.Add("Sheet1");
                var ws2 = pck.Workbook.Worksheets.Add("Sheet2");

                //Validate that formulas are copied correctly row-wise
                ws1.Cells["D3"].Formula = "SUM(A:A)";

                ws1.Cells["D3"].Copy(ws1.Cells["D2"]);
                Assert.AreEqual("SUM(A:A)", ws1.Cells["D2"].Formula);
                ws1.Cells["D3"].Copy(ws1.Cells["C3"]);
                Assert.AreEqual("SUM(#REF!)", ws1.Cells["C3"].Formula);
                ws1.Cells["D3"].Copy(ws1.Cells["E3"]);
                Assert.AreEqual("SUM(B:B)", ws1.Cells["E3"].Formula);
                ws1.Cells["D3"].Copy(ws1.Cells["D4"]);
                Assert.AreEqual("SUM(A:A)", ws1.Cells["D4"].Formula);

                ws1.Cells["D3"].Copy(ws2.Cells["D2"]);
                Assert.AreEqual("SUM(A:A)", ws2.Cells["D2"].Formula);
                ws1.Cells["D3"].Copy(ws2.Cells["C3"]);
                Assert.AreEqual("SUM(#REF!)", ws2.Cells["C3"].Formula);
                ws1.Cells["D3"].Copy(ws2.Cells["E3"]);
                Assert.AreEqual("SUM(B:B)", ws2.Cells["E3"].Formula);
                ws1.Cells["D3"].Copy(ws2.Cells["D4"]);
                Assert.AreEqual("SUM(A:A)", ws2.Cells["D4"].Formula);
            }
        }
        [TestMethod]
        public void ValidateCopyFormulasMultiCellFullColumnFixed()
        {
            using (var pck = new ExcelPackage())
            {
                var ws1 = pck.Workbook.Worksheets.Add("Sheet1");
                var ws2 = pck.Workbook.Worksheets.Add("Sheet2");

                //Validate that formulas are copied correctly row-wise
                ws1.Cells["D3"].Formula = "SUM($A:$A)";

                ws1.Cells["D3"].Copy(ws1.Cells["D2"]);
                Assert.AreEqual("SUM($A:$A)", ws1.Cells["D2"].Formula);
                ws1.Cells["D3"].Copy(ws1.Cells["C3"]);
                Assert.AreEqual("SUM($A:$A)", ws1.Cells["C3"].Formula);
                ws1.Cells["D3"].Copy(ws1.Cells["E3"]);
                Assert.AreEqual("SUM($A:$A)", ws1.Cells["E3"].Formula);
                ws1.Cells["D3"].Copy(ws1.Cells["D4"]);
                Assert.AreEqual("SUM($A:$A)", ws1.Cells["D4"].Formula);

                ws1.Cells["D3"].Copy(ws2.Cells["D2"]);
                Assert.AreEqual("SUM($A:$A)", ws2.Cells["D2"].Formula);
                ws1.Cells["D3"].Copy(ws2.Cells["C3"]);
                Assert.AreEqual("SUM($A:$A)", ws2.Cells["C3"].Formula);
                ws1.Cells["D3"].Copy(ws2.Cells["E3"]);
                Assert.AreEqual("SUM($A:$A)", ws2.Cells["E3"].Formula);
                ws1.Cells["D3"].Copy(ws2.Cells["D4"]);
                Assert.AreEqual("SUM($A:$A)", ws2.Cells["D4"].Formula);
            }
        }
        [TestMethod]
        public void ValidateCopyFormulasMultiCellFullRow()
        {
            using (var pck = new ExcelPackage())
            {
                var ws1 = pck.Workbook.Worksheets.Add("Sheet1");
                var ws2 = pck.Workbook.Worksheets.Add("Sheet2");

                //Validate that formulas are copied correctly row-wise
                ws1.Cells["D3"].Formula = "SUM(1:1)";

                ws1.Cells["D3"].Copy(ws1.Cells["D2"]);
                Assert.AreEqual("SUM(#REF!)", ws1.Cells["D2"].Formula);
                ws1.Cells["D3"].Copy(ws1.Cells["C3"]);
                Assert.AreEqual("SUM(1:1)", ws1.Cells["C3"].Formula);
                ws1.Cells["D3"].Copy(ws1.Cells["E3"]);
                Assert.AreEqual("SUM(1:1)", ws1.Cells["E3"].Formula);
                ws1.Cells["D3"].Copy(ws1.Cells["D4"]);
                Assert.AreEqual("SUM(2:2)", ws1.Cells["D4"].Formula);

                ws1.Cells["D3"].Copy(ws2.Cells["D2"]);
                Assert.AreEqual("SUM(#REF!)", ws2.Cells["D2"].Formula);
                ws1.Cells["D3"].Copy(ws2.Cells["C3"]);
                Assert.AreEqual("SUM(1:1)", ws2.Cells["C3"].Formula);
                ws1.Cells["D3"].Copy(ws2.Cells["E3"]);
                Assert.AreEqual("SUM(1:1)", ws2.Cells["E3"].Formula);
                ws1.Cells["D3"].Copy(ws2.Cells["D4"]);
                Assert.AreEqual("SUM(2:2)", ws2.Cells["D4"].Formula);
            }
        }
        [TestMethod]
        public void ValidateCopyFormulasMultiCellFullRowFixed()
        {
            using (var pck = new ExcelPackage())
            {
                var ws1 = pck.Workbook.Worksheets.Add("Sheet1");
                var ws2 = pck.Workbook.Worksheets.Add("Sheet2");

                //Validate that formulas are copied correctly row-wise
                ws1.Cells["D3"].Formula = "SUM($1:$1)";

                ws1.Cells["D3"].Copy(ws1.Cells["D2"]);
                Assert.AreEqual("SUM($1:$1)", ws1.Cells["D2"].Formula);
                ws1.Cells["D3"].Copy(ws1.Cells["C3"]);
                Assert.AreEqual("SUM($1:$1)", ws1.Cells["C3"].Formula);
                ws1.Cells["D3"].Copy(ws1.Cells["E3"]);
                Assert.AreEqual("SUM($1:$1)", ws1.Cells["E3"].Formula);
                ws1.Cells["D3"].Copy(ws1.Cells["D4"]);
                Assert.AreEqual("SUM($1:$1)", ws1.Cells["D4"].Formula);

                ws1.Cells["D3"].Copy(ws2.Cells["D2"]);
                Assert.AreEqual("SUM($1:$1)", ws2.Cells["D2"].Formula);
                ws1.Cells["D3"].Copy(ws2.Cells["C3"]);
                Assert.AreEqual("SUM($1:$1)", ws2.Cells["C3"].Formula);
                ws1.Cells["D3"].Copy(ws2.Cells["E3"]);
                Assert.AreEqual("SUM($1:$1)", ws2.Cells["E3"].Formula);
                ws1.Cells["D3"].Copy(ws2.Cells["D4"]);
                Assert.AreEqual("SUM($1:$1)", ws2.Cells["D4"].Formula);
            }
        }
        [TestMethod]
        public void Copy_Formula_From_Other_Workbook_Issue_Test()
        {
            using (var package = new ExcelPackage())
            {
                var workbook = package.Workbook;
                var sourceWs = workbook.Worksheets.Add("Sheet2");
                sourceWs.Cells["A1"].Value = 24;
                sourceWs.Cells["A2"].Value = 75;
                sourceWs.Cells["A3"].Value = 94;
                sourceWs.Cells["A4"].Value = 34;

                var ws = workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Formula = "VLOOKUP($B$1,Sheet2!A:B,2,FALSE)";

                //Literal formula copy to cell A2 - PASSES
                ws.Cells["A1"].Copy(ws.Cells[2, 1], ExcelRangeCopyOptionFlags.ExcludeFormulas);
                ws.Cells["A2"].Formula = ws.Cells[1, 1].Formula;

                Assert.IsFalse(
                    string.IsNullOrWhiteSpace(ws.Cells["A2"].Formula)
                    , "A2 formula should be set"
                );

                Assert.AreEqual(
                    ws.Cells["A1"].Formula
                    , ws.Cells["A2"].Formula
                    , $"{ws.Cells["A2"].Formula} != {ws.Cells["A1"].Formula}"
                );

                //Cell copy to cell A3 - FAILS
                ws.Cells["A1"].Copy(ws.Cells["A3"]);

                Assert.IsFalse(
                    string.IsNullOrWhiteSpace(ws.Cells["A3"].Formula)
                    , "A3 formula should be set"
                );

                Assert.AreEqual("VLOOKUP($B$1,'SHEET2'!A:B,2,FALSE)", ws.Cells["A3"].Formula);

            }        
        }
    }
}

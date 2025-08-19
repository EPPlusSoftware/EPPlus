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
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class FormulaR1C1Tests
    {
        private ExcelPackage _pck;
        private ExcelWorksheet _sheet;
        [TestInitialize]
        public void Initialize()
        {
            _pck = new ExcelPackage();
            _sheet = _pck.Workbook.Worksheets.Add("R1C1");
        }

        [TestCleanup]
        public void Cleanup()
        {
            _pck.Dispose();
        }

        [TestMethod]
        public void RC2()
        {
            string fR1C1 = "RC2";
            _sheet.Cells[5, 1].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 1].Formula;
            Assert.AreEqual("$B5", f);
            _sheet.Cells[5, 1].Formula = f;
            Assert.AreEqual(fR1C1, _sheet.Cells[5,1].FormulaR1C1);
        }
        [TestMethod]
        public void C()
        {
            string fR1C1 = "SUMIFS(C,C2,RC1)";
            _sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 3].Formula;
            Assert.AreEqual("SUMIFS(C:C,$B:$B,$A5)", f);
            _sheet.Cells[5, 3].Formula = f;
            Assert.AreEqual(fR1C1, _sheet.Cells[5, 3].FormulaR1C1);
        }
        [TestMethod]
        public void C2Abs()
        {
            string fR1C1 = "SUM(C2)";
            _sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 3].Formula;
            Assert.AreEqual("SUM($B:$B)", f);
        }
        [TestMethod]
        public void C2AbsWithSheet()
        {
            string fR1C1 = "SUM(A!C2)";
            _sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 3].Formula;
            Assert.AreEqual("SUM(A!$B:$B)", f);
        }
        [TestMethod]
        public void C2()
        {
            string fR1C1 = "SUM(C2)";
            _sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 3].Formula;
            Assert.AreEqual("SUM($B:$B)", f);
            _sheet.Cells[5, 3].Formula = f;
            Assert.AreEqual(fR1C1, _sheet.Cells[5, 3].FormulaR1C1);
        }
        [TestMethod]
        public void R2Abs()
        {
            string fR1C1 = "SUM(R2)";
            _sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 3].Formula;
            Assert.AreEqual("SUM($2:$2)",f);

            fR1C1 = "SUM(TEST2!R2)";
            _sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
            f = _sheet.Cells[5, 3].Formula;
            Assert.AreEqual("SUM(TEST2!$2:$2)", f);

        }
        [TestMethod]
        public void R2()
        {
            string fR1C1 = "SUM(R2)";
            _sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 3].Formula;
            Assert.AreEqual("SUM($2:$2)", f);
            _sheet.Cells[5, 3].Formula = f;
            Assert.AreEqual(fR1C1, _sheet.Cells[5, 3].FormulaR1C1);
        }
        [TestMethod]
        public void RCRelativeToAB()
        {
            string fR1C1 = "SUMIFS(C,C2,RC1)";
            _sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 3].Formula;
            Assert.AreEqual("SUMIFS(C:C,$B:$B,$A5)", f);
        }
        [TestMethod]
        public void RRelativeToAB()
        {
            string fR1C1 = "SUMIFS(R,C2,RC1)";
            _sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 3].Formula;
            Assert.AreEqual("SUMIFS(5:5,$B:$B,$A5)", f);
        }
        [TestMethod]
        public void RCRelativeToABToR1C1()
        {
            string fR1C1 = "SUMIFS(C,C2,RC1)";
            _sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 3].Formula;
            Assert.AreEqual("SUMIFS(C:C,$B:$B,$A5)", f);
            _sheet.Cells[5, 3].Formula = f;
            Assert.AreEqual(fR1C1, _sheet.Cells[5, 3].FormulaR1C1);
        }
        [TestMethod]
        public void RCRelativeToABToR1C1_2()
        {
            string fR1C1 = "SUM(RC9:RC[-1])";
            _sheet.Cells[5, 13].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 13].Formula;
            Assert.AreEqual("SUM($I5:L5)", f);
            _sheet.Cells[5, 13].Formula = f;
            Assert.AreEqual(fR1C1, _sheet.Cells[5, 13].FormulaR1C1);

            //"RC{colShort} - SUM(RC21:RC12)";
        }
        [TestMethod]
        public void RCFixToABToR1C1_2()
        {
            string fR1C1 = "RC28-SUM(RC12:RC21)";
            _sheet.Cells[6, 13].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[6, 13].Formula;
            Assert.AreEqual("$AB6-SUM($L6:$U6)", f);
            _sheet.Cells[6, 13].Formula = f;
            Assert.AreEqual(fR1C1, _sheet.Cells[6, 13].FormulaR1C1);
        }
        [TestMethod]
        public void SimpleRelativeR1C1()
        {
            string fR1C1 = "R[-1]C[-5]";
            var c = _sheet.Cells[7, 7];
            c.FormulaR1C1 = fR1C1;
            string f = c.Formula;
            Assert.AreEqual("B6", f);
            c.Formula = f;
            Assert.AreEqual(fR1C1, c.FormulaR1C1);
        }
        [TestMethod]
        public void SimpleAbsR1C1()
        {
            string fR1C1 = "R1C5";
            var c = _sheet.Cells[8, 8];
            c.FormulaR1C1 = fR1C1;
            string f = c.Formula;
            Assert.AreEqual("$E$1", f);
            c.Formula = f;
            Assert.AreEqual(fR1C1, c.FormulaR1C1);
        }
        [TestMethod]
        public void FullTwoColumn()
        {
            string formula = "VLOOKUP(C2,A:B,1,0)";
            var c = _sheet.Cells["D2"];
            c.Formula = formula;
            Assert.AreEqual(c.FormulaR1C1, "VLOOKUP(RC[-1],C[-3]:C[-2],1,0)");
            c.FormulaR1C1 = c.FormulaR1C1;
            Assert.AreEqual(c.Formula, formula);
        }
        [TestMethod]
        public void FullColumn()
        {
            string formula = "VLOOKUP(C2,A:A,1,0)";
            var c = _sheet.Cells["D2"];
            c.Formula = formula;
            Assert.AreEqual(c.FormulaR1C1, "VLOOKUP(RC[-1],C[-3],1,0)");
            c.FormulaR1C1 = c.FormulaR1C1;
            Assert.AreEqual(c.Formula, formula);
        }
        [TestMethod]
        public void FullTwoRow()
        {
            string formula = "VLOOKUP(C3,1:2,1,0)";
            var c = _sheet.Cells["D3"];
            c.Formula = formula;
            Assert.AreEqual(c.FormulaR1C1, "VLOOKUP(RC[-1],R[-2]:R[-1],1,0)");
            c.FormulaR1C1 = c.FormulaR1C1;
            Assert.AreEqual(c.Formula, formula);
        }
        [TestMethod]
        public void FullRow()
        {
            string formula = "VLOOKUP(C3,1:1,1,0)";
            var c = _sheet.Cells["D3"];
            c.Formula = formula;
            Assert.AreEqual(c.FormulaR1C1, "VLOOKUP(RC[-1],R[-2],1,0)");
            c.FormulaR1C1 = c.FormulaR1C1;
            Assert.AreEqual(c.Formula, formula);
        }

        [TestMethod]
        public void OutOfRangeCol()
        {
            _sheet.Cells["a3"].FormulaR1C1 = "R[-3]C";
            Assert.AreEqual("#REF!", _sheet.Cells["a3"].Formula);
            _sheet.Cells["a3"].FormulaR1C1 = "R[-2]C";
            Assert.AreEqual("A1", _sheet.Cells["a3"].Formula);

            _sheet.Cells["B3"].FormulaR1C1 = "RC[-2]";
            Assert.AreEqual("#REF!", _sheet.Cells["B3"].Formula);
            _sheet.Cells["B3"].FormulaR1C1 = "RC[-1]";
            Assert.AreEqual("A3", _sheet.Cells["B3"].Formula);

        }
    }
}

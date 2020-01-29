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
using System.Text.RegularExpressions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest
{
    [TestClass]
    public class ExcelCellBaseTest
    {
        #region UpdateFormulaReferences Tests
        [TestMethod]
        public void UpdateFormulaReferencesOnTheSameSheet()
        {
            var result = ExcelCellBase.UpdateFormulaReferences("C3", 3, 3, 2, 2, "sheet", "sheet");
            Assert.AreEqual("F6", result);
        }

        [TestMethod]
        public void UpdateFormulaReferencesIgnoresIncorrectSheet()
        {
            var result = ExcelCellBase.UpdateFormulaReferences("C3", 3, 3, 2, 2, "sheet", "other sheet");
            Assert.AreEqual("C3", result);
        }

        [TestMethod]
        public void UpdateFormulaReferencesFullyQualifiedReferenceOnTheSameSheet()
        {
            var result = ExcelCellBase.UpdateFormulaReferences("'sheet name here'!C3", 3, 3, 2, 2, "sheet name here", "sheet name here");
            Assert.AreEqual("'sheet name here'!F6", result);
        }

        [TestMethod]
        public void UpdateFormulaReferencesFullyQualifiedCrossSheetReferenceArray()
        {
            var result = ExcelCellBase.UpdateFormulaReferences("SUM('sheet name here'!B2:D4)", 3, 3, 3, 3, "cross sheet", "sheet name here");
            Assert.AreEqual("SUM('sheet name here'!B2:G7)", result);
        }

        [TestMethod]
        public void UpdateFormulaReferencesFullyQualifiedReferenceOnADifferentSheet()
        {
            var result = ExcelCellBase.UpdateFormulaReferences("'updated sheet'!C3", 3, 3, 2, 2, "boring sheet", "updated sheet");
            Assert.AreEqual("'updated sheet'!F6", result);
        }

        [TestMethod]
        public void UpdateFormulaReferencesReferencingADifferentSheetIsNotUpdated()
        {
            var result = ExcelCellBase.UpdateFormulaReferences("'boring sheet'!C3", 3, 3, 2, 2, "boring sheet", "updated sheet");
            Assert.AreEqual("'boring sheet'!C3", result);
        }
        #endregion

        #region UpdateCrossSheetReferenceNames Tests
        [TestMethod]
        public void UpdateFormulaSheetReferences()
        {
          var result = ExcelCellBase.UpdateSheetNameInFormula("5+'OldSheet'!$G3+'Some Other Sheet'!C3+SUM(1,2,3)", "OldSheet", "NewSheet");
          Assert.AreEqual("5+'NewSheet'!$G3+'Some Other Sheet'!C3+SUM(1,2,3)", result);
        }
        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void UpdateFormulaSheetReferencesNullOldSheetThrowsException()
        {
            ExcelCellBase.UpdateSheetNameInFormula("formula", null, "sheet2");
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void UpdateFormulaSheetReferencesEmptyOldSheetThrowsException()
        {
            ExcelCellBase.UpdateSheetNameInFormula("formula", string.Empty, "sheet2");
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void UpdateFormulaSheetReferencesNullNewSheetThrowsException()
        {
            ExcelCellBase.UpdateSheetNameInFormula("formula", "sheet1", null);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void UpdateFormulaSheetReferencesEmptyNewSheetThrowsException()
        {
            ExcelCellBase.UpdateSheetNameInFormula("formula", "sheet1", string.Empty);
        }
        #endregion
    }
}

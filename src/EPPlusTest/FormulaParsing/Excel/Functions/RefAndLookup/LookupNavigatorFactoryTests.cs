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
using System.IO;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class LookupNavigatorFactoryTests
    {
        private ExcelPackage _excelPackage;
        private ParsingContext _context;

        [TestInitialize]
        public void Initialize()
        {
            _excelPackage = new ExcelPackage(new MemoryStream());
            _excelPackage.Workbook.Worksheets.Add("Test");
            _context = ParsingContext.Create();
            _context.ExcelDataProvider = new EpplusExcelDataProvider(_excelPackage);
            _context.Scopes.NewScope(FormulaRangeAddress.Empty);
        }

        [TestCleanup]
        public void Cleanup()
        {
            _excelPackage.Dispose();
        }

        [TestMethod]
        public void Should_Return_ExcelLookupNavigator_When_Range_Is_Set()
        {
            var args = new LookupArguments(FunctionsHelper.CreateArgs(8, "A:B", 1), ParsingContext.Create(_excelPackage));
            var navigator = LookupNavigatorFactory.Create(LookupDirection.Horizontal, args, _context);
            Assert.IsInstanceOfType(navigator, typeof(ExcelLookupNavigator));
        }

        [TestMethod]
        public void Should_Return_ArrayLookupNavigator_When_Array_Is_Supplied()
        {
            var args = new LookupArguments(FunctionsHelper.CreateArgs(8, FunctionsHelper.CreateArgs(1,2), 1), ParsingContext.Create(_excelPackage));
            var navigator = LookupNavigatorFactory.Create(LookupDirection.Horizontal, args, _context);
            Assert.IsInstanceOfType(navigator, typeof(ArrayLookupNavigator));
        }
    }
}

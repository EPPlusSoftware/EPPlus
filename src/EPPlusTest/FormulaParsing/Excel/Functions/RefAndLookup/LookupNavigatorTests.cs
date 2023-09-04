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
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using FakeItEasy;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using EPPlusTest.FormulaParsing.TestHelpers;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace EPPlusTest.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class LookupNavigatorTests
    {
        const string WorksheetName = "";
        private LookupArguments GetArgs(params object[] args)
        {
            var lArgs = FunctionsHelper.CreateArgs(args);
            return new LookupArguments(lArgs, ParsingContext.Create());
        }

        private ParsingContext GetContext(ExcelDataProvider provider)
        {
            var ctx = ParsingContext.Create();
            ctx.Scopes.NewScope(new RangeAddress(){Worksheet = WorksheetName, FromCol = 1, FromRow = 1});
            ctx.ExcelDataProvider = provider;
            return ctx;
        }

        //[TestMethod]
        //public void NavigatorShouldEvaluateFormula()
        //{
        //    var provider = MockRepository.GenerateStub<ExcelDataProvider>();
        //    provider.Stub(x => x.GetCellValue(WorksheetName,0, 0)).Return(new ExcelCell(3);
        //    provider.Stub(x => x.GetCellValue(WorksheetName,1, 0)).Return("B5");
        //    var args = GetArgs(4, "A1:B2", 1);
        //    var context = GetContext(provider);
        //    var parser = MockRepository.GenerateMock<FormulaParser>(provider);
        //    context.Parser = parser;
        //    var navigator = new LookupNavigator(LookupDirection.Vertical, args, context);
        //    navigator.MoveNext();
        //    parser.AssertWasCalled(x => x.Parse("B5"));
        //}

        [TestMethod, Ignore]
        public void CurrentValueShouldBeFirstCell()
        {
            var provider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 1)).Returns(3);
            A.CallTo(() => provider.GetCellValue(WorksheetName,2, 1)).Returns(4);
            var args = GetArgs(3, "A1:B2", 1);
            var navigator = LookupNavigatorFactory.Create(LookupDirection.Vertical, args, GetContext(provider));
            Assert.AreEqual(3, navigator.CurrentValue);
        }

        [TestMethod, Ignore]
        public void MoveNextShouldReturnFalseIfLastCell()
        {
            var provider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns(3);
            A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 1)).Returns(4);
            var args = GetArgs(3, "A1:B1", 1);
            var navigator = LookupNavigatorFactory.Create(LookupDirection.Vertical, args, GetContext(provider));
            Assert.IsFalse(navigator.MoveNext());
        }

        [TestMethod]
        public void HasNextShouldBeTrueIfNotLastCell()
        {
            var provider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => provider.GetDimensionEnd(A<string>.Ignored)).Returns(new ExcelCellAddress(5, 5));
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 1)).Returns(3);
            A.CallTo(() => provider.GetCellValue(WorksheetName,2, 1)).Returns(4);
            var args = GetArgs(3, "A1:B2", 1);
            var navigator = LookupNavigatorFactory.Create(LookupDirection.Vertical, args, GetContext(provider));
            Assert.IsTrue(navigator.MoveNext());
        }

        [TestMethod, Ignore]
        public void MoveNextShouldNavigateVertically()
        {
            var provider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 1)).Returns(3);
            A.CallTo(() => provider.GetCellValue(WorksheetName,2, 1)).Returns(4);
            A.CallTo(() => provider.GetDimensionEnd(A<string>.Ignored)).Returns(new ExcelCellAddress(100, 10));
            var args = GetArgs(6, "A1:B2", 1);
            var navigator = LookupNavigatorFactory.Create(LookupDirection.Vertical, args, GetContext(provider));
            navigator.MoveNext();
            Assert.AreEqual(4, navigator.CurrentValue);
        }

        [TestMethod]
        public void MoveNextShouldIncreaseIndex()
        {
            var provider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => provider.GetDimensionEnd(A<string>.Ignored)).Returns(new ExcelCellAddress(5, 5));
            A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns(3);
            A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 2)).Returns(4);
            var args = GetArgs(6, "A1:B2", 1);
            var navigator = LookupNavigatorFactory.Create(LookupDirection.Vertical, args, GetContext(provider));
            Assert.AreEqual(0, navigator.Index);
            navigator.MoveNext();
            Assert.AreEqual(1, navigator.Index);
        }

        [TestMethod, Ignore]
        public void GetLookupValueShouldReturnCorrespondingValue()
        {
            var provider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => provider.GetDimensionEnd(A<string>.Ignored)).Returns(new ExcelCellAddress(5, 5));
            A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns(3);
            A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 2)).Returns(4);
            var args = GetArgs(6, "A1:B2", 2);
            var navigator = LookupNavigatorFactory.Create(LookupDirection.Vertical, args, GetContext(provider));
            Assert.AreEqual(4, navigator.GetLookupValue());
        }

        [TestMethod, Ignore]
        public void GetLookupValueShouldReturnCorrespondingValueWithOffset()
        {
            var provider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => provider.GetDimensionEnd(A<string>.Ignored)).Returns(new ExcelCellAddress(5, 5));
            A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns(3);
            A.CallTo(() => provider.GetCellValue(WorksheetName, 3, 3)).Returns(4);
            var args = new LookupArguments(3, "A1:A4", 3, 2, false,null);
            var navigator = LookupNavigatorFactory.Create(LookupDirection.Vertical, args, GetContext(provider));
            Assert.AreEqual(4, navigator.GetLookupValue());
        }
    }
}

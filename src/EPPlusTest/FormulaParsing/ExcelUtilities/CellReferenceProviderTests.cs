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
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using FakeItEasy;

namespace EPPlusTest.ExcelUtilities
{
    [TestClass]
    public class CellReferenceProviderTests
    {
        private ExcelDataProvider _provider;

        [TestInitialize]
        public void Setup()
        {
            _provider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => _provider.ExcelMaxRows).Returns(5000);
        }

        [TestMethod]
        public void ShouldReturnReferencedSingleAddress()
        {
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(FormulaRangeAddress.Empty);
            parsingContext.Configuration.SetLexer(new Lexer(parsingContext.Configuration.FunctionRepository, parsingContext.NameValueProvider));
            parsingContext.RangeAddressFactory = new RangeAddressFactory(_provider, parsingContext);
            var provider = new CellReferenceProvider();
            var result = provider.GetReferencedAddresses("A1", parsingContext);
            Assert.AreEqual("A1", result.First());
        }

        [TestMethod]
        public void ShouldReturnReferencedMultipleAddresses()
        {
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(FormulaRangeAddress.Empty);
            parsingContext.Configuration.SetLexer(new Lexer(parsingContext.Configuration.FunctionRepository, parsingContext.NameValueProvider));
            parsingContext.RangeAddressFactory = new RangeAddressFactory(_provider, parsingContext);
            var provider = new CellReferenceProvider();
            var result = provider.GetReferencedAddresses("A1:A2", parsingContext);
            Assert.AreEqual("A1", result.First());
            Assert.AreEqual("A2", result.Last());
        }
    }
}

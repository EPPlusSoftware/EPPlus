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
using OfficeOpenXml.Filter;

namespace EPPlusTest.Filter
{
    [TestClass]
    public class FilterWildCardMatcherTests
    {
        [TestMethod]
        public void MatchBeginingWith()
        {
            Assert.IsTrue(FilterWildCardMatcher.Match("Value 1", "val*"));
        }
        [TestMethod]
        public void MatchEndsWith()
        {
            Assert.IsTrue(FilterWildCardMatcher.Match("Value 1", "*ue 1"));
        }
        [TestMethod]
        public void MatchEndsWithDouble()
        {
            Assert.IsTrue(FilterWildCardMatcher.Match("Value 11", "*1"));
        }
        [TestMethod]
        public void MatchContainsSingleChar()
        {
            Assert.IsTrue(FilterWildCardMatcher.Match("Value 1", "val?e 1"));
        }
        [TestMethod]
        public void MatchContains()
        {
            Assert.IsTrue(FilterWildCardMatcher.Match("Value 1", "*ue*"));
        }
        [TestMethod]
        public void MatchContainsDouble()
        {
            Assert.IsTrue(FilterWildCardMatcher.Match("Valuee 1", "*u*e*1"));
        }
        [TestMethod]
        public void DontMatchContainsDouble()
        {
            Assert.IsFalse(FilterWildCardMatcher.Match("Valuee 1", "*u*e*2"));
        }
        [TestMethod]
        public void MatchContainsAndSingleChar()
        {
            Assert.IsTrue(FilterWildCardMatcher.Match("Value 1", "*??ue*"));
        }
        [TestMethod]
        public void DontMatchContainsAndSingleChar()
        {
            Assert.IsFalse(FilterWildCardMatcher.Match("Value 1", "*??alue*"));
        }
        [TestMethod]
        public void MatchSingleCharsFull()
        {
            Assert.IsTrue(FilterWildCardMatcher.Match("Value 1", "???????"));
        }
        [TestMethod]
        public void DontMatchSingleCharsFullLess()
        {
            Assert.IsFalse(FilterWildCardMatcher.Match("Value 1", "??????"));
        }
        [TestMethod]
        public void DontMatchSingleCharsFullMore()
        {
            Assert.IsFalse(FilterWildCardMatcher.Match("Value 1", "????????"));
        }
        [TestMethod]
        public void MatchBlank()
        {
            Assert.IsTrue(FilterWildCardMatcher.Match("", ""));
        }
        [TestMethod]
        public void MatchBlankAll()
        {
            Assert.IsTrue(FilterWildCardMatcher.Match("", "*"));
        }
    }
}

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
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace EPPlusTest.ExcelUtilities
{
    [TestClass]
    public class WildCardValueMatcher2Tests
    {
        private WildCardValueMatcher2 _matcher;

        [TestInitialize]
        public void Setup()
        {
            _matcher = new WildCardValueMatcher2();
        }

        [TestMethod]
        public void IsMatchShouldReturn0WhenSingleCharWildCardMatches()
        {
            var string1 = "a?c?";
            var string2 = "abcd";
            var result = _matcher.IsMatch(string1, string2);
            Assert.AreEqual(0, result);
        }

        [TestMethod]
        public void IsMatchShouldReturn0WhenMultipleCharWildCardMatches()
        {
            var string1 = "a*c.";
            var string2 = "abcc.";
            var result = _matcher.IsMatch(string1, string2);
            Assert.AreEqual(0, result);
        }

        [TestMethod]
        public void ShouldHandleTildeAndAsterisk1()
        {
            var string1 = "a*c";
            var string2 = "abc";
            var result1 = _matcher.IsMatch("a~*c", string1);
            Assert.AreEqual(0, result1);
            var result2 = _matcher.IsMatch("a~*c", string2);

            Assert.AreEqual(-1, result2);
        }

        [TestMethod]
        public void ShouldHandleTildeAndAsterisk2()
        {
            var string1 = "a*cde";
            var string2 = "abcd";
            var result1 = _matcher.IsMatch("a~*c*", string1);
            Assert.AreEqual(0, result1);
            var result2 = _matcher.IsMatch("a~*c*", string2);

            Assert.AreEqual(-1, result2);
        }

        [TestMethod]
        public void ShouldHandleTildeAndQuestionMark1()
        {
            var string1 = "a?c";
            var string2 = "abc";
            var result1 = _matcher.IsMatch("a~?c", string1);
            Assert.AreEqual(0, result1);
            var result2 = _matcher.IsMatch("a~?c", string2);

            Assert.AreEqual(-1, result2);
        }

        [TestMethod]
        public void ShouldHandleTildeAndQuestionMark2()
        {
            var string1 = "a?cde";
            var string2 = "abcde";
            var result1 = _matcher.IsMatch("a~?c?e", string1);
            Assert.AreEqual(0, result1);
            var result2 = _matcher.IsMatch("a~?c?e", string2);

            Assert.AreEqual(-1, result2);
        }
    }
}

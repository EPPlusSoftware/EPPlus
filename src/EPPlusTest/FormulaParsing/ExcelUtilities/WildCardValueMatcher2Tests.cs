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
using System.Diagnostics;
using Microsoft.VisualBasic;

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
            var pattern = "a?c?";
            var candidate = "abcd";
            var result = _matcher.IsMatch(pattern, candidate);
            Assert.AreEqual(0, result);
        }

        [TestMethod]
        public void ShouldHandleCandidateShorterThanPattern()
        {
            var pattern = "*~*";
            var candidate = "#";
            var result = _matcher.IsMatch(pattern, candidate);
            Assert.AreNotEqual(0, result);
        }

        [TestMethod]
        public void ShouldHandleMultiCharMatch_Match1()
        {
            var string1 = "a*c";
            var string2 = "a123c654564abc";
            var result = _matcher.IsMatch(string1, string2);
            Assert.AreEqual(0, result);
        }

        [TestMethod]
        public void ShouldHandleMultiCharMatch_Match2()
        {
            // TODO: make this work...
            var string1 = "*c?a*";
            var string2 = "ac2ac654564abc";
            var result = _matcher.IsMatch(string1, string2);
            Assert.AreEqual(0, result);
        }

        [TestMethod]
        public void ShouldHandleMultiCharMatch_Match3()
        {
            var string1 = "a*ca*cade?";
            var string2 = "abcabcadef";
            var result = _matcher.IsMatch(string1, string2);
            Assert.AreEqual(0, result);
        }

        [TestMethod]
        public void ShouldHandleMultiCharMatch_NoMatch1()
        {
            var string1 = "a*c";
            var string2 = "a123c654564abd";
            var result = _matcher.IsMatch(string1, string2);
            Assert.AreNotEqual(0, result);
        }

        [TestMethod]
        public void ShouldHandleMultiCharMatch_NoMatch2()
        {
            var string1 = "a*ca*cade?";
            var string2 = "abcadddcade";
            var result = _matcher.IsMatch(string1, string2);
            Assert.AreNotEqual(0, result);
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
        public void ShouldHandleTildeAndAsterisk3()
        {
            var string1 = "a*";
            var result1 = _matcher.IsMatch("*~*", string1);
            Assert.AreEqual(0, result1);
        }

        [TestMethod]
        public void ShouldHandleTildeAndAsterisk4()
        {
            var string1 = "*a";
            var result1 = _matcher.IsMatch("~*a", string1);
            Assert.AreEqual(0, result1);
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

        [TestMethod]
        public void ShouldHandleTildeAndQuestionMark3()
        {
            var string1 = "?c";
            var result1 = _matcher.IsMatch("~?c", string1);
            Assert.AreEqual(0, result1);
        }

        [TestMethod]
        public void ShouldHandleTilde1()
        {
            var string1 = "~";
            var result1 = _matcher.IsMatch("~", string1);
            Assert.AreEqual(0, result1);
        }

        [TestMethod]
        public void ShouldHandleTilde2()
        {
            var string1 = "a~b";
            var result1 = _matcher.IsMatch("a~b", string1);
            Assert.AreEqual(0, result1);
        }

        [TestMethod]
        public void ShouldHandleTilde3()
        {
            var string1 = "a~b";
            var result1 = _matcher.IsMatch("a~~?", string1);
            Assert.AreEqual(0, result1);
        }

        [TestMethod]
        public void ShouldHandleTilde4()
        {
            var string1 = "a~?";
            var result1 = _matcher.IsMatch("a~~~?", string1);
            Assert.AreEqual(0, result1);
            var string2 = "a~b";
            var result2 = _matcher.IsMatch("a~~~?", string2);
            Assert.AreNotEqual(0, result2);
        }

        [TestMethod]
        public void ShouldHandleNull()
        {
            string string2 = default;
            var result2 = _matcher.IsMatch("a~?c?e", string2);
            Assert.AreNotEqual(0, result2);
        }

        [TestMethod]
        public void ShouldHandleEmptyString()
        {
            var string2 = string.Empty;
            var result2 = _matcher.IsMatch("a~?c?e", string2);
            Assert.AreNotEqual(0, result2);
        }

        [TestMethod]
        public void ShouldHandleWhitespace1()
        {
            var string2 = " ";
            var result2 = _matcher.IsMatch("a~?c?e", string2);
            Assert.AreNotEqual(0, result2);
        }

        [TestMethod]
        public void ShouldHandleWhitespace2()
        {
            var string2 = " ";
            var result2 = _matcher.IsMatch(" ", string2);
            Assert.AreEqual(0, result2);
        }
    }
}

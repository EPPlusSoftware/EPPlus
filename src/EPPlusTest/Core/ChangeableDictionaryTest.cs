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
using OfficeOpenXml.Core;
using System.Diagnostics;

namespace EPPlusTest.Core
{
    [TestClass]
    public class ChangeableDictionaryTest
    {
        [TestMethod]
        public void Add()
        {
            ChangeableDictionary<int> d = Setup();

            //Validate order
            var prev = -1;
            foreach (var item in d)
            {
                Assert.IsTrue(prev < item);
            }

            //Validate values
            Assert.AreEqual(20, d[2]);
            Assert.AreEqual(60, d[6]);
            Assert.AreEqual(90, d[9]);
        }
        [TestMethod]
        public void Insert()
        {
            ChangeableDictionary<int> d = Setup();

            d.InsertAndShift(5, 1);
            d.Add(5, 500);
            Assert.AreEqual(500, d[5]);
            Assert.AreEqual(50, d[6]);
            Assert.AreEqual(80, d[9]);
            Assert.AreEqual(90, d[10]);
        }
        [TestMethod]
        public void Remove()
        {
            ChangeableDictionary<int> d = Setup();

            d.RemoveAndShift(5);
            Assert.AreEqual(20, d[2]);
            Assert.AreEqual(40, d[4]);
            Assert.AreEqual(60, d[5]);
            Assert.AreEqual(70, d[6]);
        }

        private static ChangeableDictionary<int> Setup()
        {
            var d = new ChangeableDictionary<int>();
            d.Add(2, 20);
            d.Add(8, 80);
            d.Add(5, 50);
            d.Add(0, 0);
            d.Add(3, 30);
            d.Add(6, 60);
            d.Add(4, 40);
            d.Add(7, 70);
            d.Add(9, 90);
            return d;
        }
    }
}

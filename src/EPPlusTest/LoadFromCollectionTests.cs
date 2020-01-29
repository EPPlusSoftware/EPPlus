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
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace EPPlusTest
{
    [TestClass]
    public class LoadFromCollectionTests
    {
        internal abstract class BaseClass
        {
            public string Id { get; set; }
            public string Name { get; set; }
        }

        internal class Implementation : BaseClass
        {
            public int Number { get; set; }
        }

        internal class Aclass
        {
            public string Id { get; set; }
            public string Name { get; set; }
            public int Number { get; set; }
        }

        [TestMethod]
        public void ShouldUseAclassProperties()
        {
            var items = new List<Aclass>()
            {
                new Aclass(){ Id = "123", Name = "Item 1", Number = 3}
            };
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("sheet");
                sheet.Cells["C1"].LoadFromCollection(items, true, TableStyles.Dark1);

                Assert.AreEqual("Id", sheet.Cells["C1"].Value);
            }
        }

        [TestMethod]
        public void ShouldUseBaseClassProperties()
        {
            var items = new List<BaseClass>()
            {
                new Implementation(){ Id = "123", Name = "Item 1", Number = 3}
            };
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("sheet");
                sheet.Cells["C1"].LoadFromCollection(items, true, TableStyles.Dark1);

                Assert.AreEqual("Id", sheet.Cells["C1"].Value);
            }
        }

        [TestMethod]
        public void ShouldUseAnonymousProperties()
        {
            var objs = new List<BaseClass>()
            {
                new Implementation(){ Id = "123", Name = "Item 1", Number = 3}
            };
            var items = objs.Select(x => new {Id = x.Id, Name = x.Name}).ToList();
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("sheet");
                sheet.Cells["C1"].LoadFromCollection(items, true, TableStyles.Dark1);

                Assert.AreEqual("Id", sheet.Cells["C1"].Value);
            }
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidCastException))]
        public void ShouldThrowInvalidCastExceptionIf()
        {
            var objs = new List<BaseClass>()
            {
                new Implementation(){ Id = "123", Name = "Item 1", Number = 3}
            };
            var items = objs.Select(x => new { Id = x.Id, Name = x.Name }).ToList();
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("sheet");
                sheet.Cells["C1"].LoadFromCollection(items, true, TableStyles.Dark1, BindingFlags.Public | BindingFlags.Instance, typeof(string).GetMembers());

                Assert.AreEqual("Id", sheet.Cells["C1"].Value);
            }
        }
    }
}

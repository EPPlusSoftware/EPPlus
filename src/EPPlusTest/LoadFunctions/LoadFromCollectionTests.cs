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
using System.ComponentModel;
using System.Diagnostics;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.LoadFunctions.Params;
using OfficeOpenXml.Table;

namespace EPPlusTest.LoadFunctions
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

        internal class BClass
        {
            [DisplayName("MyId")]
            public string Id { get; set; }
            [System.ComponentModel.Description("MyName")]
            public string Name { get; set; }
            public int Number { get; set; }
        }

        internal class CamelCasedClass
        {
            public string IdOfThisInstance { get; set; }

            public string CamelCased_And_Underscored { get; set; }
        }

        [TestMethod]
        public void ShouldNotIncludeHeadersWhenPrintHeadersIsOmitted()
        {
            var items = new List<Aclass>()
            {
                new Aclass(){ Id = "123", Name = "Item 1", Number = 3},
                new Aclass(){ Id = "456", Name = "Item 2", Number = 6}
            };
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("sheet");
                sheet.Cells["C1"].LoadFromCollection(items);

                Assert.AreEqual("123", sheet.Cells["C1"].Value);
                Assert.AreEqual(6, sheet.Cells["E2"].Value);
                Assert.AreEqual(3, sheet.Dimension._fromCol);
                Assert.AreEqual(5, sheet.Dimension._toCol);
                Assert.AreEqual(1, sheet.Dimension._fromRow);
                Assert.AreEqual(2, sheet.Dimension._toRow);
            }
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
                Assert.AreEqual("123", sheet.Cells["C2"].Value);
            }
        }

        [TestMethod]
        public void ShouldUseSelectedMembers()
        {
            var items = new List<Aclass>()
            {
                new Aclass(){ Id = "123", Name = "Item 1", Number = 3}
            };
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("sheet");
                var mi = typeof(Aclass)
                    .GetProperties()
                    .Where(pi => pi.Name != "Name")
                    .Select(pi => (MemberInfo)pi)
                    .ToArray();

                sheet.Cells.LoadFromCollection(items, false, TableStyles.None, BindingFlags.Instance | BindingFlags.Public, mi);

                Assert.AreEqual("123", sheet.Cells["A1"].Value);
                Assert.AreEqual(3, sheet.Cells["B1"].Value);
            }
        }

        [TestMethod]
        public void OneMemberShouldShowValue()
        {
            var items = new List<Aclass>()
            {
                new Aclass(){ Id = "123", Name = "Item 1", Number = 3}
            };
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("sheet");
                var mi = typeof(Aclass)
                    .GetProperties()
                    .Where(pi => pi.Name == "Name")
                    .Select(pi => (MemberInfo)pi)
                    .ToArray();

                sheet.Cells.LoadFromCollection(items, false, TableStyles.None, BindingFlags.Instance | BindingFlags.Public, mi);

                Assert.AreEqual("Item 1", sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void ShouldUseDisplayNameAttribute()
        {
            var items = new List<BClass>()
            {
                new BClass(){ Id = "123", Name = "Item 1", Number = 3}
            };
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("sheet");
                sheet.Cells["C1"].LoadFromCollection(items, true, TableStyles.Dark1);

                Assert.AreEqual("MyId", sheet.Cells["C1"].Value);
            }
        }

        [TestMethod]
        public void ShouldFilterMembers()
        {
            var items = new List<BaseClass>()
            {
                new Implementation(){ Id = "123", Name = "Item 1", Number = 3}
            };
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("sheet");
                var t = typeof(Implementation);
                sheet.Cells["C1"].LoadFromCollection(items, true, TableStyles.Dark1, LoadFromCollectionParams.DefaultBindingFlags,
                    new MemberInfo[]
                    {
                        t.GetProperty("Id"),
                        t.GetProperty("Name")
                    });

                Assert.AreEqual(1, sheet.Dimension._toCol - sheet.Dimension._fromCol);
                Assert.AreEqual("Id", sheet.Cells["C1"].Value);
                Assert.AreEqual("Name", sheet.Cells["D1"].Value);
                Assert.IsNull(sheet.Cells["E1"].Value);
                Assert.AreEqual("123", sheet.Cells["C2"].Value);
                Assert.AreEqual("Item 1", sheet.Cells["D2"].Value);
                Assert.IsNull(sheet.Cells["E2"].Value);
            }
        }

        [TestMethod]
        public void ShouldUseDescriptionAttribute()
        {
            var items = new List<BClass>()
            {
                new BClass(){ Id = "123", Name = "Item 1", Number = 3}
            };
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("sheet");
                sheet.Cells["C1"].LoadFromCollection(items, true, TableStyles.Dark1);

                Assert.AreEqual("MyName", sheet.Cells["D1"].Value);
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
            var items = objs.Select(x => new { Id = x.Id, Name = x.Name }).ToList();
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

        [TestMethod]
        public void ShouldUseLambdaConfig()
        {
            var items = new List<Aclass>()
            {
                new Aclass(){ Id = "123", Name = "Item 1", Number = 3}
            };
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("sheet");
                sheet.Cells["C1"].LoadFromCollection(items, c =>
                {
                    c.PrintHeaders = true;
                    c.TableStyle = TableStyles.Dark1;
                });
                Assert.AreEqual("Id", sheet.Cells["C1"].Value);
                Assert.AreEqual("123", sheet.Cells["C2"].Value);
                Assert.AreEqual(3, sheet.Cells["E2"].Value);
                Assert.AreEqual(1, sheet.Tables.Count());
            }
        }

        [TestMethod]
        public void ShouldParseCamelCasedHeaders()
        {
            var items = new List<CamelCasedClass>()
            {
                new CamelCasedClass(){ IdOfThisInstance = "123" }
            };
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("sheet");
                sheet.Cells["C1"].LoadFromCollection(items, c =>
                {
                    c.PrintHeaders = true;
                    c.HeaderParsingType = HeaderParsingTypes.CamelCaseToSpace;
                });
                Assert.AreEqual("Id Of This Instance", sheet.Cells["C1"].Value);
            }
        }

        [TestMethod]
        public void ShouldParseCamelCasedAndUnderscoredHeaders()
        {
            var items = new List<CamelCasedClass>()
            {
                new CamelCasedClass(){ CamelCased_And_Underscored = "123" }
            };
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("sheet");
                sheet.Cells["C1"].LoadFromCollection(items, c =>
                {
                    c.PrintHeaders = true;
                    c.HeaderParsingType = HeaderParsingTypes.UnderscoreAndCamelCaseToSpace;
                });
                Assert.AreEqual("Camel Cased And Underscored", sheet.Cells["D1"].Value);
            }
        }

        [TestMethod]
        public void ShouldLoadExpandoObjects()
        {
            dynamic o1 = new ExpandoObject();
            o1.Id = 1;
            o1.Name = "TestName 1";
            dynamic o2 = new ExpandoObject();
            o2.Id = 2;
            o2.Name = "TestName 2";
            var items = new List<ExpandoObject>()
            {
                o1,
                o2
            };
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                var r = sheet.Cells["A1"].LoadFromCollection(items, true, TableStyles.None);

                Assert.AreEqual("Id", sheet.Cells["A1"].Value);
                Assert.AreEqual(1, sheet.Cells["A2"].Value);
                Assert.AreEqual("TestName 2", sheet.Cells["B3"].Value);
            }
        }
    }
}

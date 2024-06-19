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
using OfficeOpenXml.Attributes;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.LoadFunctions.Params;
using OfficeOpenXml.Table;

namespace EPPlusTest.LoadFunctions
{
    [TestClass]
    public class LoadFromCollectionTests : TestBase
    {
        [EpplusTable(AutofitColumns = true, PrintHeaders = true, TableStyle = TableStyles.Light10)]
        internal class Company
        {
            public Company(int id, string name, Uri url)
            {
                Id = id;
                Name = name;
                Url = url;
            }

            [EpplusTableColumn(Header = "Id", Order = 1)]
            public int Id
            {
                get; set;
            }

            [EpplusTableColumn(Header = "Name", Order = 2)]
            public string Name { get; set; }

            [EpplusTableColumn(Header = "Homepage", Order = 3)]
            public Uri Url { get; set; }

        }

        internal abstract class BaseClass
        {
            public string Id { get; set; }
            public string Name { get; set; }
        }

        internal class Implementation : BaseClass
        {
            public int Number { get; set; }
        }
        [System.ComponentModel.Description("The color Red")]
        internal enum AnEnum
        {
            [System.ComponentModel.Description("The color Red")]
            Red,
            [System.ComponentModel.Description("The color Blue")]
            Blue,            
            Green
        }
        internal class EnumClass
        {
            public int Id { get; set; }
            public AnEnum Enum { get; set; }
            [System.ComponentModel.Description("Nullable Enum")]
            public AnEnum? NullableEnum{ get; set; }
        }
        internal class Aclass
        {
            public string Id { get; set; }
            public string Name { get; set; }
            public int Number { get; set; }
        }

        internal class BClass
        {
            [DisplayName("MyId"), EpplusTableColumn(Order = 1)]
            public string Id { get; set; }
            [System.ComponentModel.Description("MyName"), EpplusTableColumn(Order = 2)]
            public string Name { get; set; }
            [EpplusTableColumn(Order = 3)]
            public int Number { get; set; }
        }

        internal class CClass
        {
            [DisplayName("Another property")]
            public string AnotherProperty { get; set; }
        }

        internal class CamelCasedClass
        {
            public string IdOfThisInstance { get; set; }

            public string CamelCased_And_Underscored { get; set; }
        }

        internal class UrlClass : BClass
        {
            [EpplusIgnore]
            public string EMailAddress { get; set; }
            [EpplusTableColumn(Order = 5, Header = "My Mail To")]
            public ExcelHyperLink MailTo
            {
                get
                {
                    var url = new ExcelHyperLink("mailto:" + EMailAddress);
                    url.Display = Name;
                    return url;
                }
            }
            [EpplusTableColumn(Order = 4)]
            public Uri Url
            {
                get;
                set;
            }
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
        public void ShouldNotIncludeHeadersWhenPrintHeadersIsOmittedTransposed()
        {
            var items = new List<Aclass>()
            {
                new Aclass(){ Id = "123", Name = "Item 1", Number = 3},
                new Aclass(){ Id = "456", Name = "Item 2", Number = 6}
            };
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("sheet");
                sheet.Cells["C1"].LoadFromCollection(items, c =>
                {
                    c.Transpose = true;
                });

                Assert.AreEqual("123", sheet.Cells["C1"].Value);
                Assert.AreEqual(6, sheet.Cells["D3"].Value);
                Assert.AreEqual(3, sheet.Dimension._fromCol);
                Assert.AreEqual(4, sheet.Dimension._toCol);
                Assert.AreEqual(1, sheet.Dimension._fromRow);
                Assert.AreEqual(3, sheet.Dimension._toRow);
            }
        }

        [TestMethod]
        public void ShouldIncludeHeaders()
        {
            var items = new List<Aclass>()
            {
                new Aclass(){ Id = "123", Name = "Item 1", Number = 3},
                new Aclass(){ Id = "456", Name = "Item 2", Number = 6}
            };
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("sheet");
                sheet.Cells["C1"].LoadFromCollection(items, true);
                Assert.AreEqual("Id", sheet.Cells["C1"].Value);
            }
        }

        [TestMethod]
        public void ShouldIncludeHeadersTransposed()
        {
            var items = new List<Aclass>()
            {
                new Aclass(){ Id = "123", Name = "Item 1", Number = 3},
                new Aclass(){ Id = "456", Name = "Item 2", Number = 6}
            };
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("sheet");
                sheet.Cells["C1"].LoadFromCollection(items, c =>
                {
                    c.PrintHeaders = true;
                    c.Transpose = true;
                });
                Assert.AreEqual("Id", sheet.Cells["C1"].Value);
                Assert.AreEqual("123", sheet.Cells["D1"].Value);
                Assert.AreEqual("456", sheet.Cells["E1"].Value);
                Assert.AreEqual("Name", sheet.Cells["C2"].Value);
                Assert.AreEqual("Item 1", sheet.Cells["D2"].Value);
                Assert.AreEqual("Item 2", sheet.Cells["E2"].Value);
                Assert.AreEqual("Number", sheet.Cells["C3"].Value);
                Assert.AreEqual(3, sheet.Cells["D3"].Value);
                Assert.AreEqual(6, sheet.Cells["E3"].Value);
            }
        }

        [TestMethod]
        public void ShouldIncludeHeadersAndTableStyle()
        {
            var items = new List<Aclass>()
            {
                new Aclass(){ Id = "123", Name = "Item 1", Number = 3},
                new Aclass(){ Id = "456", Name = "Item 2", Number = 6}
            };
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("sheet");
                sheet.Cells["C1"].LoadFromCollection(items, true, TableStyles.Dark1);
                Assert.AreEqual("Id", sheet.Cells["C1"].Value);
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
        public void ShouldFilterMembersTransposed()
        {
            var items = new List<BaseClass>()
            {
                new Implementation(){ Id = "123", Name = "Item 1", Number = 3}
            };
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("sheet");
                var t = typeof(Implementation);
                sheet.Cells["C1"].LoadFromCollection(items, true, TableStyles.Dark1, true, LoadFromCollectionParams.DefaultBindingFlags,
                    new MemberInfo[]
                    {
                        t.GetProperty("Id"),
                        t.GetProperty("Name")
                    });

                Assert.AreEqual(1, sheet.Dimension._toCol - sheet.Dimension._fromCol);
                Assert.AreEqual("Id", sheet.Cells["C1"].Value);
                Assert.AreEqual("Name", sheet.Cells["C2"].Value);
                Assert.IsNull(sheet.Cells["C3"].Value);
                Assert.AreEqual("123", sheet.Cells["D1"].Value);
                Assert.AreEqual("Item 1", sheet.Cells["D2"].Value);
                Assert.IsNull(sheet.Cells["E2"].Value);
            }
        }

        [TestMethod]
        public void ShouldFilterOneMember()
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
                    });
                Assert.AreEqual("Id", sheet.Cells["C1"].Value);
                Assert.AreEqual("123", sheet.Cells["C2"].Value);
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

#if !NET35
        [TestMethod]
        public void ShouldUseDisplayAttribute()
        {
            var items = new List<CClass>()
            {
                new CClass(){ AnotherProperty = "asdjfkl?"}
            };
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("sheet");
                sheet.Cells["C1"].LoadFromCollection(items, true, TableStyles.Dark1);

                Assert.AreEqual("Another property", sheet.Cells["C1"].Value);
            }
        }
#endif

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
        [TestMethod]
        public void ShouldLoadExpandoObjectsTransposed()
        {
            dynamic o1 = new ExpandoObject();
            o1.Id = 1;
            o1.Name = "TestName 1";
            dynamic o2 = new ExpandoObject();
            o2.Id = 2;
            o2.Name = "TestName 2";
            dynamic o3 = new ExpandoObject();
            o3.Id = 3;
            o3.Name = "TestName 3";
            var items = new List<ExpandoObject>()
            {
                o1,
                o2,
                o3,
            };
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                var r = sheet.Cells["A1"].LoadFromCollection(items, c =>
                {
                    c.PrintHeaders = true;
                    c.Transpose = true;
                });
                Assert.AreEqual("A1:D2", r.Address);
                Assert.AreEqual("Id", sheet.Cells["A1"].Value);
                Assert.AreEqual(1, sheet.Cells["B1"].Value);
                Assert.AreEqual("TestName 2", sheet.Cells["C2"].Value);
            }
        }
        [TestMethod]
        public void ShouldSetHyperlinkForURIs()
        {
            var items = new List<UrlClass>()
            {
                new UrlClass{Id="1", Name="Person 1", EMailAddress="person1@somewhe.re"},
                new UrlClass{Id="2", Name="Person 2", EMailAddress="person2@somewhe.re"},
                new UrlClass{Id="2", Name="Person with Url", EMailAddress="person2@somewhe.re", Url=new Uri("https://epplussoftware.com")},
            };

            using (var package = OpenPackage("LoadFromCollectionUrls.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                var ns = package.Workbook.Styles.CreateNamedStyle("Hyperlink");
                ns.BuildInId = 8;
                ns.Style.Font.UnderLine = true;
                ns.Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(0xFF,0x05 ,0x63, 0xC1));

                var r = sheet.Cells["A1"].LoadFromCollection(items, true, TableStyles.Medium1);
                sheet.Cells["E2:E5"].StyleName = "Hyperlink";

                Assert.AreEqual("MyId", sheet.Cells["A1"].Value);
                Assert.AreEqual("MyName", sheet.Cells["B1"].Value);
                Assert.AreEqual("Number", sheet.Cells["C1"].Value);
                Assert.AreEqual("Url", sheet.Cells["D1"].Value);
                Assert.AreEqual("My Mail To", sheet.Cells["E1"].Value);

                Assert.AreEqual("1", sheet.Cells["A2"].Value);
                Assert.AreEqual("Person 2", sheet.Cells["B3"].Value);
                Assert.IsInstanceOfType(sheet.Cells["E3"].Hyperlink, typeof(ExcelHyperLink));
                Assert.AreEqual("Person 2", sheet.Cells["E3"].Value);

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void TransposeHyperlinks()
        {
            var items = new List<UrlClass>()
            {
                new UrlClass{Id="1", Name="Person 1", EMailAddress="person1@somewhe.re"},
                new UrlClass{Id="2", Name="Person 2", EMailAddress="person2@somewhe.re"},
                new UrlClass{Id="2", Name="Person with Url", EMailAddress="person2@somewhe.re", Url=new Uri("https://epplussoftware.com")},
            };

            using (var package = OpenPackage("LoadFromURIsTranspose.xlsx", true))
            {
                
                var sheet = package.Workbook.Worksheets.Add("test");

                var r = sheet.Cells["A1"].LoadFromCollection(items, true, TableStyles.Medium1, true);

                sheet.Tables[0].SyncColumnNames(ApplyDataFrom.ColumnNamesToCells);
                Assert.AreEqual("Column4", sheet.Cells["D1"].Value);

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void LoadListOfEnumWithDescription()
        {
            var items = new List<AnEnum>()
            {
                AnEnum.Red,
                AnEnum.Green,
                AnEnum.Blue
            };

            using (var package = OpenPackage("LoadFromCollectionEnumDescrAtt.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("EnumList");
                var r = sheet.Cells["A1"].LoadFromCollection(items, true, TableStyles.Medium1);
                Assert.AreEqual("The color Red", sheet.Cells["A1"].Value);
                Assert.AreEqual("Green", sheet.Cells["A2"].Value);
                Assert.AreEqual("The color Blue", sheet.Cells["A3"].Value);
                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void LoadListOfEnumWithDescriptionTransposed()
        {
            var items = new List<AnEnum>()
            {
                AnEnum.Red,
                AnEnum.Green,
                AnEnum.Blue
            };

            using (var package = OpenPackage("LoadFromCollectionEnumDescrAtt.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("EnumList");
                var r = sheet.Cells["A1"].LoadFromCollection(items, true, TableStyles.Medium1, true);
                Assert.AreEqual("The color Red", sheet.Cells["A1"].Value);
                Assert.AreEqual("Green", sheet.Cells["B1"].Value);
                Assert.AreEqual("The color Blue", sheet.Cells["C1"].Value);
                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void LoadListOfNullableEnumWithDescription()
        {
            var items = new List<AnEnum?>()
            {
                AnEnum.Red,
                AnEnum.Green,
                AnEnum.Blue
            };

            using (var package = OpenPackage("LoadFromCollectionNullableEnumDescrAtt.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("NullableEnumList");
                var r = sheet.Cells["A1"].LoadFromCollection(items, true, TableStyles.Medium1);
                Assert.AreEqual("The color Red", sheet.Cells["A1"].Value);
                Assert.AreEqual("Green", sheet.Cells["A2"].Value);
                Assert.AreEqual("The color Blue", sheet.Cells["A3"].Value);
                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void LoadListOfClassWithEnumWithDescription()
        {
            var items = new List<EnumClass>()
            {
                new EnumClass(){Id=1, Enum=AnEnum.Red, NullableEnum = AnEnum.Blue},
                new EnumClass(){Id=2, Enum=AnEnum.Blue, NullableEnum = null},
                new EnumClass(){Id=3, Enum=AnEnum.Green, NullableEnum = AnEnum.Red},
            };

            using (var package = OpenPackage("LoadFromCollectionClassWithEnumDescrAtt.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                var r = sheet.Cells["A1"].LoadFromCollection(items, true, TableStyles.Medium1);
                Assert.AreEqual("Id", sheet.Cells["A1"].Value);
                Assert.AreEqual("Enum", sheet.Cells["B1"].Value);
                Assert.AreEqual("Nullable Enum", sheet.Cells["C1"].Value);
                Assert.AreEqual(1, sheet.Cells["A2"].Value);
                Assert.AreEqual("The color Red", sheet.Cells["B2"].Value);
                Assert.AreEqual("The color Blue", sheet.Cells["C2"].Value);
                Assert.AreEqual(2, sheet.Cells["A3"].Value);
                Assert.AreEqual("The color Blue", sheet.Cells["B3"].Value);
                Assert.IsNull(sheet.Cells["C3"].Value);
                Assert.AreEqual(3, sheet.Cells["A4"].Value);
                Assert.AreEqual("Green", sheet.Cells["B4"].Value);
                Assert.AreEqual("The color Red", sheet.Cells["C4"].Value);

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void LoadListOfClassWithEnumWithDescriptionTransposed()
        {
            var items = new List<EnumClass>()
            {
                new EnumClass(){Id=1, Enum=AnEnum.Red, NullableEnum = AnEnum.Blue},
                new EnumClass(){Id=2, Enum=AnEnum.Blue, NullableEnum = null},
                new EnumClass(){Id=3, Enum=AnEnum.Green, NullableEnum = AnEnum.Red},
            };

            using (var package = OpenPackage("LoadFromCollectionClassWithEnumDescrAtt.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                var r = sheet.Cells["A1"].LoadFromCollection(items, true, TableStyles.Medium1, true);
                Assert.AreEqual("Id", sheet.Cells["A1"].Value);
                Assert.AreEqual("Enum", sheet.Cells["A2"].Value);
                Assert.AreEqual("Nullable Enum", sheet.Cells["A3"].Value);
                Assert.AreEqual(1, sheet.Cells["B1"].Value);
                Assert.AreEqual("The color Red", sheet.Cells["B2"].Value);
                Assert.AreEqual("The color Blue", sheet.Cells["B3"].Value);
                Assert.AreEqual(2, sheet.Cells["C1"].Value);
                Assert.AreEqual("The color Blue", sheet.Cells["C2"].Value);
                Assert.IsNull(sheet.Cells["C3"].Value);
                Assert.AreEqual(3, sheet.Cells["D1"].Value);
                Assert.AreEqual("Green", sheet.Cells["D2"].Value);
                Assert.AreEqual("The color Red", sheet.Cells["D3"].Value);

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void LoadWithAttributesTest()
        {
            var l = new List<Company>();
            l.Add(new Company(1, "EPPlus Software AB", new Uri("https://epplussoftware.com")));

            using (var package = OpenPackage("LoadFromCollectionAttr.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].LoadFromCollection(l, x => x.UseBuiltInStylesForHyperlinks = true);

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void LoadWithAttributesTestTransposed()
        {
            var l = new List<Company>();
            l.Add(new Company(1, "EPPlus Software AB", new Uri("https://epplussoftware.com")));

            using (var package = OpenPackage("LoadFromCollectionAttr.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].LoadFromCollection(l, x => { x.UseBuiltInStylesForHyperlinks = true; x.Transpose = true; });

                SaveAndCleanup(package);
            }
        }
    }
}

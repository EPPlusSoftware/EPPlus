using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace EPPlusTest.LoadFunctions
{
    [TestClass]
    public class LoadFromCollectionNestedTests
    {
        [TestMethod]
        public void ShouldHandleNestedInThreeLevels()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].LoadFromCollection(Person.All, o =>
            {
                o.PrintHeaders = true;
                o.TableStyle = OfficeOpenXml.Table.TableStyles.Dark2;
                o.Members = new MemberInfo[]
                {
                    typeof(Person).GetProperty("FirstName"),
                    typeof(Person).GetProperty("LastName"),
                    typeof(Person).GetProperty("Employment"),
                    typeof(Employment).GetProperty("Employer"),
                    typeof(Employer).GetProperty("Name"),
                };
            });
            Assert.AreEqual("First name", sheet.Cells["A1"].Value);
            Assert.AreEqual("LastName", sheet.Cells["B1"].Value);
            Assert.AreEqual("Employment Employer", sheet.Cells["C1"].Value);
            Assert.AreEqual("John", sheet.Cells["A2"].Value);
            Assert.AreEqual("Doe", sheet.Cells["B2"].Value);
            Assert.AreEqual("Acme Inc", sheet.Cells["C2"].Value);
            Assert.IsNull(sheet.Cells["D1"].Value);
            Assert.IsNull(sheet.Cells["D2"].Value);
        }

        [EpplusTable(TableStyle = OfficeOpenXml.Table.TableStyles.Dark5, AutofitColumns = true, PrintHeaders = true)]
        [EPPlusTableColumnSortOrder(Properties = new[] { nameof(FirstName), nameof(LastName), nameof(Age), nameof(Employment) })]
        internal class Person
        {
            public Person(string firstName, string lastName, int age, Employment employment)
            {
                FirstName = firstName;
                LastName = lastName;
                Age = age;
                Employment = employment;
            }

            [EpplusTableColumn(Header = "First name")]
            public string FirstName { get; set; }

            public string LastName { get; set; }

            [EpplusTableColumn(Header = "Persons age", Order = 2)]
            public int Age { get; set; }

            [EpplusNestedTableColumn(HeaderPrefix = "Employment", Order = 1)]
            public Employment Employment { get; set; }


            public static IEnumerable<Person> All
            {
                get
                {
                    var rnd = new Random();
                    var persons = new List<Person>
                    {
                        new Person(
                            "John",
                            "Doe",
                            rnd.Next(15, 100),
                            new Employment
                            {
                                StartDate = DateTime.Now.AddDays(-1 * rnd.Next(100, 1000)),
                                Salary = rnd.Next(20000, 40000),
                                Employer = new Employer
                                {
                                    Name = "Acme Inc"
                                }
                            })
                    };
                    return persons;
                }
            }
        }


        [EPPlusTableColumnSortOrder(Properties = new[] { nameof(Employer), nameof(Salary), nameof(StartDate) })]
        internal class Employment
        {
            [EpplusTableColumn(Header = "Start date")]
            public DateTime StartDate { get; set; }

            public decimal Salary { get; set; }

#nullable enable
            [EpplusNestedTableColumn]
            public Employer? Employer { get; set; }
#nullable disable
        }

        internal class Employer
        {
            [EpplusTableColumn(Header = "Employer")]
            public string Name { get; set; }
        }
    }
}

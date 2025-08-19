using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Attributes;
using OfficeOpenXml.LoadFunctions;
using System.Threading;

namespace EPPlusTest.LoadFunctions
{
    [TestClass]
    public class LoadFromCollectionNumberFormatProviderTests
    {
        private ExcelPackage _package;
        private ExcelWorksheet _sheet;

        [TestInitialize]
        public void Initialize()
        {
            _package = new ExcelPackage();
            _sheet = _package.Workbook.Worksheets.Add("Sheet1");
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
            _package = null;
        }

        #region Test classes

        public class MyNumberFormatProvider : IExcelNumberFormatProvider
        {
            public const int SwedishCurrencyFormat = 1;
            string IExcelNumberFormatProvider.GetFormat(int numberFormatId)
            {
                switch(numberFormatId)
                {
                    case SwedishCurrencyFormat:
                        return "#,##0.00\\ \"kr\"";
                    default:
                        return string.Empty;
                } 
            }
        }

        [EpplusTable(NumberFormatProviderType = typeof(MyNumberFormatProvider))]
        public class NumberFormatWithTableAttribute
        {
            [EpplusTableColumn(Header = "First name")]
            public string Name { get; set; }

            [EpplusTableColumn(Header = "Salary", NumberFormatId = MyNumberFormatProvider.SwedishCurrencyFormat)]
            public decimal Salary { get; set; }
        }

        public class NumberFormatWithoutTableAttribute
        {
            [EpplusTableColumn(Header = "First name")]
            public string Name { get; set; }

            [EpplusTableColumn(Header = "Salary", NumberFormatId = MyNumberFormatProvider.SwedishCurrencyFormat)]
            public decimal Salary { get; set; }
        }

        #endregion

        [TestMethod]
        public void ShouldSetNumberFormatFromExcalTableAttribute()
        {
            var currentCulture = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            var items = new List<NumberFormatWithTableAttribute>
            {
                new NumberFormatWithTableAttribute { Name = "Joe", Salary = 1000 }
            };
            _sheet.Cells["A1"].LoadFromCollection(items, o => o.PrintHeaders = true);
            Assert.AreEqual("Joe", _sheet.Cells["A2"].Value);
            Assert.AreEqual(1000m, _sheet.Cells["B2"].Value);
            Assert.AreEqual("1,000.00 kr", _sheet.Cells["B2"].Text);
            Thread.CurrentThread.CurrentCulture = currentCulture;

        }

        [TestMethod]
        public void ShouldSetNumberFormatFromExcalTableAttribute_WhenPrintHeaderIsFalse()
        {
            var currentCulture = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            var items = new List<NumberFormatWithTableAttribute>
            {
                new NumberFormatWithTableAttribute { Name = "Joe", Salary = 1000 }
            };
            _sheet.Cells["A1"].LoadFromCollection(items, o => o.PrintHeaders = false);
            Assert.AreEqual("Joe", _sheet.Cells["A1"].Value);
            Assert.AreEqual(1000m, _sheet.Cells["B1"].Value);
            Assert.AreEqual("1,000.00 kr", _sheet.Cells["B1"].Text);
            Thread.CurrentThread.CurrentCulture = currentCulture;

        }

        [TestMethod, ExpectedException(typeof(ArgumentNullException))]
        public void ShouldThrowIfNoNumberFormatProviderSet()
        {
            var items = new List<NumberFormatWithoutTableAttribute>
            {
                new NumberFormatWithoutTableAttribute { Name = "Joe", Salary = 1000 }
            };
            _sheet.Cells["A1"].LoadFromCollection(items, o => o.PrintHeaders = true);

        }

        [TestMethod]
        public void ShouldUseNumberFormatProviderSetViaParams()
        {
            var currentCulture = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            var items = new List<NumberFormatWithoutTableAttribute>
            {
                new NumberFormatWithoutTableAttribute { Name = "Joe", Salary = 1000 }
            };
            _sheet.Cells["A1"].LoadFromCollection(items, o =>
            {
                o.PrintHeaders = true;
                o.SetNumberFormatProvider(new MyNumberFormatProvider());
            });
            Assert.AreEqual("Joe", _sheet.Cells["A2"].Value);
            Assert.AreEqual(1000m, _sheet.Cells["B2"].Value);
            Assert.AreEqual("1,000.00 kr", _sheet.Cells["B2"].Text);
            Thread.CurrentThread.CurrentCulture = currentCulture;
        }

    }
}

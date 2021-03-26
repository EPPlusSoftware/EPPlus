using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Core
{
    [TestClass]
    public class LoadSaveTests : TestBase
    {
        static ExcelPackage _pck;

        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("LoadSaveTest.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void CheckCfLfIsRetained()
        {
            using (var p1 = new ExcelPackage())
            {
                var expected = "Line1\r\nLine2";
                var ws = p1.Workbook.Worksheets.Add("CrLf");
                ws.Cells["A1"].Value = expected;
                Assert.AreEqual(expected, ws.Cells["A1"].Value);

                ws.Save();

                using (var p2 = new ExcelPackage(p1.Stream))
                {
                    ws = p1.Workbook.Worksheets["CrLf"];
                    Assert.AreEqual(expected, ws.Cells["A1"].Value);
                }
            }
        }
        [TestMethod]
        public void SaveTwiceShouldNotCorruptPackage()
        {
            using (var p = new ExcelPackage())
            {
                var ws=p.Workbook.Worksheets.Add("SaveTwice");
                p.Workbook.Properties.Application = "EPPlus";
                ws.Cells["A1"].Value = "A1";
                p.Workbook.Properties.Title = "EPPlus";
                p.Save();
                var length = p.Stream.Length;
                var b = p.GetAsByteArray();

                Assert.AreEqual(length, b.Length);
            }
        }
        [TestMethod]
        public async Task SaveTwiceShouldNotCorruptPackageAsync()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("SaveTwice");
                p.Workbook.Properties.Application = "EPPlus";
                ws.Cells["A1"].Value = "A1";
                p.Workbook.Properties.Title = "EPPlus";
                p.Save();
                var length = p.Stream.Length;
                var b = await p.GetAsByteArrayAsync();

                Assert.AreEqual(length, b.Length);
            }
        }

        [TestMethod]
        public void ChartSheetShouldNotThrowException()
        {
            using (var p = new ExcelPackage())
            {
                var s1 = p.Workbook.Worksheets.Add("Table1");
                var s2 = p.Workbook.Worksheets.AddChart("Chart1",
                                  OfficeOpenXml.Drawing.Chart.eChartType.Area);
                var s3 = p.Workbook.Worksheets.Add("Table2");

                DataTable dt = new DataTable();

                dt.Columns.Add(new DataColumn("Title", typeof(string)));
                dt.Columns.Add(new DataColumn("Count", typeof(int)));

                var r1 = dt.NewRow();
                var r2 = dt.NewRow();
                var r3 = dt.NewRow();

                r1.ItemArray = new object[] { "Title", 20 };
                r2.ItemArray = new object[] { "Title", 20 };
                r3.ItemArray = new object[] { "Title", 20 };

                dt.Rows.Add(r1);
                dt.Rows.Add(r2);
                dt.Rows.Add(r3);

                s1.Cells[1, 1, 3, 2].LoadFromDataTable(dt, true, OfficeOpenXml.Table.TableStyles.None);
            }
        }

        [TestMethod]
        public void VerifyInvalidXmlUnicodeChar()
        {
            string s1 = "String with \ufffe char";
            string s2 = "Second string with \uffff char";
            using (var p1 = new ExcelPackage())
            {
                var ws = p1.Workbook.Worksheets.Add("Sheet1");
                ws.SetValue(1, 1, s1);
                ws.SetValue(2, 1, s2);
                p1.Save();

                using (var p2 = new ExcelPackage(p1.Stream))
                {
                    Assert.AreEqual(s1, p2.Workbook.Worksheets[0].Cells["A1"].Value);
                    Assert.AreEqual(s2, p2.Workbook.Worksheets[0].Cells["A2"].Value);
                }
            }
        }
    }
}

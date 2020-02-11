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
    public class LoadSaveTests
    {
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
    }
}

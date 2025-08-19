using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace EPPlusTest.Issues
{
    [TestClass]
    public class ImportIssues : TestBase
    {
        [TestMethod]
        public void s823()
        {
            using (var p = OpenTemplatePackage("s823.xlsx"))
            {
                var workbook = p.Workbook;
                string[][] arr = [["I work"]];
                string[][] arr3 = [["I work"], ["I work2"], ["I work3"]];
                var dt = new DataTable();
                dt.Columns.Add("Column1", typeof(string));

                foreach (var row in arr3)
                {
                    dt.Rows.Add(row);
                }

                var workSheet = workbook.Worksheets["Sheet1"];
                workSheet.Cells["B2"].LoadFromArrays(arr);          //This works
                workSheet.Cells["D1"].LoadFromArrays(arr3);         //This doesn't
                workSheet.Cells["F1"].LoadFromDataTable(dt, false); //This doesn't

                SaveAndCleanup(p);
            }
        }
    }
}

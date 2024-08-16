using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Database
{
    [TestClass]
    public class DbFunctionTests : TestBase
    {
        //[TestMethod]

        //public void WorkBookTest()
        //{
        //    var wbPath = "C:\\Users\\HannesAlm\\Downloads\\dbtest.xlsx";
        //    //Same test as from microsoft examples
        //    using (ExcelPackage package = new ExcelPackage(new FileInfo(wbPath)))
        //    {
        //        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

        //        //And set 9

        //        worksheet.Cells["A20"].Formula = "DSTDEV(A5:E11, D5, A1:A3)";
        //        worksheet.Cells["A21"].Formula = "DSTDEVP(A5:E11, D5, A1:A3)";
        //        worksheet.Cells["A22"].Formula = "DPRODUCT(A5:E11, D5, A1:A3)";

        //        worksheet.Calculate();

        //        package.Save();
        //    }
        //}
    }
}

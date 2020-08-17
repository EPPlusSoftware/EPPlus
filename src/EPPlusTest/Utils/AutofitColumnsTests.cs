using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Utils
{
    [TestClass]
    public class AutofitColumnsTests
    {
        [TestMethod]
        public void AutofitCols1()
        {
            const int nCols = 50;
            const int nRows = 5000;
            var chars = "jklöasuiopweqrkljösadf789023478907asfjklöqwe7r89sgfjört903sdidfgjls";
            var rnd = new Random();
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                for(var col = 1; col <= nCols; col++)
                {
                    for(var row = 1; row <= nRows; row++)
                    {
                        var length = rnd.Next(3, chars.Length - 1);
                        if (col % 3 == 0 && length > 40) length -= 40;
                        sheet.Cells[row, col].Value = chars.Substring(0, length);
                    }
                }

                sheet.Cells[1, 1, nRows, nCols].AutoFitColumns();
                package.SaveAs(new FileInfo("c:\\Temp\\autofit.xlsx"));
            }
        }
    }
}

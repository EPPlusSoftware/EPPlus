using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using System.IO;

namespace EPPlusTest.Core.Worksheet
{
    [TestClass]
    public class AutofitMultThreadingTests
    {
        [TestMethod]
        public void Test1()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            var strs = new List<string> { "abcdef", "cdefAB321", "123BJFDOEKSJGF" };
            var rnd = new Random();
            for(var col = 1; col < 40; col++)
            {
                for(var row = 1; row < 10; row++)
                {
                    var sb= new StringBuilder();                    
                    for(var i2 = 0; i2 < rnd.Next(3); i2++)
                    {
                        var i = rnd.Next(strs.Count);
                        var str = strs[i];
                        sb.Append(str);
                    }
                    sheet.Cells[row, col].Value = sb.ToString();
                }
            }
            sheet.Cells[sheet.Dimension.Address].AutoFitColumns2(true);
            var path = @"c:\Temp\AutofitMulti.xlsx";
            if (File.Exists(path)) File.Delete(path);
            package.SaveAs(path);
        }
    }
}

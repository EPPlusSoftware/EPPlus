using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class LinestWorkbookTests : TestBase
    {
        [TestMethod]
        public void Test1()
        {
            using (var package = OpenTemplatePackage(@"LinestTest.xlsx"))
            {
                package.Workbook.Worksheets.Copy("Sheet1", "TestSheet");
                var excelSheet = package.Workbook.Worksheets["Sheet1"];
                var sheet = package.Workbook.Worksheets["TestSheet"];
                sheet.Calculate();
                var sheet2 = package.Workbook.Worksheets.Add("Sheet5");
                for(var c = 1; c <= sheet.Dimension.End.Column; c++)
                {
                    for(var r = 1; r <= sheet.Dimension.End.Row; r++)
                    {
                        sheet2.Cells[r, c].Value = sheet.GetValue(r, c);
                    }
                }
                SaveWorkbook(@"LinestTestResults.xlsx", package);
                //CompareRange(excelSheet, sheet, "O10:R14");
               

            }
        }

        private void CompareRange(ExcelWorksheet excel, ExcelWorksheet sheet, string address)
        {
            var excelRange = excel.Cells[address];
            var xlRange = new object[excelRange.End.Column - excelRange.Start.Column + 1, excelRange.End.Row - excelRange.Start.Row + 1];
            var tRange = new object[excelRange.End.Column - excelRange.Start.Column + 1, excelRange.End.Row - excelRange.Start.Row + 1];
            for (var c = excelRange.Start.Column; c <= excelRange.End.Column; c++)
            {
                for(var r = excelRange.Start.Row; r <= excelRange.End.Row; r++)
                {
                    var cIx = c - excelRange.Start.Column;
                    var rIx = r - excelRange.Start.Row;
                    xlRange[cIx, rIx] = excel.GetValue(r, c);
                    tRange[cIx, rIx] = sheet.GetValue(r, c);
                }
            }

            var dir = @"c:\temp\hannes";
            if(!Directory.Exists(dir)) Directory.CreateDirectory(dir);
            var filePath = Path.Combine(dir, address.Replace(':', '-') + ".xlsx");
            if(File.Exists(filePath))
            {
                File.Delete(filePath);
            }

            var resultOffset = 8;
            using (var package = new ExcelPackage(filePath))
            {
                var resSheet = package.Workbook.Worksheets.Add("Result");
                for(var c = 0; c <= excelRange.End.Column - excelRange.Start.Column; c++)
                {
                    for(var r = 0; r <= (excelRange.End.Row - excelRange.Start.Row); r++)
                    {
                        resSheet.Cells[r + 1, c + 1].Value = xlRange[c, r];
                        resSheet.Cells[r + 1, c + 1 + resultOffset].Value = tRange[c, r];
                    }
                }
                package.Save();
            }
            
        }
    }
}

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
                sheet.ClearFormulaValues();
                sheet.Calculate();
                var sheet2 = package.Workbook.Worksheets.Add("Sheet5");
                for (var c = 1; c <= sheet.Dimension.End.Column; c++)
                {
                    for (var r = 1; r <= sheet.Dimension.End.Row; r++)
                    {
                        sheet2.Cells[r, c].Value = sheet.GetValue(r, c);
                    }
                }
                SaveWorkbook(@"LinestTestResults2.xlsx", package);
                //CompareRange(excelSheet, sheet, "O10:R14");


            }
        }

        [TestMethod]
        public void Test2()
        {
            using (var package = OpenTemplatePackage(@"LinestTest.xlsx"))
            {
                var oSheet = package.Workbook.Worksheets["Sheet1"];
                var sheet = package.Workbook.Worksheets.Add("EPPlus calc");
                // SET 1
                oSheet.Cells["B4:F9"].Copy(sheet.Cells["B4:F9"]);
                //sheet.Cells["B11"].Formula = "LINEST(B4:B9,C4:F9,TRUE,TRUE)";

                //SET 2
                oSheet.Cells["I4:L7"].Copy(sheet.Cells["I4:L7"]);
                //sheet.Cells["I9"].Formula = "LINEST(I4:I7,J4:L7,TRUE,TRUE)";
                //sheet.Cells["I15"].Formula = "LINEST(I4:I7,J4:L7,FALSE,TRUE)";

                // SET 3
                oSheet.Cells["O4:R8"].Copy(sheet.Cells["O4:R8"]);
                sheet.Cells["O10"].Formula = "LINEST(O4:O8,P4:R8,TRUE,TRUE)";
                //sheet.Cells["O16"].Formula = "LINEST(O4:O8,P4:R8,FALSE,TRUE)";

                // SET 4
                oSheet.Cells["B24:F25"].Copy(sheet.Cells["B24:F25"]);
                //sheet.Cells["B27"].Formula = "LINEST(B24:B25,C24:F25,TRUE,TRUE)";
                //sheet.Cells["B33"].Formula = "LINEST(B24:B25,C24:F25,FALSE,TRUE)";

                // SET 5
                oSheet.Cells["I24:M33"].Copy(sheet.Cells["I24:M33"]);
                //sheet.Cells["I35"].Formula = "LINEST(I24:I33,J24:M33,TRUE,TRUE)";
                //sheet.Cells["I41"].Formula = "LINEST(I24:I33,J24:M33,FALSE,TRUE)";

                // SET 6
                oSheet.Cells["P24:T33"].Copy(sheet.Cells["P24:T33"]);
                //sheet.Cells["P35"].Formula = "LINEST(P24:P33,Q24:T33,TRUE,TRUE)";
                //sheet.Cells["41"].Formula = "LINEST(P24:P33;Q24:T33;FALSE;TRUE)";

                // SET 7
                oSheet.Cells["B50:D54"].Copy(sheet.Cells["B50:D54"]);
                //sheet.Cells["B56"].Formula = "LINEST(B50:B54,C50:D54,TRUE,TRUE)";
                //sheet.Cells["B62"].Formula = "LINEST(B50:B54,C50:D54,FALSE,TRUE)";

                // SET 8
                oSheet.Cells["G50:I54"].Copy(sheet.Cells["G50:I54"]);
                //sheet.Cells["G56"].Formula = "LINEST(G50:G54,H50:I54,TRUE,TRUE)";
                //sheet.Cells["G62"].Formula = "LINEST(G50:G54,H50:I54,FALSE,TRUE)";
                //sheet.Cells[""].Formula = "";
                //sheet.Cells[""].Formula = "";
                //sheet.Cells[""].Formula = "";
                //sheet.Cells[""].Formula = "";
                //sheet.Cells[""].Formula = "";
                //sheet.Cells[""].Formula = "";
                //sheet.Cells["A1:O25"].UseImplicitItersection = false;
                sheet.Calculate();
                var p10 = sheet.Cells["P16"].Value;
                var sheet2 = package.Workbook.Worksheets.Add("Sheet5");
                for (var c = 1; c <= sheet.Dimension.End.Column; c++)
                {
                    for (var r = 1; r <= sheet.Dimension.End.Row; r++)
                    {
                        sheet2.Cells[r, c].Value = sheet.GetValue(r, c);
                    }
                }
                SaveWorkbook(@"LinestTestResults3.xlsx", package);
                //CompareRange(excelSheet, sheet, "O10:R14");


            }
        }

        [TestMethod]
        public void Test3()
        {
            using (var package = OpenTemplatePackage(@"LinestTest.xlsx"))
            {
                var wb = package.Workbook;
                var sheet = package.Workbook.Worksheets["Sheet1"];
                package.Workbook.Worksheets.Add("Sheet5");
                //SET 1
                PerformLinest(wb, sheet, "B11", "LINEST(B4:B9,C4:F9,TRUE,TRUE)", "B4:F9", "B11:F15");
                PerformLinest(wb, sheet, "B17", "LINEST(B4:B9,C4:F9,FALSE,TRUE)", "B4:F9", "B17:F21");
                //SET 2
                PerformLinest(wb, sheet, "I9", "LINEST(I4:I7,J4:L7,TRUE,TRUE)", "I4:L7", "I9:L13");
                PerformLinest(wb, sheet, "I15", "LINEST(I4:I7,J4:L7,FALSE,TRUE)", "I4:L7", "I15:L19");
                //SET 3
                PerformLinest(wb, sheet, "O10", "LINEST(O4:O8,P4:R8,TRUE,TRUE)", "O4:R8", "O10:R14");
                PerformLinest(wb, sheet, "O16", "LINEST(O4:O8,P4:R8,FALSE,TRUE)", "O4:R8", "O16:R20");
                //SET 4
                PerformLinest(wb, sheet, "B27", "LINEST(B24:B25,C24:F25,TRUE,TRUE)", "B24:F25", "B27:F31");
                PerformLinest(wb, sheet, "B33", "LINEST(B24:B25,C24:F25,FALSE,TRUE)", "B24:F25", "B33:F37");
                //SET 5
                PerformLinest(wb, sheet, "I35", "LINEST(I24:I33,J24:M33,TRUE,TRUE)", "I24:M33", "I35:M39");
                PerformLinest(wb, sheet, "I41", "LINEST(I24:I33,J24:M33,FALSE,TRUE)", "I24:M33", "I41:M45");
                //SET 6
                PerformLinest(wb, sheet, "P35", "LINEST(P24:P33,Q24:T33,TRUE,TRUE)", "P24:T33", "P35:T39");
                PerformLinest(wb, sheet, "P41", "LINEST(P24:P33,Q24:T33,FALSE,TRUE)", "P24:T33", "P41:T45");
                //SET 7
                PerformLinest(wb, sheet, "B56", "LINEST(B50:B54,C50:D54,TRUE,TRUE)", "B50:D54", "B56:D60");
                PerformLinest(wb, sheet, "B62", "LINEST(B50:B54,C50:D54,FALSE,TRUE)", "B50:D54", "B62:D66");
                //SET 8
                PerformLinest(wb, sheet, "G56", "LINEST(G50:G54,H50:I54,TRUE,TRUE)", "G50:I54", "G56:I60");
                PerformLinest(wb, sheet, "G62", "LINEST(G50:G54,H50:I54,FALSE,TRUE)", "G50:I54", "G62:I66");
                //SET 9
                PerformLinest(wb, sheet, "K56", "LINEST(K50:K54,L50:N54,TRUE,TRUE)", "K50:N54", "K56:N60");
                PerformLinest(wb, sheet, "K62", "LINEST(K50:K54,L50:N54,FALSE,TRUE)", "K50:N54", "K62:N66");
                //SET 10
                PerformLinest(wb, sheet, "P56", "LINEST(P50:S50,P51:S53,TRUE,TRUE)", "P50:S53", "P56:S60");
                PerformLinest(wb, sheet, "P62", "LINEST(P50:S50,P51:S53,FALSE,TRUE)", "P50:S53", "P62:S66");
                //SET 11
                //PerformLinest(wb, sheet, "N73", "LINEST(B72:B1071,C72:L1071,TRUE,TRUE)", "B72:L1071", "N73:X77");
                //PerformLinest(wb, sheet, "N80", "LINEST(B72:B1071,C72:L1071,FALSE,TRUE)", "B72:L1071", "N80:X84");
                //SET 12
                PerformLinest(wb, sheet, "N97", "LINEST(N88:N95,O88:T95,TRUE,TRUE)", "N88:T95", "N97:T101");
                PerformLinest(wb, sheet, "N103", "LINEST(N88:N95,O88:T95,FALSE,TRUE)", "N88:T95", "N103:T107");
                SaveWorkbook(@"LinestTestResults4.xlsx", package);
            }
        }

        [TestMethod]
        public void TestOne()
        {
            using (var package = OpenTemplatePackage(@"LinestTest.xlsx"))
            {
                var wb = package.Workbook;
                var sheet = package.Workbook.Worksheets["Sheet1"];
                package.Workbook.Worksheets.Add("Sheet5");
                //Insert debug line here
                PerformLinest(wb, sheet, "B11", "LINEST(B4:B9,C4:F9,TRUE,TRUE)", "B4:F9", "B11:F15");
            }
        }

        private void PerformLinest(ExcelWorkbook workbook, ExcelWorksheet source, string linestAddress, string linestFormula, string inputRange, string outputRange)
        {
            var tmpSheet = workbook.Worksheets.Add("tmp");
            source.Cells[inputRange].Copy(tmpSheet.Cells[inputRange]);
            tmpSheet.Cells[linestAddress].Formula = linestFormula;
            tmpSheet.Calculate();
            var targetSheet = workbook.Worksheets["Sheet5"];
            tmpSheet.Cells[outputRange].Copy(targetSheet.Cells[outputRange], ExcelRangeCopyOptionFlags.ExcludeFormulas);
            workbook.Worksheets.Delete(tmpSheet);

        }

        private void CompareRange(ExcelWorksheet excel, ExcelWorksheet sheet, string address)
        {
            var excelRange = excel.Cells[address];
            var xlRange = new object[excelRange.End.Column - excelRange.Start.Column + 1, excelRange.End.Row - excelRange.Start.Row + 1];
            var tRange = new object[excelRange.End.Column - excelRange.Start.Column + 1, excelRange.End.Row - excelRange.Start.Row + 1];
            for (var c = excelRange.Start.Column; c <= excelRange.End.Column; c++)
            {
                for (var r = excelRange.Start.Row; r <= excelRange.End.Row; r++)
                {
                    var cIx = c - excelRange.Start.Column;
                    var rIx = r - excelRange.Start.Row;
                    xlRange[cIx, rIx] = excel.GetValue(r, c);
                    tRange[cIx, rIx] = sheet.GetValue(r, c);
                }
            }

            var dir = @"c:\temp\hannes";
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
            var filePath = Path.Combine(dir, address.Replace(':', '-') + ".xlsx");
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }

            var resultOffset = 8;
            using (var package = new ExcelPackage(filePath))
            {
                var resSheet = package.Workbook.Worksheets.Add("Result");
                for (var c = 0; c <= excelRange.End.Column - excelRange.Start.Column; c++)
                {
                    for (var r = 0; r <= (excelRange.End.Row - excelRange.Start.Row); r++)
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


using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using System;
using System.Diagnostics;
using System.IO;

namespace EPPlusTest
{

    public class TestClass
    {
        public string name;
        int age;

        public TestClass() { }

        public TestClass(string name) => this.name = name;

        public TestClass(string name, int age)
        {
            this.name = name;
            this.age = age;
        }
    }

    [TestClass]
    public class Test635 : TestBase
    {
        [TestMethod]
        public void AnyValidationTest()
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("test");
                var customValidation = ws.DataValidations.AddCustomValidation("A1:A5");

                customValidation.ShowErrorMessage = true;
                customValidation.Formula.ExcelFormula = "ISNUMBER(A1)";
                customValidation.ErrorTitle = "Input invalid in Data Validation";
                customValidation.Error = $"Value in cell must be number!!";
                customValidation.ErrorStyle = ExcelDataValidationWarningStyle.stop;

                package.SaveAs("C:/temp/TempExample.xlsx");
            }
        }

        [TestMethod]
        public void CanReadAndSaveExtLstPackageSafely()
        {
            using (ExcelPackage package = OpenTemplatePackage("ExtLstDataValidationValidation.xlsx"))
            {
                SaveAndCleanup(package);

                ExcelPackage p = new ExcelPackage("C:\\epplusTest\\Testoutput\\ExtLstDataValidationValidation.xlsx");

                Assert.IsTrue(p.Workbook.Worksheets[0].DataValidations.Count > 0);
            }
        }

        [TestMethod]
        public void SpeedTestDataValidations()
        {
            Debug.WriteLine($"----Loading Package------");
            var stopWatch = Stopwatch.StartNew();
            using (var package = new ExcelPackage())
            {
                Debug.WriteLine($"{stopWatch.Elapsed}");
                stopWatch.Stop();

                var sheet = package.Workbook.Worksheets.Add("Example");

                Debug.WriteLine($"----Adding Validation to Package------");
                stopWatch = Stopwatch.StartNew();
                for (int i = 0; i < 100000; i++)
                {
                    var validation = sheet.DataValidations.AddDateTimeValidation("AA" + (i + 1).ToString());

                    validation.Formula.ExcelFormula = "2022/01/05";
                    validation.Formula2.ExcelFormula = "2022/01/06";
                }
                Debug.WriteLine($"{stopWatch.Elapsed}");
                Assert.IsTrue(stopWatch.Elapsed.Seconds < 40);
                stopWatch.Stop();

                Debug.WriteLine($"----Saving Package------");
                stopWatch = Stopwatch.StartNew();
                SaveAndCleanup(package);

                Debug.WriteLine($"{stopWatch.Elapsed}");
                stopWatch.Stop();
            }
        }

        public void DateTimeTest()
        {
            FileInfo testFile = new FileInfo("C:\\Users\\OssianEdström\\Documents\\DateTimeEx.xlsx");

            ExcelPackage package = new ExcelPackage(testFile);
            var workSheet = package.Workbook.Worksheets[0];
            var validation = workSheet.DataValidations.AddDateTimeValidation("A2");

            validation.Formula.Value = DateTime.Now;
            validation.Formula2.Value = new DateTime(2022, 07, 03, 15, 00, 00);

            string path = @"C:\Users\OssianEdström\Documents\testNew.xlsx";
            Stream stream = File.Create(path);
            package.SaveAs(stream);
            stream.Close();

            SaveAndCleanup(package);
        }

        public void TimeTest()
        {
            FileInfo testFile = new FileInfo("C:\\Users\\OssianEdström\\Documents\\TimeExample.xlsx");

            ExcelPackage package = new ExcelPackage(testFile);
            var workSheet = package.Workbook.Worksheets[0];
            var validation = workSheet.DataValidations.AddTimeValidation("A2");

            validation.Formula.Value.Hour = 9;
            validation.Formula2.Value.Hour = 17;

            string path = @"C:\Users\OssianEdström\Documents\testNew.xlsx";
            Stream stream = File.Create(path);
            package.SaveAs(stream);
            stream.Close();

            SaveAndCleanup(package);
        }

        public void ListTest()
        {
            FileInfo testFile = new FileInfo("C:\\Users\\OssianEdström\\Documents\\ListEx.xlsx");

            ExcelPackage package = new ExcelPackage(testFile);

            var workSheetArr = package.Workbook.Worksheets;

            var workSheet = workSheetArr[0];

            var validation = workSheet.DataValidations.AddListValidation("D3");

            validation.ErrorStyle = ExcelDataValidationWarningStyle.warning;

            validation.Formula.Values.Add("1");
            validation.Formula.Values.Add("2");
            validation.Formula.Values.Add("3");

            validation.Formula.Values.Add("15");

            string path = @"C:\Users\OssianEdström\Documents\testNew.xlsx";
            Stream stream = File.Create(path);
            package.SaveAs(stream);
            stream.Close();

            SaveAndCleanup(package);
        }

        public void DecimalTest()
        {
            FileInfo testFile = new FileInfo("C:\\Users\\OssianEdström\\Documents\\DecimalEx.xlsx");

            ExcelPackage package = new ExcelPackage(testFile);

            var workSheetArr = package.Workbook.Worksheets;

            var workSheet = workSheetArr[0];

            var validation = workSheet.DataValidations.AddDecimalValidation("B3");

            validation.ErrorStyle = ExcelDataValidationWarningStyle.warning;

            validation.Formula.Value = 1.5;
            validation.Formula2.Value = 5.5;

            string path = @"C:\Users\OssianEdström\Documents\testNew.xlsx";
            Stream stream = File.Create(path);
            package.SaveAs(stream);
            stream.Close();

            SaveAndCleanup(package);
        }


        public void ValidateExtLstANDLocalDataValidation()
        {
            using (var P = new ExcelPackage(@"C:\epplusTest\TestOutput\extLstTest.xlsx"))
            {
                ExcelWorksheet sheet = P.Workbook.Worksheets.Add("NewSheet");
                ExcelWorksheet sheet2 = P.Workbook.Worksheets.Add("ExtSheet");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 5;


                var validationLocal = sheet2.DataValidations.AddIntegerValidation("B1");

                validationLocal.ErrorStyle = ExcelDataValidationWarningStyle.stop;
                validationLocal.PromptTitle = "Enter a integer value here";
                validationLocal.Prompt = "Value should be between 1 and 5";
                validationLocal.ShowInputMessage = true;
                validationLocal.ErrorTitle = "An invalid value was entered";
                validationLocal.Error = "Value must be between 1 and 5";
                validationLocal.ShowErrorMessage = true;
                validationLocal.Operator = ExcelDataValidationOperator.between;
                validationLocal.Formula.Value = 6;
                validationLocal.Formula2.ExcelFormula = "=ExtSheet!A2";

                var validation = sheet2.DataValidations.AddIntegerValidation("A1");

                validation.ErrorStyle = ExcelDataValidationWarningStyle.stop;
                validation.PromptTitle = "Enter a integer value here";
                validation.Prompt = "Value should be between 1 and 5";
                validation.ShowInputMessage = true;
                validation.ErrorTitle = "An invalid value was entered";
                validation.Error = "Value must be between 1 and 5";
                validation.ShowErrorMessage = true;
                validation.Operator = ExcelDataValidationOperator.between;
                validation.Formula.ExcelFormula = "NewSheet!A1";
                validation.Formula2.ExcelFormula = "NewSheet!A2";

                SaveAndCleanup(P);

                //TODO: Assert that sheets are valid xmls here.
            }
        }
    }
}

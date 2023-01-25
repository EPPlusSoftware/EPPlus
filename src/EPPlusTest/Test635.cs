using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using System;
using System.Collections.Generic;
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
        public void Stuff()
        {
            FileInfo testFile = new FileInfo("C:\\Users\\OssianEdström\\Documents\\ListEx.xlsx");

            ExcelPackage package = new ExcelPackage(testFile);

            var workSheetArr = package.Workbook.Worksheets;

            var workSheet = workSheetArr[0];

            var validation = workSheet.DataValidations.AddListValidation("D3");

            validation.ErrorStyle = ExcelDataValidationWarningStyle.warning;

            validation.Formula.ExcelFormula = "$A$1:$A$7";

            // workSheet.DataValidations.Remove(workSheet.DataValidations[0]);

            //var validation = workSheet.DataValidations.AddIntegerValidation("B3");

            //validation.PromptTitle = "This is an IntegerTest TITLE";
            //validation.Prompt = "This is an IntegerTest";
            //validation.Formula.Value = 1;
            //validation.Formula2.Value = 5;
            //validation.ShowInputMessage = true;

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

        [TestMethod]
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

                // Alternatively:
                //var validation = sheet.Cells["A1:A2"].DataValidation;

                //validation.AddAnyDataValidation();
                //var validation2 = validation.AddDateTimeDataValidation();

                //validation2.ErrorStyle = ExcelDataValidationWarningStyle.stop;

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

        private void SpeedTest(Func<TestClass> function)
        {
            var testList = new List<TestClass>();
            var stopWatch = Stopwatch.StartNew();

            for (int i = 0; i < 100000; i++)
            {
                TestClass test = function();

                testList.Add(test);
            }

            Debug.WriteLine($"{stopWatch.Elapsed}");
            stopWatch.Stop();
            testList.Clear();
        }


        [TestMethod]
        public void SpeedTestNewNoArgs()
        {
            Debug.WriteLine("---New Start---");
            SpeedTest(delegate { return new TestClass(); });
        }

        [TestMethod]
        public void SpeedTestActivatorNoArgs()
        {
            Debug.WriteLine("---Activator Start---");
            SpeedTest(Activator.CreateInstance<TestClass>);
        }

        [TestMethod]
        public void SpeedTestCompiledLambdaExpressionNoArgsNoWarmUp()
        {
            Debug.WriteLine("---LambdaNoWarmUp Start---");
            SpeedTest(New<TestClass>.Instance);
        }

        [TestMethod]
        public void SpeedTestCompiledLambdaExpressionNoArgs()
        {
            Debug.WriteLine("---Lambda Start---");
            //Note: This is neccesary in order to "warm-up" the instantiation;
            New<TestClass>.Instance();
            SpeedTest(New<TestClass>.Instance);

        }

        [TestMethod]
        public void SpeedTestNewArgs()
        {
            Debug.WriteLine("---NewArgs Start---");
            SpeedTest(delegate { return new TestClass("TestName", 5); });
        }

        [TestMethod]
        public void SpeedTestActivatorArgs()
        {
            Debug.WriteLine("---ActivatorArgs Start---");
            SpeedTest(
                delegate
                {
                    return (TestClass)Activator.CreateInstance(typeof(TestClass), "TestName", 5);
                }
                );
        }

        //[TestMethod]
        //public void SpeedTestCompiledLambdaArgs()
        //{
        //    Debug.WriteLine("---LambdaArgs Start---");
        //    SpeedTest(delegate { return (TestClass)typeof(TestClass).CreateInstance("TestName", 5); });
        //}
    }
}

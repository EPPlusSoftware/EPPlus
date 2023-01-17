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
            FileInfo testFile = new FileInfo("C:\\Users\\OssianEdström\\Documents\\BookInitial.xlsx");

            ExcelPackage package = new ExcelPackage(testFile);

            var wb = package.Workbook.Worksheets;

            //using (var P = OpenPackage(testFile))
            //{

            //}
            //{
            //    //ExcelWorksheet worksheet = P.Workbook.Worksheets[1];
            //    ////P.Workbook.Worksheets.Add("NewSheet");
            //    ////AddIntegerValidation(P);
            //    //string path = @"C:\test2.xlsx";
            //    //Stream stream = File.Create(path);
            //    //P.SaveAs(stream);
            //    //stream.Close();

            //    ////SaveAndCleanup(P);
            //}

        }

        private static void AddIntegerValidation(ExcelPackage package)
        {
            var sheet = package.Workbook.Worksheets.Add("integer");
            //add a validation and set values
            var validation = sheet.DataValidations.AddIntegerValidation("A1:A2");

            // Alternatively:
            //var validation = sheet.Cells["A1:A2"].DataValidation;

            //validation.AddAnyDataValidation();
            //var validation2 = validation.AddDateTimeDataValidation();

            //validation2.ErrorStyle = ExcelDataValidationWarningStyle.stop;


            validation.ErrorStyle = ExcelDataValidationWarningStyle.stop;
            validation.PromptTitle = "Enter a integer value here";
            validation.Prompt = "Value should be between 1 and 5";
            validation.ShowInputMessage = true;
            validation.ErrorTitle = "An invalid value was entered";
            validation.Error = "Value must be between 1 and 5";
            validation.ShowErrorMessage = true;
            validation.Operator = ExcelDataValidationOperator.between;
            validation.Formula.Value = 1;
            validation.Formula2.Value = 5;
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

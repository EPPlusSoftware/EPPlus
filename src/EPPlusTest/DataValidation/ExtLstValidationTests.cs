using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.DataValidation;
using System.Drawing;
using System.IO;

namespace EPPlusTest.DataValidation
{
    [TestClass]
    public class ExtLstValidationTests : TestBase
    {
        [TestMethod, Ignore]
        public void AddValidationWithFormulaOnOtherWorksheetShouldReturnExt()
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("test");
                var sheet2 = package.Workbook.Worksheets.Add("test2");
                var val = sheet1.DataValidations.AddListValidation("A1");
                val.Formula.ExcelFormula = "test2!A1:A2";
                Assert.IsInstanceOfType(val, typeof(ExcelDataValidationList));
            }
        }

        [TestMethod]
        public void CanReadWriteSimpleExtLst()
        {
            using (ExcelPackage package = new ExcelPackage(new MemoryStream()))
            {
                var ws1 = package.Workbook.Worksheets.Add("ExtTest");
                var ws2 = package.Workbook.Worksheets.Add("ExternalAdresses");

                var validation = ws1.DataValidations.AddIntegerValidation("A1");
                validation.Operator = ExcelDataValidationOperator.equal;
                ws2.Cells["A1"].Value = 5;

                validation.Formula.ExcelFormula = "sheet2!A1";

                Assert.AreEqual(((ExcelDataValidationInt)validation).InternalValidationType, InternalValidationType.ExtLst);

                var stream = new MemoryStream();
                package.SaveAs(stream);

                ExcelPackage package2 = new ExcelPackage(stream);

                var readingValidation = package2.Workbook.Worksheets[0].DataValidations[0];

                Assert.AreEqual("sheet2!A1", readingValidation.As.IntegerValidation.Formula.ExcelFormula);
                Assert.AreEqual(((ExcelDataValidationInt)readingValidation).InternalValidationType, InternalValidationType.ExtLst);
            }
        }

        [TestMethod]
        public void EnsureIsNotExtLstWhenRegularReadWrite()
        {
            using (ExcelPackage package = new ExcelPackage(new MemoryStream()))
            {
                var ws1 = package.Workbook.Worksheets.Add("ExtTest");
                var ws2 = package.Workbook.Worksheets.Add("ExternalAdresses");

                var validation = ws1.DataValidations.AddIntegerValidation("A1");
                validation.Operator = ExcelDataValidationOperator.equal;

                validation.Formula.ExcelFormula = "IF(A2=\"red\")";

                Assert.AreNotEqual(((ExcelDataValidationInt)validation).InternalValidationType, InternalValidationType.ExtLst);

                var stream = new MemoryStream();
                package.SaveAs(stream);

                ExcelPackage package2 = new ExcelPackage(stream);

                var readingValidation = package2.Workbook.Worksheets[0].DataValidations[0];

                Assert.AreEqual("IF(A2=\"red\")", readingValidation.As.IntegerValidation.Formula.ExcelFormula);
                Assert.AreNotEqual(((ExcelDataValidationInt)readingValidation).InternalValidationType, InternalValidationType.ExtLst);
            }
        }

        [TestMethod]
        public void ReadAndSaveExtLstPackage_ShouldNotThrow()
        {
            using (ExcelPackage package = OpenTemplatePackage("ExtLstDataValidationValidation.xlsx"))
            {
                var memoryStream = new MemoryStream();
                package.SaveAs(memoryStream);
                ExcelPackage p = new ExcelPackage(memoryStream);

                Assert.IsTrue(p.Workbook.Worksheets[0].DataValidations.Count > 0);
            }
        }

        ExcelPackage MakePackageWithExtLstIntValidation()
        {
            var package = new ExcelPackage(new MemoryStream());

            package.Workbook.Worksheets.Add("extValidations");
            package.Workbook.Worksheets.Add("extValidationTargets");

            var validation = package.Workbook.Worksheets[0].DataValidations.AddIntegerValidation("A1");
            validation.Operator = ExcelDataValidationOperator.equal;

            validation.Formula.ExcelFormula = "sheet2!A1";

            return package;
        }

        [TestMethod]
        public void ReadWriteWorksWithOneValidation()
        {
            var creationPackage = MakePackageWithExtLstIntValidation();

            var stream = new MemoryStream();
            creationPackage.SaveAs(stream);

            var readingPackage = new ExcelPackage(stream);

            var validation = readingPackage.Workbook.Worksheets[0].DataValidations[0];
            Assert.AreEqual(ExcelDataValidationOperator.equal, validation.Operator);
            Assert.AreEqual("sheet2!A1", validation.As.IntegerValidation.Formula.ExcelFormula);
            Assert.AreEqual(InternalValidationType.ExtLst, validation.InternalValidationType);
        }

        [TestMethod]
        public void WorksWithManyValidations()
        {
            var creationPackage = MakePackageWithExtLstIntValidation();

            var decimalValidation = creationPackage.Workbook.Worksheets[0].DataValidations.AddDecimalValidation("B1");
            decimalValidation.Operator = ExcelDataValidationOperator.between;

            decimalValidation.Formula.ExcelFormula = "sheet2!B1";
            decimalValidation.Formula2.ExcelFormula = "1.5";

            var timeValidation = creationPackage.Workbook.Worksheets[0].DataValidations.AddTimeValidation("C1");
            timeValidation.Operator = ExcelDataValidationOperator.notBetween;

            timeValidation.Formula.ExcelFormula = "sheet2!C1";
            timeValidation.Formula2.ExcelFormula = "14:00";

            var listValidation = creationPackage.Workbook.Worksheets[0].DataValidations.AddListValidation("D1");

            listValidation.Formula.ExcelFormula = "sheet2!A1, sheet2!B1, sheet2!C1";

            var textLength = creationPackage.Workbook.Worksheets[0].DataValidations.AddTextLengthValidation("E1");

            textLength.Operator = ExcelDataValidationOperator.lessThan;
            textLength.Formula.ExcelFormula = "sheet2!D1";

            var stream = new MemoryStream();
            creationPackage.SaveAs(stream);

            var readingPackage = new ExcelPackage(stream);

            var validations = readingPackage.Workbook.Worksheets[0].DataValidations;

            Assert.AreEqual(ExcelDataValidationOperator.equal, validations[0].Operator);
            Assert.AreEqual("sheet2!A1", validations[0].As.IntegerValidation.Formula.ExcelFormula);
            Assert.AreEqual(InternalValidationType.ExtLst, validations[0].InternalValidationType);

            Assert.AreEqual(ExcelDataValidationOperator.between, validations[1].Operator);
            Assert.AreEqual("sheet2!B1", validations[1].As.DecimalValidation.Formula.ExcelFormula);
            Assert.AreEqual(1.5, validations[1].As.DecimalValidation.Formula2.Value);
            Assert.AreEqual(InternalValidationType.ExtLst, validations[1].InternalValidationType);

            Assert.AreEqual(ExcelDataValidationOperator.notBetween, validations[2].Operator);
            Assert.AreEqual("sheet2!C1", validations[2].As.TimeValidation.Formula.ExcelFormula);
            Assert.AreEqual("14:00", validations[2].As.TimeValidation.Formula2.ExcelFormula);
            Assert.AreEqual(InternalValidationType.ExtLst, validations[2].InternalValidationType);

            Assert.AreEqual("sheet2!A1, sheet2!B1, sheet2!C1", validations[3].As.ListValidation.Formula.ExcelFormula);
            Assert.AreEqual(InternalValidationType.ExtLst, validations[3].InternalValidationType);

            Assert.AreEqual("sheet2!D1", validations[4].As.IntegerValidation.Formula.ExcelFormula);
            Assert.AreEqual(InternalValidationType.ExtLst, validations[4].InternalValidationType);
        }

        internal void AddDataValidations(ref ExcelWorksheet ws, bool isExtLst = false, string extSheetName = "", bool many = false)
        {
            if (isExtLst)
            {
                var intValidation = ws.DataValidations.AddIntegerValidation("A1");
                intValidation.Operator = ExcelDataValidationOperator.equal;
                intValidation.Formula.ExcelFormula = extSheetName + "!A1";
            }
            else
            {
                var intValidation = ws.DataValidations.AddIntegerValidation("A2");
                intValidation.Formula.Value = 1;
                intValidation.Formula2.Value = 3;
            }

            if (many)
            {

                if (isExtLst)
                {
                    var timeValidation = ws.DataValidations.AddTimeValidation("B1");
                    timeValidation.Operator = ExcelDataValidationOperator.between;

                    timeValidation.Formula.ExcelFormula = extSheetName + "!B1";
                    timeValidation.Formula2.ExcelFormula = extSheetName + "!B2";


                }
                else
                {
                    var timeValidation = ws.DataValidations.AddTimeValidation("B2");
                    timeValidation.Operator = ExcelDataValidationOperator.between;

                    timeValidation.Formula.ExcelFormula = "B1";
                    timeValidation.Formula.ExcelFormula = "B2";
                }
            }
        }

        //Ensures no save or load errors
        internal void SaveAndLoadAndSave(in ExcelPackage pck)
        {
            var file = pck.File;

            var stream = new MemoryStream();
            pck.SaveAs(stream);

            var loadedPackage = new ExcelPackage(stream);

            loadedPackage.File = file;

            SaveAndCleanup(loadedPackage);
        }

        [TestMethod]
        public void LocalDataValidationsShouldWorkWithExtLstValidation()
        {
            using (var pck = OpenPackage("DataValidationLocalExtLst.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("extLstTest");
                var extSheet = pck.Workbook.Worksheets.Add("extAddressSheet");

                AddDataValidations(ref ws, false);
                AddDataValidations(ref ws, true, "extAddressSheet");

                SaveAndLoadAndSave(pck);
            }
        }

        [TestMethod]
        public void LocalDataValidationsShouldWorkWithManyExtLstValidations()
        {
            using (var pck = OpenPackage("DataValidationLocalExtLstMany.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("extLstTest");
                var extSheet = pck.Workbook.Worksheets.Add("extAddressSheet");

                AddDataValidations(ref ws, false);
                AddDataValidations(ref ws, true, "extAddressSheet", true);

                SaveAndLoadAndSave(pck);
            }
        }

        [TestMethod]
        public void ManyLocalDataValidationsShouldWorkWithSingularExtLstValidations()
        {
            using (var pck = OpenPackage("DataValidationLocalManyExtLst.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("extLstTest");
                var extSheet = pck.Workbook.Worksheets.Add("extAddressSheet");

                AddDataValidations(ref ws, false, "", true);
                AddDataValidations(ref ws, true, "extAddressSheet");

                SaveAndLoadAndSave(pck);
            }

        }

        [TestMethod]
        public void ManyLocalDataValidationsShouldWorkWithManyExtLstConditionalFormattings()
        {
            using (var pck = OpenPackage("DataValidationLocalManyExtLstMany.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("extLstTest");
                var extSheet = pck.Workbook.Worksheets.Add("extAddressSheet");

                AddDataValidations(ref ws, false, "", true);
                AddDataValidations(ref ws, true, "extAddressSheet", true);

                SaveAndLoadAndSave(pck);
            }
        }

        [TestMethod]
        public void LocalMultipleAddress()
        {
            using (var pck = OpenPackage("DataValidationLocalSeperatedAddress.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("localAddressTest");

                var validation = ws.DataValidations.AddDecimalValidation("A1:A5 C5:C15 D13");

                validation.Formula.Value = 5;
                validation.Formula2.Value = 10.5;

                SaveAndLoadAndSave(pck);
            }
        }

        [TestMethod]
        public void ExtLstMultipleAddress()
        {
            using (var pck = OpenPackage("DataValidationExtLstSeperatedAddress.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("extLstAddressTest");
                var ws2 = pck.Workbook.Worksheets.Add("external");


                var validation = ws.DataValidations.AddDecimalValidation("A1:A5 C5:C15 D13");

                validation.Formula.ExcelFormula = "external!A1";
                validation.Formula2.Value = 10.5;

                SaveAndLoadAndSave(pck);
            }
        }

        [TestMethod]
        public void ExtLstAndLocalMultipleAddressShouldWork()
        {
            using (var pck = OpenPackage("DataValidationLocalExtSeperatedAddress.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("extLstAddressTest");
                var ws2 = pck.Workbook.Worksheets.Add("external");


                var extValidation = ws.DataValidations.AddDecimalValidation("A1:A5 C5:C15 D13");

                extValidation.Formula.ExcelFormula = "external!A1";
                extValidation.Formula2.Value = 10.5;

                var localValidation = ws.DataValidations.AddDecimalValidation("E1:E5 F5:F15 G13");

                localValidation.Formula.Value = 5.5;
                localValidation.Formula2.Value = 25.75;

                SaveAndLoadAndSave(pck);
            }
        }

        [TestMethod]
        public void DataValidationExtLstShouldWorkWithConditionalFormatting()
        {
            var creationPackage = MakePackageWithExtLstIntValidation();

            creationPackage.Workbook.Worksheets[0]
                .ConditionalFormatting.AddDatabar(new ExcelAddress("A1:A5"), Color.DarkBlue);

            var stream = new MemoryStream();
            creationPackage.SaveAs(stream);

            var readingPackage = new ExcelPackage(stream);
            var ws = readingPackage.Workbook.Worksheets[0];
            var validation = ws.DataValidations[0];

            Assert.AreEqual(eExcelConditionalFormattingRuleType.DataBar, ws.ConditionalFormatting[0].Type);

            Assert.AreEqual(ExcelDataValidationOperator.equal, validation.Operator);
            Assert.AreEqual("sheet2!A1", validation.As.IntegerValidation.Formula.ExcelFormula);
            Assert.AreEqual(InternalValidationType.ExtLst, validation.InternalValidationType);
        }

        [TestMethod]
        public void DataValidationExtLstShouldWorkWithConditionalFormattingMultiple()
        {
            var creationPackage = MakePackageWithExtLstIntValidation();

            var decimalValidation = creationPackage.Workbook.Worksheets[0].DataValidations.AddDecimalValidation("B1");
            decimalValidation.Operator= ExcelDataValidationOperator.between;

            decimalValidation.Formula.ExcelFormula = "sheet2!B1";
            decimalValidation.Formula2.ExcelFormula = "1.5";

            creationPackage.Workbook.Worksheets[0]
                .ConditionalFormatting.AddDatabar(new ExcelAddress("A1:A5"), Color.DarkBlue);
            creationPackage.Workbook.Worksheets[0]
                .ConditionalFormatting.AddDatabar(new ExcelAddress("A1:A5"), Color.Red);

            var stream = new MemoryStream();
            creationPackage.SaveAs(stream);

            var readingPackage = new ExcelPackage(stream);
            var ws = readingPackage.Workbook.Worksheets[0];
            var validations = ws.DataValidations;

            Assert.AreEqual(eExcelConditionalFormattingRuleType.DataBar, ws.ConditionalFormatting[0].Type);
            Assert.AreEqual(eExcelConditionalFormattingRuleType.DataBar, ws.ConditionalFormatting[1].Type);


            Assert.AreEqual(ExcelDataValidationOperator.equal, validations[0].Operator);
            Assert.AreEqual("sheet2!A1", validations[0].As.IntegerValidation.Formula.ExcelFormula);
            Assert.AreEqual(InternalValidationType.ExtLst, validations[0].InternalValidationType);

            Assert.AreEqual(ExcelDataValidationOperator.between, validations[1].Operator);
            Assert.AreEqual("sheet2!B1", validations[1].As.DecimalValidation.Formula.ExcelFormula);
            Assert.AreEqual(1.5, validations[1].As.DecimalValidation.Formula2.Value);
            Assert.AreEqual(InternalValidationType.ExtLst, validations[1].InternalValidationType);
        }

        [TestMethod]
        public void DataValidationExtLstShouldWorkWithSparklines()
        {
            var creationPackage = MakePackageWithExtLstIntValidation();

            creationPackage.Workbook.Worksheets[0]
                .SparklineGroups.Add(OfficeOpenXml.Sparkline.eSparklineType.Line, new ExcelAddress("A1:A5"), new ExcelAddress("B1:B5"));

            var stream = new MemoryStream();
            creationPackage.SaveAs(stream);

            var readingPackage = new ExcelPackage(stream);
            var ws = readingPackage.Workbook.Worksheets[0];
            var validation = ws.DataValidations[0];

            Assert.AreEqual(OfficeOpenXml.Sparkline.eSparklineType.Line, ws.SparklineGroups[0].Type);

            Assert.AreEqual(ExcelDataValidationOperator.equal, validation.Operator);
            Assert.AreEqual("sheet2!A1", validation.As.IntegerValidation.Formula.ExcelFormula);
            Assert.AreEqual(InternalValidationType.ExtLst, validation.InternalValidationType);
        }

        [TestMethod]
        public void DataValidationExtLstShouldWorkWithSparklineMultiple()
        {
            var creationPackage = MakePackageWithExtLstIntValidation();

            var decimalValidation = creationPackage.Workbook.Worksheets[0].DataValidations.AddDecimalValidation("B1");
            decimalValidation.Operator = ExcelDataValidationOperator.between;

            decimalValidation.Formula.ExcelFormula = "sheet2!B1";
            decimalValidation.Formula2.ExcelFormula = "1.5";

            creationPackage.Workbook.Worksheets[0]
                .SparklineGroups.Add(OfficeOpenXml.Sparkline.eSparklineType.Line, new ExcelAddress("A1:A5"), new ExcelAddress("B1:B5"));
            creationPackage.Workbook.Worksheets[0]
                .SparklineGroups.Add(OfficeOpenXml.Sparkline.eSparklineType.Column, new ExcelAddress("C1:C5"), new ExcelAddress("D1:D5"));

            var stream = new MemoryStream();
            creationPackage.SaveAs(stream);

            var readingPackage = new ExcelPackage(stream);
            var ws = readingPackage.Workbook.Worksheets[0];
            var validations = ws.DataValidations;

            Assert.AreEqual(OfficeOpenXml.Sparkline.eSparklineType.Line, ws.SparklineGroups[0].Type);
            Assert.AreEqual(OfficeOpenXml.Sparkline.eSparklineType.Column, ws.SparklineGroups[1].Type);

            Assert.AreEqual(ExcelDataValidationOperator.equal, validations[0].Operator);
            Assert.AreEqual("sheet2!A1", validations[0].As.IntegerValidation.Formula.ExcelFormula);
            Assert.AreEqual(InternalValidationType.ExtLst, validations[0].InternalValidationType);

            Assert.AreEqual(ExcelDataValidationOperator.between, validations[1].Operator);
            Assert.AreEqual("sheet2!B1", validations[1].As.DecimalValidation.Formula.ExcelFormula);
            Assert.AreEqual(1.5, validations[1].As.DecimalValidation.Formula2.Value);
            Assert.AreEqual(InternalValidationType.ExtLst, validations[1].InternalValidationType);
        }

        [TestMethod]
        public void DataValidationExtLstShouldWorkWithConditionalFormattingANDSparklineSingle()
        {
            var creationPackage = MakePackageWithExtLstIntValidation();

            creationPackage.Workbook.Worksheets[0]
                .SparklineGroups.Add(OfficeOpenXml.Sparkline.eSparklineType.Line, new ExcelAddress("A1:A5"), new ExcelAddress("B1:B5"));
            creationPackage.Workbook.Worksheets[0]
                .ConditionalFormatting.AddDatabar(new ExcelAddress("A1:A5"), Color.DarkBlue);

            var stream = new MemoryStream();
            creationPackage.SaveAs(stream);

            var readingPackage = new ExcelPackage(stream);
            var ws = readingPackage.Workbook.Worksheets[0];
            var validation = ws.DataValidations[0];

            Assert.AreEqual(OfficeOpenXml.Sparkline.eSparklineType.Line, ws.SparklineGroups[0].Type);
            Assert.AreEqual(eExcelConditionalFormattingRuleType.DataBar, ws.ConditionalFormatting[0].Type);

            Assert.AreEqual(ExcelDataValidationOperator.equal, validation.Operator);
            Assert.AreEqual("sheet2!A1", validation.As.IntegerValidation.Formula.ExcelFormula);
            Assert.AreEqual(InternalValidationType.ExtLst, validation.InternalValidationType);
        }

        [TestMethod]
        public void DataValidationExtLstShouldWorkWithConditionalFormattingANDSparklineMultiple()
        {
            var creationPackage = MakePackageWithExtLstIntValidation();

            creationPackage.Workbook.Worksheets[0]
                .SparklineGroups.Add(OfficeOpenXml.Sparkline.eSparklineType.Line, new ExcelAddress("A1:A5"), new ExcelAddress("B1:B5"));
            creationPackage.Workbook.Worksheets[0]
                .SparklineGroups.Add(OfficeOpenXml.Sparkline.eSparklineType.Column, new ExcelAddress("C1:C5"), new ExcelAddress("D1:D5"));

            creationPackage.Workbook.Worksheets[0]
                .ConditionalFormatting.AddDatabar(new ExcelAddress("A1:A5"), Color.DarkBlue);
            creationPackage.Workbook.Worksheets[0]
                .ConditionalFormatting.AddDatabar(new ExcelAddress("A1:A5"), Color.Red);


            var decimalValidation = creationPackage.Workbook.Worksheets[0].DataValidations.AddDecimalValidation("B1");
            decimalValidation.Operator = ExcelDataValidationOperator.between;

            decimalValidation.Formula.ExcelFormula = "sheet2!B1";
            decimalValidation.Formula2.ExcelFormula = "1.5";

            var stream = new MemoryStream();
            creationPackage.SaveAs(stream);

            var readingPackage = new ExcelPackage(stream);
            var ws = readingPackage.Workbook.Worksheets[0];
            var validations = ws.DataValidations;

            Assert.AreEqual(OfficeOpenXml.Sparkline.eSparklineType.Line, ws.SparklineGroups[0].Type);
            Assert.AreEqual(OfficeOpenXml.Sparkline.eSparklineType.Column, ws.SparklineGroups[1].Type);

            Assert.AreEqual(eExcelConditionalFormattingRuleType.DataBar, ws.ConditionalFormatting[0].Type);
            Assert.AreEqual(eExcelConditionalFormattingRuleType.DataBar, ws.ConditionalFormatting[1].Type);

            Assert.AreEqual(ExcelDataValidationOperator.equal, validations[0].Operator);
            Assert.AreEqual("sheet2!A1", validations[0].As.IntegerValidation.Formula.ExcelFormula);
            Assert.AreEqual(InternalValidationType.ExtLst, validations[0].InternalValidationType);

            Assert.AreEqual(ExcelDataValidationOperator.between, validations[1].Operator);
            Assert.AreEqual("sheet2!B1", validations[1].As.DecimalValidation.Formula.ExcelFormula);
            Assert.AreEqual(1.5, validations[1].As.DecimalValidation.Formula2.Value);
            Assert.AreEqual(InternalValidationType.ExtLst, validations[1].InternalValidationType);
        }

    }
}

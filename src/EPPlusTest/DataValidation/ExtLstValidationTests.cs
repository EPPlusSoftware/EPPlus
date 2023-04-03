using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
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

                validation.Formula.ExcelFormula = "IF(A2=\"red\"";

                Assert.AreNotEqual(((ExcelDataValidationInt)validation).InternalValidationType, InternalValidationType.ExtLst);

                var stream = new MemoryStream();
                package.SaveAs(stream);

                ExcelPackage package2 = new ExcelPackage(stream);

                var readingValidation = package2.Workbook.Worksheets[0].DataValidations[0];

                Assert.AreEqual("IF(A2=\"red\"", readingValidation.As.IntegerValidation.Formula.ExcelFormula);
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
    }
}

using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.DataValidation;

namespace EPPlusTest.DataValidation
{
    [TestClass]
    public class UidTests : ValidationTestBase
    {
        [TestInitialize]
        public void Setup()
        {
            SetupTestData();
        }

        [TestCleanup]
        public void Cleanup()
        {
            CleanupTestData();
            _dataValidationNode = null;
        }

        [TestMethod]
        public void UidShouldBeSetOnValidations()
        {
            // Arrange
            LoadXmlTestData("A1", "decimal", "1.3");
            var id = ExcelDataValidation.NewId();
            // Act
            var validation = new ExcelDataValidationDecimal(id, "A1", _sheet.Name);
            // Assert
            Assert.AreEqual(id, validation.Uid);
        }
    }
}

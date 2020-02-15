using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.DataValidation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
            var validation = new ExcelDataValidationDecimal(_sheet, id, "A1", ExcelDataValidationType.Decimal, _dataValidationNode, _namespaceManager);
            // Assert
            Assert.AreEqual(id, validation.Uid);
        }
    }
}

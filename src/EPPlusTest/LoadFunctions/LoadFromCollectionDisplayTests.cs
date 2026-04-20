/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************/
#if !NET35
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using EPPlusTest.LoadFunctions.AttributesTestClasses;
using System.Collections.Generic;

namespace EPPlusTest.LoadFunctions
{
    [TestClass]
    public class LoadFromCollectionDisplayTests
    {
        private ExcelPackage _package;
        private ExcelWorksheet _sheet;

        [TestInitialize]
        public void Initialize()
        {
            _package = new ExcelPackage();
            _sheet = _package.Workbook.Worksheets.Add("test");
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }

        #region Order precedence tests

        [TestMethod]
        public void ShouldUseEpplusTableColumnOrderOverDisplayOrder()
        {
            // Arrange
            // EpplusTableColumn Order: Id=3, Name=1, Description=2
            // Display Order:           Id=1, Name=3, Description=2
            // Expected: EpplusTableColumn.Order wins
            var items = new List<ClassWithEpplusTableColumnAndDisplayOrder>
            {
                new ClassWithEpplusTableColumnAndDisplayOrder
                {
                    Id = 1,
                    Name = "Test",
                    Description = "A test"
                }
            };

            // Act
            _sheet.Cells["A1"].LoadFromCollection(items, true);

            // Assert - column order should follow EpplusTableColumnAttribute.Order
            Assert.AreEqual("Name Column", _sheet.Cells["A1"].Value, "First column should be Name (EpplusTableColumn Order=1)");
            Assert.AreEqual("Description Column", _sheet.Cells["B1"].Value, "Second column should be Description (EpplusTableColumn Order=2)");
            Assert.AreEqual("Id Column", _sheet.Cells["C1"].Value, "Third column should be Id (EpplusTableColumn Order=3)");
        }

        [TestMethod]
        public void ShouldUseEpplusTableColumnOrderWithNegativeValues()
        {
            // Arrange - matches the customer scenario with Order = -90
            // EpplusTableColumn Order: NumRegistro=-90, Nombre=1, Descripcion=2
            // Display Order:           NumRegistro=5,   Nombre=1, Descripcion=2
            // Expected: EpplusTableColumn.Order wins, NumRegistro first
            var items = new List<ClassWithNegativeEpplusOrder>
            {
                new ClassWithNegativeEpplusOrder
                {
                    NumRegistro = 42,
                    Nombre = "Test",
                    Descripcion = "A test"
                }
            };

            // Act
            _sheet.Cells["A1"].LoadFromCollection(items, true);

            // Assert
            Assert.AreEqual("NumRegistro", _sheet.Cells["A1"].Value, "First column should be NumRegistro (EpplusTableColumn Order=-90)");
            Assert.AreEqual("Nombre", _sheet.Cells["B1"].Value, "Second column should be Nombre (EpplusTableColumn Order=1)");
            Assert.AreEqual("Descripcion", _sheet.Cells["C1"].Value, "Third column should be Descripcion (EpplusTableColumn Order=2)");
        }

        [TestMethod]
        public void ShouldUseEpplusTableColumnOrderWithNegativeValues_DataOrder()
        {
            // Arrange - verify that the data row also follows the correct column order
            var items = new List<ClassWithNegativeEpplusOrder>
            {
                new ClassWithNegativeEpplusOrder
                {
                    NumRegistro = 42,
                    Nombre = "Test",
                    Descripcion = "A test"
                }
            };

            // Act
            _sheet.Cells["A1"].LoadFromCollection(items, true);

            // Assert - row 2 is data
            Assert.AreEqual(42, _sheet.Cells["A2"].Value, "First data column should be NumRegistro");
            Assert.AreEqual("Test", _sheet.Cells["B2"].Value, "Second data column should be Nombre");
            Assert.AreEqual("A test", _sheet.Cells["C2"].Value, "Third data column should be Descripcion");
        }

        [TestMethod]
        public void ShouldFallBackToDisplayOrderWhenEpplusOrderNotSet()
        {
            // Arrange
            // EpplusTableColumn exists but Order is NOT explicitly set
            // Display Order: Id=3, Name=1, Description=2
            // Expected: falls back to DisplayAttribute.Order
            var items = new List<ClassWithEpplusNoOrderAndDisplayWithOrder>
            {
                new ClassWithEpplusNoOrderAndDisplayWithOrder
                {
                    Id = 1,
                    Name = "Test",
                    Description = "A test"
                }
            };

            // Act
            _sheet.Cells["A1"].LoadFromCollection(items, true);

            // Assert - should follow DisplayAttribute.Order as fallback
            Assert.AreEqual("Name Column", _sheet.Cells["A1"].Value, "First column should be Name (Display Order=1)");
            Assert.AreEqual("Description Column", _sheet.Cells["B1"].Value, "Second column should be Description (Display Order=2)");
            Assert.AreEqual("Id Column", _sheet.Cells["C1"].Value, "Third column should be Id (Display Order=3)");
        }

        [TestMethod]
        public void ShouldUseDisplayOrderWhenNoEpplusTableColumnAttribute()
        {
            // Arrange - only DisplayAttribute present
            // Display Order: Id=3, Name=1, Description=2
            var items = new List<ClassWithDisplayOrderOnly>
            {
                new ClassWithDisplayOrderOnly
                {
                    Id = 1,
                    Name = "Test",
                    Description = "A test"
                }
            };

            // Act
            _sheet.Cells["A1"].LoadFromCollection(items, true);

            // Assert
            Assert.AreEqual("Name Column", _sheet.Cells["A1"].Value, "First column should be Name (Display Order=1)");
            Assert.AreEqual("Description Column", _sheet.Cells["B1"].Value, "Second column should be Description (Display Order=2)");
            Assert.AreEqual("Id Column", _sheet.Cells["C1"].Value, "Third column should be Id (Display Order=3)");
        }

        #endregion

        #region Header precedence tests

        [TestMethod]
        public void ShouldUseEpplusHeaderOverDisplayName()
        {
            // Arrange - both EpplusTableColumn.Header and Display.Name are set
            // Expected: EpplusTableColumn.Header wins
            var items = new List<ClassWithEpplusHeaderAndDisplayName>
            {
                new ClassWithEpplusHeaderAndDisplayName
                {
                    Id = 1,
                    Name = "Test"
                }
            };

            // Act
            _sheet.Cells["A1"].LoadFromCollection(items, true);

            // Assert
            Assert.AreEqual("EPPlus Id", _sheet.Cells["A1"].Value, "Header should come from EpplusTableColumn.Header");
            Assert.AreEqual("EPPlus Name", _sheet.Cells["B1"].Value, "Header should come from EpplusTableColumn.Header");
        }

        [TestMethod]
        public void ShouldFallBackToDisplayNameWhenEpplusHeaderNotSet()
        {
            // Arrange - EpplusTableColumn has Order but NOT Header
            // Display has Name set
            // Expected: falls back to DisplayAttribute.GetName()
            var items = new List<ClassWithEpplusOrderAndDisplayName>
            {
                new ClassWithEpplusOrderAndDisplayName
                {
                    Id = 1,
                    Name = "Test",
                    Description = "A test"
                }
            };

            // Act
            _sheet.Cells["A1"].LoadFromCollection(items, true);

            // Assert - headers should come from Display.Name, order from EpplusTableColumn.Order
            Assert.AreEqual("The Name", _sheet.Cells["A1"].Value, "Header should fall back to Display.Name");
            Assert.AreEqual("The Description", _sheet.Cells["B1"].Value, "Header should fall back to Display.Name");
            Assert.AreEqual("The Id", _sheet.Cells["C1"].Value, "Header should fall back to Display.Name");
        }

        [TestMethod]
        public void ShouldUseDisplayNameAsHeaderWhenNoEpplusAttribute()
        {
            // Arrange - only DisplayAttribute present
            var items = new List<ClassWithDisplayOrderOnly>
            {
                new ClassWithDisplayOrderOnly
                {
                    Id = 1,
                    Name = "Test",
                    Description = "A test"
                }
            };

            // Act
            _sheet.Cells["A1"].LoadFromCollection(items, true);

            // Assert
            Assert.AreEqual("Name Column", _sheet.Cells["A1"].Value);
            Assert.AreEqual("Description Column", _sheet.Cells["B1"].Value);
            Assert.AreEqual("Id Column", _sheet.Cells["C1"].Value);
        }

        [TestMethod]
        public void ShouldUseGetNameForDisplayAttributeWithResourceType()
        {
            // Arrange - DisplayAttribute uses ResourceType for localization
            // GetName() should return the localized value from the resource class,
            // not the raw Name property (which is just the resource key).
            // Display Name="IdHeader" with ResourceType=TestDisplayResources
            //   -> GetName() returns "Identifier"
            //   -> Name returns "IdHeader"
            var items = new List<ClassWithDisplayResourceType>
            {
                new ClassWithDisplayResourceType
                {
                    Id = 1,
                    Name = "Test",
                    Description = "A test"
                }
            };

            // Act
            _sheet.Cells["A1"].LoadFromCollection(items, true);

            // Assert - headers should be the localized values from TestDisplayResources
            Assert.AreEqual("Identifier", _sheet.Cells["A1"].Value, "Header should be localized value from resource class via GetName()");
            Assert.AreEqual("Full Name", _sheet.Cells["B1"].Value, "Header should be localized value from resource class via GetName()");
            Assert.AreEqual("Item Description", _sheet.Cells["C1"].Value, "Header should be localized value from resource class via GetName()");
        }

        [TestMethod]
        public void ShouldUseGetNameForDisplayAttributeWithResourceType_NoEpplusAttribute()
        {
            // Arrange - same as above but verify the resource type pattern also works
            // when no EpplusTableColumnAttribute is present (regression test)
            var items = new List<ClassWithDisplayResourceTypeOnly>
            {
                new ClassWithDisplayResourceTypeOnly
                {
                    Id = 1,
                    Name = "Test",
                    Description = "A test"
                }
            };

            // Act
            _sheet.Cells["A1"].LoadFromCollection(items, true);

            // Assert
            Assert.AreEqual("Identifier", _sheet.Cells["A1"].Value, "Header should be localized value from resource class via GetName()");
            Assert.AreEqual("Full Name", _sheet.Cells["B1"].Value, "Header should be localized value from resource class via GetName()");
            Assert.AreEqual("Item Description", _sheet.Cells["C1"].Value, "Header should be localized value from resource class via GetName()");
        }

        #endregion
    }
}
#endif
/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * Required Notice: Copyright (C) EPPlus Software AB. 
 * https://epplussoftware.com
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "" as is "" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
  Date               Author                       Change
 *******************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *******************************************************************************/
using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.Reflection;
using System.Linq;
using System.Runtime.ExceptionServices;

namespace EPPlusTest
{
	[TestClass]
	public class WorksheetsTests : TestBase
	{
		private ExcelPackage package;
		private ExcelWorkbook workbook;

		[TestInitialize]
		public void TestInitialize()
		{
			package = new ExcelPackage();
			workbook = package.Workbook;
			workbook.Worksheets.Add("NEW1");
		}

		[TestMethod]
		public void ConfirmFileStructure()
		{
			Assert.IsNotNull(package, "Package not created");
			Assert.IsNotNull(workbook, "No workbook found");
		}

		[TestMethod]
		public void ShouldBeAbleToDeleteAndThenAdd()
		{
			workbook.Worksheets.Add("NEW2");
			workbook.Worksheets.Delete(1);
			workbook.Worksheets.Add("NEW3");
		}

		[TestMethod]
		public void DeleteByNameWhereWorkSheetExists()
		{
		    workbook.Worksheets.Add("NEW2");
			workbook.Worksheets.Delete("NEW2");
        }

		[TestMethod, ExpectedException(typeof(ArgumentException))]
		public void DeleteByNameWhereWorkSheetDoesNotExist()
		{
			workbook.Worksheets.Add("NEW2");
			workbook.Worksheets.Delete("NEW3");
		}

		[TestMethod]
		public void MoveBeforeByNameWhereWorkSheetExists()
		{
			workbook.Worksheets.Add("NEW2");
			workbook.Worksheets.Add("NEW3");
			workbook.Worksheets.Add("NEW4");
			workbook.Worksheets.Add("NEW5");

			workbook.Worksheets.MoveBefore("NEW4", "NEW2");

			CompareOrderOfWorksheetsAfterSaving(package);
		}

		[TestMethod]
		public void MoveAfterByNameWhereWorkSheetExists()
		{
			workbook.Worksheets.Add("NEW2");
			workbook.Worksheets.Add("NEW3");
			workbook.Worksheets.Add("NEW4");
			workbook.Worksheets.Add("NEW5");

			workbook.Worksheets.MoveAfter("NEW4", "NEW2");

			CompareOrderOfWorksheetsAfterSaving(package);
		}

		[TestMethod]
		public void MoveBeforeByPositionWhereWorkSheetExists()
		{
			workbook.Worksheets.Add("NEW2");
			workbook.Worksheets.Add("NEW3");
			workbook.Worksheets.Add("NEW4");
			workbook.Worksheets.Add("NEW5");

			workbook.Worksheets.MoveBefore(4, 2);

			CompareOrderOfWorksheetsAfterSaving(package);
		}

		[TestMethod]
		public void MoveAfterByPositionWhereWorkSheetExists()
		{
			workbook.Worksheets.Add("NEW2");
			workbook.Worksheets.Add("NEW3");
			workbook.Worksheets.Add("NEW4");
			workbook.Worksheets.Add("NEW5");

			workbook.Worksheets.MoveAfter(4, 2);

			CompareOrderOfWorksheetsAfterSaving(package);
		}

        [TestMethod]
        public void MoveToStartByNameWhereWorkSheetExists()
        {
            workbook.Worksheets.Add("NEW2");

            workbook.Worksheets.MoveToStart("NEW2");

            Assert.AreEqual("NEW2", workbook.Worksheets.First().Name);
        }

        [TestMethod]
        public void MoveToEndByNameWhereWorkSheetExists()
        {
            workbook.Worksheets.Add("NEW2");

            workbook.Worksheets.MoveToEnd("NEW1");

            Assert.AreEqual("NEW1", workbook.Worksheets.Last().Name);
        }
		[TestMethod]
		public void ShouldHandleResizeOfIndexWhenExceed8Items()
		{
			using (var p = new ExcelPackage())
			{
				ExcelWorksheet wsStart = p.Workbook.Worksheets.Add($"Copy");
				for (int i = 0; i < 7; i++)
				{
					ExcelWorksheet wsNew = p.Workbook.Worksheets.Add($"Sheet{i}");
					p.Workbook.Worksheets.MoveBefore(wsStart.Name, wsNew.Name);
				}
			}
		}
		[TestMethod]
		public void MoveBeforeByName8Worksheets()
		{
			workbook.Worksheets.Add("NEW2");
			workbook.Worksheets.Add("NEW3");
			workbook.Worksheets.Add("NEW4");
			workbook.Worksheets.Add("NEW5");
			workbook.Worksheets.Add("NEW6");
			workbook.Worksheets.Add("NEW7");
			workbook.Worksheets.Add("NEW8");

			workbook.Worksheets.MoveBefore("NEW8", "NEW1");
			Assert.AreEqual("NEW7", workbook.Worksheets.Last().Name);
			Assert.AreEqual("NEW8", workbook.Worksheets.First().Name);
			Assert.AreEqual("NEW1", workbook.Worksheets[1].Name);
		}
		private static void CompareOrderOfWorksheetsAfterSaving(ExcelPackage editedPackage)
		{
			var packageStream = new MemoryStream();
			editedPackage.SaveAs(packageStream);

			var newPackage = new ExcelPackage(packageStream);
            var positionId = newPackage._worksheetAdd;
			foreach (var worksheet in editedPackage.Workbook.Worksheets)
			{
				Assert.AreEqual(worksheet.Name, newPackage.Workbook.Worksheets[positionId].Name, "Worksheets are not in the same order");
				positionId++;
			}
		}
        [TestMethod]
        public void CheckAddedWorksheetWithInvalidName()
        {
            if (workbook.Worksheets["[NEW2]"] == null)
                workbook.Worksheets.Add("[NEW2]");

			Assert.IsNotNull(workbook.Worksheets["[NEW2]"]);
        }

        [TestMethod]
        public void DeletingSheetMovesSelectedSheetCorrectly()
        {
            using (var package = OpenPackage("deletedSheets.xlsx", true))
            {
                package.Workbook.Worksheets.Add("VisibleSheet1");
                package.Workbook.Worksheets.Add("HiddenSheet1").Hidden = eWorkSheetHidden.Hidden;
                package.Workbook.Worksheets.Add("VisibleSheet2");
                package.Workbook.Worksheets.Add("HiddenSheet2").Hidden = eWorkSheetHidden.Hidden;
                package.Workbook.Worksheets.Add("HiddenSheet3").Hidden = eWorkSheetHidden.VeryHidden;
                package.Workbook.Worksheets.Add("VisibleSheet3");
                package.Workbook.Worksheets.Add("VisibleSheet4");
                package.Workbook.Worksheets.Add("HiddenSheet4").Hidden = eWorkSheetHidden.Hidden;
                package.Workbook.View.ActiveTab = 2;
                package.Workbook.Worksheets.Delete("VisibleSheet2");
                Assert.AreEqual(4, package.Workbook.View.ActiveTab);
                package.Workbook.View.ActiveTab = package.Workbook.Worksheets.GetByName("VisibleSheet4").Index;
                package.Workbook.Worksheets.Delete("VisibleSheet4");
                Assert.AreEqual(package.Workbook.Worksheets.GetByName("VisibleSheet3").Index, package.Workbook.View.ActiveTab);
                package.Workbook.Worksheets.Delete("HiddenSheet4");
                Assert.AreEqual(package.Workbook.Worksheets.GetByName("VisibleSheet3").Index, package.Workbook.View.ActiveTab);
                package.Workbook.Worksheets.Delete("VisibleSheet3");
                Assert.AreEqual(0, package.Workbook.View.ActiveTab);
                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void DeletingSheetBeforeSelectedSheetMovesCorrectly()
        {
            using (var package = OpenPackage("deletedSheets.xlsx", true))
            {
                package.Workbook.Worksheets.Add("VisibleSheet1");
                package.Workbook.Worksheets.Add("HiddenSheet1").Hidden = eWorkSheetHidden.Hidden;
                package.Workbook.Worksheets.Add("VisibleSheet2");
                package.Workbook.Worksheets.Add("HiddenSheet2").Hidden = eWorkSheetHidden.Hidden;
                package.Workbook.Worksheets.Add("HiddenSheet3").Hidden = eWorkSheetHidden.VeryHidden;
                package.Workbook.Worksheets.Add("VisibleSheet3");
                package.Workbook.Worksheets.Add("VisibleSheet4");
                package.Workbook.Worksheets.Add("HiddenSheet4").Hidden = eWorkSheetHidden.Hidden;

                package.Workbook.View.ActiveTab = 2;

                package.Workbook.Worksheets.Delete("VisibleSheet1");

                Assert.AreEqual(1, package.Workbook.View.ActiveTab);

                package.Workbook.View.ActiveTab = 4;

                package.Workbook.Worksheets.Delete("HiddenSheet3");

                Assert.AreEqual(package.Workbook.Worksheets.GetByName("VisibleSheet3").Index, package.Workbook.View.ActiveTab);

                package.Workbook.Worksheets.Delete("VisibleSheet4");
                package.Workbook.Worksheets.Delete("VisibleSheet3");

                Assert.AreEqual(package.Workbook.Worksheets.GetByName("VisibleSheet2").Index, package.Workbook.View.ActiveTab);

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void DeletedSheetMovesCorrectlyIsWorksheet1Based()
        {
            using (var package = OpenPackage("deletedSheets.xlsx", true))
            {
                package.Compatibility.IsWorksheets1Based = true;

                package.Workbook.Worksheets.Add("VisibleSheet1");
                package.Workbook.Worksheets.Add("VisibleSheet2");
                package.Workbook.Worksheets.Add("HiddenSheet2").Hidden = eWorkSheetHidden.Hidden;
                package.Workbook.Worksheets.Add("HiddenSheet3").Hidden = eWorkSheetHidden.VeryHidden;
                package.Workbook.Worksheets.Add("VisibleSheet3");
                package.Workbook.Worksheets.Add("VisibleSheet4");
                package.Workbook.Worksheets.Add("HiddenSheet4").Hidden = eWorkSheetHidden.Hidden;

                package.Workbook.View.ActiveTab = 5;

                package.Workbook.Worksheets.Delete("HiddenSheet4");

                Assert.AreEqual(5, package.Workbook.View.ActiveTab);

                package.Workbook.Worksheets.Delete("VisibleSheet4");

                Assert.AreEqual(4, package.Workbook.View.ActiveTab);

                package.Workbook.View.ActiveTab = 1;

                package.Workbook.Worksheets.Delete("VisibleSheet2");

                Assert.AreEqual(3, package.Workbook.View.ActiveTab);

                package.Workbook.View.ActiveTab = 0;

                package.Workbook.Worksheets.Delete("VisibleSheet1");

                Assert.AreEqual(2, package.Workbook.View.ActiveTab);

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void NoVisibleSheetShouldThrow()
        {
            using (var package = new ExcelPackage("ExceptionSheet.xlsx"))
            {
                package.Workbook.Worksheets.Add("VisibleSheet1");
                package.Workbook.Worksheets.Add("HiddenSheet1").Hidden = eWorkSheetHidden.Hidden;
                package.Workbook.Worksheets.Add("HiddenSheet2").Hidden = eWorkSheetHidden.Hidden;
                package.Workbook.Worksheets.Delete(0);
                package.Save();
            }
        }
    }
}

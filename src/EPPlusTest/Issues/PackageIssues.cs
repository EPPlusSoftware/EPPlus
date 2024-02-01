using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System.IO;
namespace EPPlusTest.Issues
{
	[TestClass]
	public class PackageIssues : TestBase
	{
		[ClassInitialize]
		public static void Init(TestContext context)
		{
		}
		[ClassCleanup]
		public static void Cleanup()
		{
		}
		[TestInitialize]
		public void Initialize()
		{
		}
		[TestMethod]
		public void s593()
		{
			using (ExcelPackage package = OpenTemplatePackage("s593.xlsx"))
			{
				SaveWorkbook("s593FirstSave.xlsx", package);
				package.Workbook.Worksheets.Add("Sheet1");
				SaveWorkbook("s593Second.xlsx", package);
				Assert.AreEqual(73, package.Workbook.Worksheets[0].Part._rels.Count);

				using (var savedPackage = OpenTemplatePackage("s593Second.xlsx"))
				{
					Assert.AreEqual(13, savedPackage.Workbook.Styles.Dxfs.Count);
					Assert.AreEqual(4, savedPackage.Workbook.Worksheets[0].Part._rels.Count);
				}
			}
		}
	}
}
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
namespace EPPlusTest.Issues
{
	[TestClass]
	public class DrawingIssues : TestBase
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
		public void s633()
		{
			using (var p = OpenTemplatePackage("s633.xlsx"))
			{
				var sheet = p.Workbook.Worksheets[0];
				var pic=sheet.Drawings[0].As.Picture;
			}
		}
	}
}

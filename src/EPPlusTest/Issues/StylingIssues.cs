using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
namespace EPPlusTest
{
	[TestClass]
	public class StylingIssues : TestBase
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
		public void i1291()
		{
			using(var p=OpenPackage("i1291.xlsx", true))
			{
				var ws = p.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["A1"].Style.Font.Name = "+Headings";
				SaveAndCleanup(p);
			}
		}
	}
}

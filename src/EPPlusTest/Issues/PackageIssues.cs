using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
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
        [TestMethod]
        public void i1440()
        {
            using (ExcelPackage package = OpenTemplatePackage("i1440.xlsx"))
            {
				var ws = package.Workbook.Worksheets[0];
            }
        }
        [TestMethod]
        public void s703()
        {
			var fi = GetTemplateFile("errorfile.xlsx");
			if (fi != null)
			{
                var formFile = fi.OpenRead();

                using (ExcelPackage package = new ExcelPackage(formFile))
                {
                    var ws = package.Workbook.Worksheets[0];
                    SaveWorkbook("s703.xlsx", package);
                }
                formFile.Close();
                formFile.Dispose();
            }
        }
        [TestMethod]
        public void i1530()
        {
            using (ExcelPackage package = OpenTemplatePackage("i1530.xlsx"))
            {
                var ws = package.Workbook.Worksheets[0];
				Assert.AreEqual("a", ws.Cells["A2"].Value);
            }
        }
        [TestMethod]
        public void i1388()
        {
            using (ExcelPackage package = OpenTemplatePackage("i1388.xlsx"))
            {
                var ws = package.Workbook.Worksheets[0];
                SaveAndCleanup(package);
            }
        }
    }
 }
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
        [TestMethod]
        public void s711()
        {
            using (ExcelPackage package = OpenTemplatePackage("s711.xlsx"))
            {
                var ws = package.Workbook.Worksheets[0];
                SaveAndCleanup(package);
            }
        }
        [TestMethod, Ignore] //This test is set to be ignored as it test creates a workbook with the worksheet xml exceeding 2GB. This causes the the stream and the rolling buffer to read directly from the zip stream.
        public void s699() 
        {
            using var p = OpenPackage("s699.xlsx");
            var ws = p.Workbook.Worksheets.Add("Sheet1");
            for (int c = 1; c <= 250; c++)
            {
                for (int r = 1; r <= 250000; r++)
                {
                    ws.SetValue(r, c, c + r);
                }
            }
            SaveAndCleanup(p);
        }
        [TestMethod, Ignore] 
        public void s699_Read()
        {
            using var p = OpenPackage("s699.xlsx");
            var ws = p.Workbook.Worksheets["Sheet1"];

            SaveWorkbook("s699-resaved.xlsx", p);
        }

    }
}
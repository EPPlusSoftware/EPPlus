using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Drawing;
namespace EPPlusTest.Issues
{
    [TestClass]
	public class PictureIssues : TestBase

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
		public void i1389()
		{
			using (var p = OpenPackage("i1389.xlsx", true))
			{
				p.Settings.ImageSettings.PrimaryImageHandler = new GenericImageHandler();
				var ws = p.Workbook.Worksheets.Add("Sheet1");
				var stream = GetImageMemoryStream("i1389.jpg");
				ExcelPicture pic = ws.Drawings.AddPicture("s", stream);
				SaveAndCleanup(p);
			}
		}
        [TestMethod]
        public void i1982()
        {
            using (var p = OpenTemplatePackage("PictureRead.xlsx"))
            {
				var ws1 = p.Workbook.Worksheets[0];
				foreach(var d in ws1.Drawings)
				{
                    var pic = d.As.Picture;
					Assert.IsNotNull(pic.Image.Bounds);
					var pic2 = new ExcelImage();
                }

                SaveAndCleanup(p);
            }
        }
    }
}

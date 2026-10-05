using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Utils.FileUtils;
using System.IO;

namespace EPPlusTest.Utils
{
    [TestClass]
	public class FileHelperTests

	{
		static readonly char separator = Path.DirectorySeparatorChar;

        [TestMethod]
		public void ValidateGetRelativeFile()
		{
            var file = FileHelper.GetRelativeFile(new FileInfo($"FileSource.xlsx"), new FileInfo($"FileTarget.xlsx"));
			Assert.AreEqual($"FileTarget.xlsx", file);

			file = FileHelper.GetRelativeFile(new FileInfo($"c:{separator}FileSource.xlsx"), new FileInfo($"c:{separator}Dir1{separator}FileTarget.xlsx"));
			Assert.AreEqual($"Dir1{separator}FileTarget.xlsx", file);

			file = FileHelper.GetRelativeFile(new FileInfo($"c:{separator}Dir1{separator}FileSource.xlsx"), new FileInfo($"c:{separator}FileTarget.xlsx"));
			Assert.AreEqual($"..{separator}FileTarget.xlsx", file);

			file = FileHelper.GetRelativeFile(new FileInfo($"c:{separator}Dir1{separator}Dir2{separator}FileSource.xlsx"), new FileInfo($"c:{separator}Dir1{separator}Dir1{separator}FileTarget.xlsx"));
			Assert.AreEqual($"..{separator}Dir1{separator}FileTarget.xlsx", file);

			file = FileHelper.GetRelativeFile(new FileInfo($"c:{separator}Dir1{separator}FileSource.xlsx"), new FileInfo($"c:{separator}Dir1{separator}Dir1{separator}FileTarget.xlsx"));
			Assert.AreEqual($"Dir1{separator}FileTarget.xlsx", file);
		}
	}
}

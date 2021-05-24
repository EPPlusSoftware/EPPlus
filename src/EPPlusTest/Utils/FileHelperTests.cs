using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Utils
{
	[TestClass]
	public class FileHelperTests

	{
		[TestMethod]
		public void ValidateGetRelativeFile()
		{
			var file = FileHelper.GetRelativeFile(new System.IO.FileInfo("FileSource.xlsx"), new System.IO.FileInfo("FileTarget.xlsx"));
			Assert.AreEqual("FileTarget.xlsx", file);

			file = FileHelper.GetRelativeFile(new System.IO.FileInfo("c:\\FileSource.xlsx"), new System.IO.FileInfo("c:\\Dir1\\FileTarget.xlsx"));
			Assert.AreEqual("Dir1\\FileTarget.xlsx", file);

			file = FileHelper.GetRelativeFile(new System.IO.FileInfo("c:\\Dir1\\FileSource.xlsx"), new System.IO.FileInfo("c:\\FileTarget.xlsx"));
			Assert.AreEqual("..\\FileTarget.xlsx", file);

			file = FileHelper.GetRelativeFile(new System.IO.FileInfo("c:\\Dir1\\Dir2\\FileSource.xlsx"), new System.IO.FileInfo("c:\\Dir1\\Dir1\\FileTarget.xlsx"));
			Assert.AreEqual("..\\Dir1\\FileTarget.xlsx", file);

			file = FileHelper.GetRelativeFile(new System.IO.FileInfo("c:\\Dir1\\FileSource.xlsx"), new System.IO.FileInfo("c:\\Dir1\\Dir1\\FileTarget.xlsx"));
			Assert.AreEqual("Dir1\\FileTarget.xlsx", file);
		}
	}
}

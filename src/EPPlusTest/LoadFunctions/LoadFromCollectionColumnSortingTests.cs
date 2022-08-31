using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Attributes;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.LoadFunctions
{
	[EpplusTable(PrintHeaders = true)]
	public class Columns17
	{

		[EpplusTableColumn(Header = "C01")]
		public int C01 { get; set; }

		[EpplusTableColumn(Header = "C02")]
		public int C02 { get; set; }

		[EpplusTableColumn(Header = "C03")]
		public int C03 { get; set; }

		[EpplusTableColumn(Header = "C04")]
		public int C04 { get; set; }

		[EpplusTableColumn(Header = "C05")]
		public int C05 { get; set; }

		[EpplusTableColumn(Header = "C06")]
		public int C06 { get; set; }

		[EpplusTableColumn(Header = "C07")]
		public int C07 { get; set; }

		[EpplusTableColumn(Header = "C08")]
		public int C08 { get; set; }

		[EpplusTableColumn(Header = "C09")]
		public int C09 { get; set; }

		[EpplusTableColumn(Header = "C10")]
		public int C10 { get; set; }

		[EpplusTableColumn(Header = "C11")]
		public int C11 { get; set; }

		[EpplusTableColumn(Header = "C12")]
		public int C12 { get; set; }

		[EpplusTableColumn(Header = "C13")]
		public int C13 { get; set; }

		[EpplusTableColumn(Header = "C14")]
		public int C14 { get; set; }

		[EpplusTableColumn(Header = "C15")]
		public int C15 { get; set; }

		[EpplusTableColumn(Header = "C16")]
		public int C16 { get; set; }

		[EpplusTableColumn(Header = "C17")]
		public int C17 { get; set; }

	}

	[TestClass]
    public class LoadFromCollectionColumnSortingTests
    {
		/// <summary>
		/// The reason for this test is that the .NET sorting function used seems to change sort algorithm when more than 16 items are sorted.
		/// Therefore we must use the index of the column (the order that the properties are returned when using reflection on the class)
		/// to sort. If this isn't done the sorting will generate a strange result.
		/// </summary>
        [TestMethod]
        public void ShouldUseIndexWhenMoreThan17Properties()
        {
			using (var excel = new ExcelPackage())
			{
				var tableData1 = Enumerable.Range(1, 10)
				.Select(_ => new Columns17
				{
					C01 = 1,
					C02 = 2,
					C03 = 3,
					C04 = 4,
					C05 = 5,
					C06 = 6,
					C07 = 7,
					C08 = 8,
					C09 = 9,
					C10 = 10,
					C11 = 11,
					C12 = 12,
					C13 = 13,
					C14 = 14,
					C15 = 15,
					C16 = 16,
					C17 = 17
				}).ToArray();
				var sheet = excel.Workbook.Worksheets.Add("16Columns");
				sheet.Cells["A1"].LoadFromCollection(tableData1);

				for(int i = 1; i < 18; i++)
                {
					var expected = i < 10 ? "C0" + i : "C" + i;
					Assert.AreEqual(expected, sheet.Cells[1, i].Value, $"Value of cell [[1, {i}] vas not {expected}");
                }
				
			}
		}
	}
}

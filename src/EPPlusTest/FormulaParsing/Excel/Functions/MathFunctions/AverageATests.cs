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
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Ranges;

namespace EPPlusTest.FormulaParsing.Excel.Functions.MathFunctions
{
	[TestClass]
	public class AverageATests
	{
		[TestMethod]
		public void AverageALiterals()
		{
			// For literals, AverageA always parses and include numeric strings, date strings, bools, etc.
			// The only exception is unparsable string literals, which cause a #VALUE.
			AverageA average = new AverageA();
			var date1 = new DateTime(2013, 1, 5);
			var date2 = new DateTime(2013, 1, 15);
			double value1 = 1000;
			double value2 = 2000;
			double value3 = 6000;
			double value4 = 1;
			double value5 = date1.ToOADate();
			double value6 = date2.ToOADate();
			var result = average.Execute(new FunctionArgument[]
			{
				new FunctionArgument(value1.ToString("n")),
				new FunctionArgument(value2),
				new FunctionArgument(value3.ToString("n")),
				new FunctionArgument(true),
				new FunctionArgument(date1),
				new FunctionArgument(date2.ToString("d"))
			}, ParsingContext.Create());
			Assert.AreEqual((value1 + value2 + value3 + value4 + value5 + value6) / 6, result.Result, $"Current Culture is {Thread.CurrentThread.CurrentCulture.Name}");
        }

		[TestMethod]
		public void AverageACellReferences()
		{
			// For cell references, AverageA divides by all cells, but only adds actual numbers, dates, and booleans.
			ExcelPackage package = new ExcelPackage();
			var ctx = ParsingContext.Create(package);
			var worksheet = package.Workbook.Worksheets.Add("Test");
			double[] values =
			{
				0,
				2000,
				0,
				1,
				new DateTime(2013, 1, 5).ToOADate(),
				0
			};
			ExcelRange range1 = worksheet.Cells[1, 1];
			range1.Formula = "\"1000\"";
			range1.Calculate();
			var range2 = worksheet.Cells[1, 2];
			range2.Value = 2000;
			var range3 = worksheet.Cells[1, 3];
			range3.Formula = $"\"{new DateTime(2013, 1, 5).ToString("d")}\"";
			range3.Calculate();
			var range4 = worksheet.Cells[1, 4];
			range4.Value = true;
			var range5 = worksheet.Cells[1, 5];
			range5.Value = new DateTime(2013, 1, 5);
			var range6 = worksheet.Cells[1, 6];
			range6.Value = "Test";
			AverageA average = new AverageA();
			var rangeInfo1 = new RangeInfo(worksheet, 1, 1, 1, 3, ctx);
			var rangeInfo2 = new RangeInfo(worksheet, 1, 4, 1, 4, ctx);
			var rangeInfo3 = new RangeInfo(worksheet, 1, 5, 1, 6, ctx);
			var context = ParsingContext.Create();
			var address = new FormulaRangeAddress(context);
			address.FromRow = address.ToRow = address.FromCol = address.ToCol = 2;
			var result = average.Execute(new FunctionArgument[]
			{
				new FunctionArgument(rangeInfo1),
				new FunctionArgument(rangeInfo2),
				new FunctionArgument(rangeInfo3)
			}, context);
			Assert.AreEqual(values.Average(), result.Result);
		}

		[TestMethod]
		public void AverageAArray()
		{
			// For arrays, AverageA completely ignores booleans.  It divides by strings and numbers, but only
			// numbers are added to the total.  Real dates cannot be specified and string dates are not parsed.
			AverageA average = new AverageA();
			var date = new DateTime(2013, 1, 15);
			double[] values =
			{
				0,
				2000,
				0,
				0,
				0,
				0
			};
			var range = new InMemoryRange(6, 1);
			range.SetValue(0, 0, 1000.ToString("n"));
			range.SetValue(1, 0, 2000);
            range.SetValue(2, 0, 6000.ToString("n"));
            range.SetValue(3, 0, true);
            range.SetValue(4, 0, date.ToString("d"));
            range.SetValue(5, 0,"test");
            var result = average.Execute(
				new List<FunctionArgument>()
				{
					new FunctionArgument(range)
				}
			, ParsingContext.Create());
			Assert.AreEqual(values.Average(), result.Result);
		}

		[TestMethod]
		[ExpectedException(typeof(ExcelErrorValueException))]
		public void AverageAUnparsableLiteral()
		{
			// In the case of literals, any unparsable string literal results in a #VALUE.
			AverageA average = new AverageA();
			var result = average.Execute(new FunctionArgument[]
			{
				new FunctionArgument(1000),
				new FunctionArgument("Test")
			}, ParsingContext.Create());
		}
	}
}

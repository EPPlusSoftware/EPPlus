using System;
using System.Drawing;
using System.IO;
using System.Xml;
using EPPlusTest.Properties;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System.Diagnostics;
using System.Reflection;
using OfficeOpenXml.Drawing.Theme;

namespace EPPlusTest.Drawing
{
    [TestClass]
    public class EquationTests : TestBase
    {
        [TestMethod]
        public void Equations()
        {
            FileInfo fileInfo = new FileInfo(@"C:\epplusTest\Workbooks\Equations02.xlsx");
            using var p = new ExcelPackage(fileInfo);
            foreach (var d in p.Workbook.Worksheets[0].Drawings)
            {
                Console.WriteLine(d.Name);
                Console.WriteLine(d.DrawingType.ToString());
                Console.WriteLine(" ");
            }
        }

    }
}

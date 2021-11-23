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
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;
using System.Reflection;
using OfficeOpenXml.Drawing;
using System.Threading.Tasks;
using System.Collections.Generic;

namespace EPPlusTest
{
    [TestClass]
    public abstract class TestBase
    {
        private class SalesData
        {
            public string Continent { get; set; }
            public string Country { get; set; }
            public string State { get; set; }
            public double Sales { get; set; }

        }
        private class GeoData
        {
            public string Country { get; set; }
            public string State { get; set; }
            public double Sales { get; set; }

        }
        protected static FileInfo _file;
        protected static string _clipartPath ="";
        protected static string _worksheetPath = @"c:\epplusTest\Testoutput\";
        protected static string _testInputPath = AppContext.BaseDirectory + "\\workbooks\\";
        protected static string _testInputPathOptional = @"c:\epplusTest\workbooks\";
        public TestContext TestContext { get; set; }
        
        public static void InitBase()
        {
            _clipartPath = Path.Combine(Path.GetTempPath(), @"EPPlus clipart");
            if (!Directory.Exists(_clipartPath))
            {
                Directory.CreateDirectory(_clipartPath);
            }
            if(Environment.GetEnvironmentVariable("EPPlusTestInputPath")!=null)
            {
                _testInputPathOptional = Environment.GetEnvironmentVariable("EPPlusTestInputPath");
            }
            var asm = Assembly.GetExecutingAssembly();
            var validExtensions = new[]
                {
                    ".gif", ".wmf"
                };

            foreach (var name in asm.GetManifestResourceNames())
            {
                foreach (var ext in validExtensions)
                {
                    if (name.EndsWith(ext, StringComparison.OrdinalIgnoreCase))
                    {
                        string fileName = name.Replace("EPPlusTest.Resources.", "");
                        using (var stream = asm.GetManifestResourceStream(name))
                        using (var file = File.Create(Path.Combine(_clipartPath, fileName)))
                        {
                            stream.CopyTo(file);
                        }
                        break;
                    }
                }
            }
            
            var di=new DirectoryInfo(_worksheetPath);            
            _worksheetPath = di.FullName + "\\";
        }
        /// <summary>
        /// Saves and disposes a package
        /// </summary>
        /// <param name="pck"></param>
        protected static void SaveAndCleanup(ExcelPackage pck)
        {
            if (pck.Workbook.Worksheets.Count > 0)
            {
                pck.Save();
            }
            pck.Dispose();
        }

        protected static bool ExistsPackage(string name)
        {
            var fi = new FileInfo(_worksheetPath + name);
            return fi.Exists;
        }
        protected static void AssertIfNotExists(string name)
        {
            if(!ExistsPackage(name))
            {
                Assert.Inconclusive($"{_worksheetPath}{name} workbook is missing");
            }
        }
        protected static ExcelPackage OpenPackage(string name, bool delete=false)
        {
            CreateWorksheetPathIfNotExists();
            _file = new FileInfo(_worksheetPath + name);
            if(delete && _file.Exists)
            {
                _file.Delete();
            }
            return new ExcelPackage(_file);
        }
        protected static async Task<ExcelPackage> OpenPackageAsync(string name, bool delete = false, string password=null)
        {
            CreateWorksheetPathIfNotExists();
            var _file = new FileInfo(_worksheetPath + name);
            if (delete && _file.Exists)
            {
                _file.Delete();
            }
            var p = new ExcelPackage();
            if (password == null)
            {
                await p.LoadAsync(_file).ConfigureAwait(false);
            }
            else
            {
                await p.LoadAsync(_file, password).ConfigureAwait(false);
            }
            return p;
        }

        static void CreateWorksheetPathIfNotExists()
        {
            CreatePathIfNotExists(_worksheetPath);
        }
        protected static void CreatePathIfNotExists(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }
        protected static ExcelPackage OpenTemplatePackage(string name)
        {
            var t = new FileInfo(_testInputPath  + name);
            if (t.Exists)
            {
                var file = new FileInfo(_worksheetPath + name);
                return new ExcelPackage(file, t);
            }
            else
            {
                t = new FileInfo(_testInputPathOptional + name);
                if (t.Exists)
                {
                    var file = new FileInfo(_worksheetPath + name);
                    return new ExcelPackage(file, t);  
                }
                t = new FileInfo(_worksheetPath + name);
                if (t.Exists)
                {
                    return new ExcelPackage(t);
                }
                Assert.Inconclusive($"Template {name} does not exist in path {_testInputPath}");
            }
            return null;
        }

        internal void IsNullRange(ExcelRange address)
        {
            for(int row=address._fromRow;row<=address._toRow;row++)
            {
                for (int col = address._fromCol; col <= address._toCol; col++)
                {
                    Assert.IsNull(address._worksheet.Cells[row, col].Value);
                }
            }
        }
        protected void SaveWorkbook(string name, ExcelPackage pck)
            {
            if (pck.Workbook.Worksheets.Count == 0) return;
            var fi = new FileInfo(_worksheetPath + name);
            if (fi.Exists)
            {
                //fi.Delete();
            }
            pck.SaveAs(fi);
        }
        protected static readonly DateTime _loadDataStartDate = new DateTime(DateTime.Today.Year-1, 11, 1);
        /// <summary>
        /// Loads 4 columns of {date, numeric, string, numeric}
        /// </summary>
        /// <param name="ws">The worksheet </param>
        /// <param name="noItems">Number of items</param>
        /// <param name="startColumn">The start column</param>
        /// <param name="startRow">The start row</param>
        /// <param name="addHyperlinkColumn">Add a column with hyperlinks</param>
        protected static void LoadTestdata(ExcelWorksheet ws, int noItems = 100, int startColumn=1, int startRow=1, bool addHyperlinkColumn=false)
        {
            ws.SetValue(1, startColumn, "Date");
            ws.SetValue(1, startColumn + 1, "NumValue");
            ws.SetValue(1, startColumn + 2, "StrValue");
            ws.SetValue(1, startColumn + 3, "NumFormattedValue");
            if(addHyperlinkColumn)
            {
                ws.SetValue(1, startColumn + 4, "HyperLink");
            }

            DateTime dt = _loadDataStartDate;
            int row = 1;
            for (int i = 1; i < noItems; i++)
            {
                row = startRow + i;
                ws.SetValue(row, startColumn, dt);
                ws.SetValue(row, startColumn + 1, row);
                ws.SetValue(row, startColumn + 2, $"Value {row}");
                ws.SetValue(row, startColumn + 3, row * 33);
                if (addHyperlinkColumn)
                {
                    ws.Cells[row, startColumn + 4].SetHyperlink(new Uri("https://epplussoftware.com"));
                }

                dt = dt.AddDays(1);
            }
            ws.Cells[startRow, startColumn, row, startColumn].Style.Numberformat.Format = "yyyy-MM-dd";
            ws.Cells.AutoFitColumns();
        }
        protected static void LoadHierarkiTestData(ExcelWorksheet ws)
        {

            var l = new List<SalesData>
            {
                new SalesData{ Continent="Europe", Country="Sweden", State = "Stockholm", Sales = 154 },
                new SalesData{ Continent="Asia", Country="Vietnam", State = "Ho Chi Minh", Sales= 88 },
                new SalesData{ Continent="Europe", Country="Sweden", State = "Västerås", Sales = 33 },
                new SalesData{ Continent="Asia", Country="Japan", State = "Tokyo", Sales= 534 },
                new SalesData{ Continent="Europe", Country="Germany", State = "Frankfurt", Sales = 109 },
                new SalesData{ Continent="Asia", Country="Vietnam", State = "Hanoi", Sales= 322 },
                new SalesData{ Continent="Asia", Country="Japan", State = "Osaka", Sales= 88 },
                new SalesData{ Continent="North America", Country="Canada", State = "Vancover", Sales= 99 },
                new SalesData{ Continent="Asia", Country="China", State = "Peking", Sales= 205 },
                new SalesData{ Continent="North America", Country="Canada", State = "Toronto", Sales= 138 },
                new SalesData{ Continent="Europe", Country="France", State = "Lyon", Sales = 185 },
                new SalesData{ Continent="North America", Country="USA", State = "Boston", Sales= 155 },
                new SalesData{ Continent="Europe", Country="France", State = "Paris", Sales = 127 },
                new SalesData{ Continent="North America", Country="USA", State = "New York", Sales= 330 },
                new SalesData{ Continent="Europe", Country="Germany", State = "Berlin", Sales = 210 },
                new SalesData{ Continent="North America", Country="USA", State = "San Fransico", Sales= 411 },
            };

            ws.Cells["A1"].LoadFromCollection(l, true, OfficeOpenXml.Table.TableStyles.Medium12);
        }
        protected static void LoadGeoTestData(ExcelWorksheet ws)
        {

            var l = new List<GeoData>
            {
                new GeoData{ Country="Sweden", State = "Stockholm", Sales = 154 },
                new GeoData{ Country="Sweden", State = "Jämtland", Sales = 55 },
                new GeoData{ Country="Sweden", State = "Västerbotten", Sales = 44},
                new GeoData{ Country="Sweden", State = "Dalarna", Sales = 33 },
                new GeoData{ Country="Sweden", State = "Uppsala", Sales = 22 },
                new GeoData{ Country="Sweden", State = "Skåne", Sales = 47 },
                new GeoData{ Country="Sweden", State = "Halland", Sales = 88 },
                new GeoData{ Country="Sweden", State = "Norrbotten", Sales = 99 },
                new GeoData{ Country="Sweden", State = "Västra Götaland", Sales = 120 },
                new GeoData{ Country="Sweden", State = "Södermanland", Sales = 57 },
            };

            ws.Cells["A1"].LoadFromCollection(l, true, OfficeOpenXml.Table.TableStyles.Medium12);
        }
        protected static ExcelRangeBase LoadItemData(ExcelWorksheet ws)
        {
            ws.Cells["K1"].Value = "Item";
            ws.Cells["L1"].Value = "Category";
            ws.Cells["M1"].Value = "Stock";
            ws.Cells["N1"].Value = "Price";
            ws.Cells["O1"].Value = "Date for grouping";

            ws.Cells["K2"].Value = "Crowbar";
            ws.Cells["L2"].Value = "Hardware";
            ws.Cells["M2"].Value = 12;
            ws.Cells["N2"].Value = 85.2;
            ws.Cells["O2"].Value = new DateTime(2010, 1, 31);

            ws.Cells["K3"].Value = "Crowbar";
            ws.Cells["L3"].Value = "Hardware";
            ws.Cells["M3"].Value = 15;
            ws.Cells["N3"].Value = 12.2;
            ws.Cells["O3"].Value = new DateTime(2010, 2, 28);

            ws.Cells["K4"].Value = "Hammer";
            ws.Cells["L4"].Value = "Hardware";
            ws.Cells["M4"].Value = 550;
            ws.Cells["N4"].Value = 72.7;
            ws.Cells["O4"].Value = new DateTime(2010, 3, 31);

            ws.Cells["K5"].Value = "Hammer";
            ws.Cells["L5"].Value = "Hardware";
            ws.Cells["M5"].Value = 120;
            ws.Cells["N5"].Value = 11.3;
            ws.Cells["O5"].Value = new DateTime(2010, 4, 30);

            ws.Cells["K6"].Value = "Crowbar";
            ws.Cells["L6"].Value = "Hardware";
            ws.Cells["M6"].Value = 120;
            ws.Cells["N6"].Value = 173.2;
            ws.Cells["O6"].Value = new DateTime(2010, 5, 31);

            ws.Cells["K7"].Value = "Hammer";
            ws.Cells["L7"].Value = "Hardware";
            ws.Cells["M7"].Value = 1;
            ws.Cells["N7"].Value = 4.2;
            ws.Cells["O7"].Value = new DateTime(2010, 6, 30);

            ws.Cells["K8"].Value = "Saw";
            ws.Cells["L8"].Value = "Hardware";
            ws.Cells["M8"].Value = 4;
            ws.Cells["N8"].Value = 33.12;
            ws.Cells["O8"].Value = new DateTime(2010, 6, 28);

            ws.Cells["K9"].Value = "Screwdriver";
            ws.Cells["L9"].Value = "Hardware";
            ws.Cells["M9"].Value = 1200;
            ws.Cells["N9"].Value = 45.2;
            ws.Cells["O9"].Value = new DateTime(2010, 8, 31);

            ws.Cells["K10"].Value = "Apple";
            ws.Cells["L10"].Value = "Groceries";
            ws.Cells["M10"].Value = 807;
            ws.Cells["N10"].Value = 1.2;
            ws.Cells["O10"].Value = new DateTime(2010, 9, 30);

            ws.Cells["K11"].Value = "Butter";
            ws.Cells["L11"].Value = "Groceries";
            ws.Cells["M11"].Value = 52;
            ws.Cells["N11"].Value = 7.2;
            ws.Cells["O11"].Value = new DateTime(2010, 10, 31);
            ws.Cells["O2:O11"].Style.Numberformat.Format = "yyyy-MM-dd";
            return ws.Cells["K1:O11"];
        }

        protected static void SetDateValues(ExcelWorksheet _ws, int noItems=100)
        {
            /* Set dates in numeric column */
            _ws.SetValue(50, 2, new DateTime(2018, 12, 15));
            _ws.SetValue(51, 2, new DateTime(2018, 12, 16));
            _ws.SetValue(52, 2, new DateTime(2018, 12, 17));
            _ws.Cells[50, 2, 52, 2].Style.Numberformat.Format = "yyyy-MM-dd";

            _ws.Cells[1, 1, noItems, 1].Style.Numberformat.Format = "yyyy-MM-dd";
            _ws.Cells[2, 4, noItems, 4].Style.Numberformat.Format = "#,##0.00";
        }
        protected int GetRowFromDate(DateTime date)
        {
            var startDate = new DateTime(DateTime.Today.Year-1, 11, 1);
            if (startDate > date)
                return 2;
            else
                return (date - startDate).Days + 2;
        }
        protected static ExcelWorksheet TryGetWorksheet(ExcelPackage pck, string worksheetName)
        {
            var ws = pck.Workbook.Worksheets[worksheetName];
            if (ws == null) Assert.Inconclusive($"{worksheetName} worksheet is missing");
            return ws;
        }
        protected static ExcelShape TryGetShape(ExcelPackage pck, string wsName)
        {
            var ws = pck.Workbook.Worksheets[wsName];
            if (ws == null) Assert.Inconclusive($"{wsName} worksheet is missing");
            var shape = (ExcelShape)ws.Drawings[0];
            return shape;
        }
        protected static FileInfo GetResourceFile(string fileName)
        {
            string path = AppContext.BaseDirectory;
            while (!Directory.Exists(path + "\\Resources") && path.Length > 4)
            {
                path = new DirectoryInfo(path + "\\..").FullName;
            }
            if(path.Length > 4)
            {
                return new FileInfo(path + "\\Resources\\" + fileName);
            }
            else
            {
                return null;
            }
        }
        protected void AssertIsNull(ExcelRangeBase range)
        {
            foreach (var r in range)
            {
                Assert.IsNotNull(r.Value);
            }
        }


        protected void AssertNoChange(ExcelRangeBase range)
        {
            foreach (var r in range)
            {
                Assert.AreEqual(r.Address, r.Value);
            }
        }

        protected static void SetValues(ExcelWorksheet ws, int rowcols)
        {
            for (int r = 1; r <= rowcols; r++)
            {
                for (int c = 1; c <= rowcols; c++)
                {
                    ws.Cells[r, c].Value = ExcelCellBase.GetAddress(r, c);
                }
            }
        }

    }
}

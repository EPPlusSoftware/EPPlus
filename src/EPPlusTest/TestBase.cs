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

namespace EPPlusTest
{
    [TestClass]
    public abstract class TestBase
    {
        protected static FileInfo _file;
        protected static string _clipartPath ="";
        protected static string _worksheetPath = @"c:\epplusTest\Testoutput\";
        protected static string _testInputPath = AppContext.BaseDirectory + "\\workbooks\\";
        protected static string _testInputPathOptional = @"c:\epplusTest\Testoutput\";
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

        private static void CreateWorksheetPathIfNotExists()
        {
            if (!Directory.Exists(_worksheetPath))
            {
                Directory.CreateDirectory(_worksheetPath);
            }
        }

        protected static ExcelPackage OpenTemplatePackage(string name)
        {
            var t = new FileInfo(_testInputPath  + name);
            if (t.Exists)
            {
                var _file = new FileInfo(_worksheetPath + name);
                return new ExcelPackage(_file, t);
            }
            else
            {
                t = new FileInfo(_testInputPathOptional + name);
                if (t.Exists)
                {
                    var _file = new FileInfo(_worksheetPath + name);
                    return new ExcelPackage(_file, t);
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
                fi.Delete();
            }
            pck.SaveAs(fi);
        }
        protected static readonly DateTime _loadDataStartDate = new DateTime(DateTime.Today.Year-1, 11, 1);
        protected static void LoadTestdata(ExcelWorksheet ws, int noItems = 100, int startColumn=1, int startRow=1)
        {
            ws.SetValue(1, startColumn, "Date");
            ws.SetValue(1, startColumn + 1, "NumValue");
            ws.SetValue(1, startColumn + 2, "StrValue");
            ws.SetValue(1, startColumn + 3, "NumFormatedValue");

            DateTime dt = _loadDataStartDate;
            int row = 1;
            for (int i = 1; i < noItems; i++)
            {
                row = startRow + i;
                ws.SetValue(row, startColumn, dt);
                ws.SetValue(row, startColumn + 1, row);
                ws.SetValue(row, startColumn + 2, $"Value {row}");
                ws.SetValue(row, startColumn + 3, row * 33);
                dt = dt.AddDays(1);
            }
            ws.Cells[startRow, 1, row, 1].Style.Numberformat.Format = "yyyy-MM-dd";
            ws.Cells.AutoFitColumns();
        }
        protected static void SetDateValues(ExcelWorksheet _ws, int noItems=100)
        {
            /** Set dates in numeric column **/
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

    }
}

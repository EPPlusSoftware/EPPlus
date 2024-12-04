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
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Sparkline;
using System.Threading;
using System.Collections.Generic;
using System.Drawing;

namespace EPPlusTest
{
    [TestClass]
    public class MultiThreadingTests : TestBase
    {
        static ExcelPackage _pck;
        string _pckfile;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenPackage("MultiThreading.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            var dirName = _pck.File.DirectoryName;
            var fileName = _pck.File.FullName;

            SaveAndCleanup(_pck);
            if (File.Exists(fileName))
            {
                File.Copy(fileName, dirName + "\\WorksheetRead.xlsx", true);
            }
        }
        [TestMethod]
        public void AddMultipleWorksheetsWithStyling()
        {
            var noTheads = 10;
            using(var p=OpenPackage("MulitThreadWorksheet.xlsx", true))
            {
                var pool = new List<Thread>();
                for(int i=0;i<noTheads;i++)
                {
                    var thread = new Thread(LoadWorksheet);
                    pool.Add(thread);
                    var ws = p.Workbook.Worksheets.Add($"Sheet{i + 1}");                    
                    thread.Start(ws);
                }

                while(pool.Count>0)
                {
                    if (pool[0].IsAlive==false)
                    {
                        pool.RemoveAt(0);
                        continue;
                    }
                    Thread.Sleep(100);
                }
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void AddMultipleWorksheetsWithStyling_NoThreading()
        {
            using (var p = OpenPackage("NoThreading.xlsx", true))
            {
                for (int i = 0; i < 10; i++)
                {
                    var ws = p.Workbook.Worksheets.Add($"Sheet{i + 1}");
                    LoadWorksheet(ws);
                }
                SaveAndCleanup(p);
            }
        }
        public static void LoadWorksheet(object wsObject)
        {
            var ws = wsObject as ExcelWorksheet;
            for(int r=1;r<=1000;r++)
            {
                for (int c = 1; c <= 10; c++)
                {
                    ws.Cells[r, c].Value = c + r * 10;
                    ws.Cells[r, c].Style.Font.Color.SetColor(Color.FromArgb(c + r * 10));
                    ws.Cells[r, c].Style.Fill.SetBackground(Color.FromArgb(c + r*20));
                    ws.Cells[r, c].Style.Numberformat.Format = "#,##0";
                }
            }
        }
    }
}

/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
  07/07/2023         EPPlus Software AB       Epplus 7
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.ConditionalFormatting;
using System.Threading;
using System.Drawing;
namespace EPPlusTest
{
    [TestClass]
    public class ReadTemplate : TestBase
    {
        [TestMethod]
        public void ReadBlankStream()
        {
            MemoryStream stream = new MemoryStream();
            using (ExcelPackage pck = new ExcelPackage(stream))
            {
                var ws = pck.Workbook.Worksheets.Add("Perf");
                pck.SaveAs(stream);
            }
            stream.Close();
        }
        [TestMethod]
        public void OpenXlts()
        {
            using (var pck = OpenTemplatePackage("Template.xltx"))
            {
                var ws=pck.Workbook.Worksheets[0];
                SaveWorkbook("Template.xlsx", pck);
            }
        }
    }
}

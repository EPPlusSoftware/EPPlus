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
 * but WITHOUT ANY WARRANTY; without even the implied warranty ofD:\kod\EPPlusSoftware\EPPlus\src\EPPlusTest\Workbooks\Template.xltx
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
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlusTest
{
    [TestClass]
    public class DTS_FailingTests
    {

        [TestMethod]
        public void DeleteWorksheetWithReferencedImage()
        {
            var ms = new MemoryStream();
            using (var pck = new ExcelPackage())
            {
                var ws = pck.Workbook.Worksheets.Add("original");
                ws.Drawings.AddPicture("Pic1", Properties.Resources.Test1);
                pck.Workbook.Worksheets.Copy("original", "copy");
                pck.SaveAs(ms);
            }
            ms.Position = 0;

            using (var pck = new ExcelPackage(ms))
            {
                var ws = pck.Workbook.Worksheets["original"];
                pck.Workbook.Worksheets.Delete(ws);
                pck.Save();
            }
        }

        [TestMethod]
        public void CopyAndDeleteWorksheetWithImage()
        {
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var ws = pck.Workbook.Worksheets.Add("original");
                ws.Drawings.AddPicture("Pic1", Properties.Resources.Test1);
                pck.Workbook.Worksheets.Copy("original", "copy");
                pck.Workbook.Worksheets.Delete(ws);
                pck.Save();
            }
        }
    }
}

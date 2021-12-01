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
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;

namespace EPPlusTest.Core.Worksheet
{
    [TestClass]
    public class AutofitColumnsSkiaTests : TestBase
    {
        static ExcelPackage _pck;

        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("Skia.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }

        [TestMethod]
        public void SaveCharToCellShouldBeWrittenAsString()
        {
            var ws = _pck.Workbook.Worksheets.Add("autofit");
            ws.Cells["A1"].Value = "Autofit columns Skia";
            ws.Cells["B1"].Value = "Autofit columns Skia. Autofit columns Skia";
            ws.Cells["C1"].Value = "Autofit columns Skia. Autofit columns Skia. Autofit columns Skia";
            ws.Cells["D1"].Value = "Autofit columns Skia. Autofit columns Skia. Autofit columns Skia. Autofit columns Skia";
            ws.Cells.AutoFitColumns();
        }
        [TestMethod]
        public void AddJpgImageViaSkia()
        {
            var ws = _pck.Workbook.Worksheets.Add("pictures");

            using (var ms1 = new MemoryStream(Properties.Resources.Test1JpgByteArray))
            {
                var image1 = ws.Drawings.AddPicture("jpg1", ms1, OfficeOpenXml.Drawing.ePictureType.Jpg);
            }
            using (var ms2 = new MemoryStream(Properties.Resources.VmlPatternImagePngByteArray))
            {
                var image2 = ws.Drawings.AddPicture("png1", ms2, OfficeOpenXml.Drawing.ePictureType.Png);
            }
            //using (var ms3 = Properties.Resources.VectorDrawing)
            //{
            //    var image2 = ws.Drawings.AddPicture("wmf1", ms3, OfficeOpenXml.Drawing.ePictureType.Wmf);
            //}
        }
        [TestMethod]
        public void AddJpgImageViaImageChart()
        {
            var ws = _pck.Workbook.Worksheets.Add("picturesIS");

            using (var ms1 = new MemoryStream(Properties.Resources.Test1JpgByteArray))
            {
                var image1 = ws.Drawings.AddPicture("jpg1", ms1, OfficeOpenXml.Drawing.ePictureType.Jpg);
            }
            using (var ms2 = new MemoryStream(Properties.Resources.VmlPatternImagePngByteArray))
            {
                var image2 = ws.Drawings.AddPicture("png1", ms2, OfficeOpenXml.Drawing.ePictureType.Png);
            }
            using (var ms3 = new MemoryStream(Properties.Resources.CodeEmfByteArray))
            {
                var image3 = ws.Drawings.AddPicture("emf1", ms3, OfficeOpenXml.Drawing.ePictureType.Emf);
            }
            using (var ms4 = new MemoryStream(Properties.Resources.Svg1ByteArray))
            {
                var image4 = ws.Drawings.AddPicture("svg1", ms4, OfficeOpenXml.Drawing.ePictureType.Svg);
            }
            using (var ms5 = new MemoryStream(Properties.Resources.Svg2ByteArray))
            {
                var image5 = ws.Drawings.AddPicture("svg2", ms5, OfficeOpenXml.Drawing.ePictureType.Svg);
            }

        }
    }
}

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

            using (var msGif = new MemoryStream(Properties.Resources.BitmapImageGif))
            {
                var imageGif = ws.Drawings.AddPicture("gif1", msGif, OfficeOpenXml.Drawing.ePictureType.Gif);
                imageGif.SetPosition(40, 0, 0, 0);
            }

            using (var msBmp = new MemoryStream(Properties.Resources.BitmapImageGif))
            {
                var imagebmp = ws.Drawings.AddPicture("bmp1", msBmp, OfficeOpenXml.Drawing.ePictureType.Bmp);
                imagebmp.SetPosition(40, 0, 10, 0);
            }

            using (var ms1 = new MemoryStream(Properties.Resources.Test1JpgByteArray))
            {
                var image1 = ws.Drawings.AddPicture("jpg1", ms1, OfficeOpenXml.Drawing.ePictureType.Jpg);
            }
            using (var ms2 = new MemoryStream(Properties.Resources.VmlPatternImagePngByteArray))
            {
                var image2 = ws.Drawings.AddPicture("png1", ms2, OfficeOpenXml.Drawing.ePictureType.Png);
                image2.SetPosition(0, 0, 10, 0);
            }
            using (var ms22 = new MemoryStream(Properties.Resources.Png2ByteArray))
            {
                var image22 = ws.Drawings.AddPicture("png2", ms22, OfficeOpenXml.Drawing.ePictureType.Png);
                image22.SetPosition(0, 0, 20, 0);
            }
            using (var ms23 = new MemoryStream(Properties.Resources.Png3ByteArray))
            {
                var image23 = ws.Drawings.AddPicture("png3", ms23, OfficeOpenXml.Drawing.ePictureType.Png);
                image23.SetPosition(0, 0, 30, 0);
            }
            using (var ms3 = new MemoryStream(Properties.Resources.CodeEmfByteArray))
            {
                var image3 = ws.Drawings.AddPicture("emf1", ms3, OfficeOpenXml.Drawing.ePictureType.Emf);
                image3.SetPosition(0, 0, 40, 0);
            }
            using (var ms4 = new MemoryStream(Properties.Resources.Svg1ByteArray))
            {
                var image4 = ws.Drawings.AddPicture("svg1", ms4, OfficeOpenXml.Drawing.ePictureType.Svg);
                image4.SetPosition(0, 0, 50, 0);
            }
            using (var ms5 = new MemoryStream(Properties.Resources.Svg2ByteArray))
            {
                var image5 = ws.Drawings.AddPicture("svg2", ms5, OfficeOpenXml.Drawing.ePictureType.Svg);
                image5.SetPosition(0, 0, 60, 0);
                image5.SetSize(25);
            }
            using (var ms6 = Properties.Resources.VectorDrawing)
            {
                var image6 = ws.Drawings.AddPicture("wmf", ms6, OfficeOpenXml.Drawing.ePictureType.Wmf);
                image6.SetPosition(0, 0, 70, 0);
            }
            
            using (var msTif = Properties.Resources.CodeTif)
            {
                var imageTif = ws.Drawings.AddPicture("tif1", msTif, OfficeOpenXml.Drawing.ePictureType.Tif);
                imageTif.SetPosition(0, 0, 80, 0);
            }
        }
        [TestMethod]
        public void AddWebPImages()
        {
            var ws = _pck.Workbook.Worksheets.Add("picturesWebP");

            //386*395
            using (var msWebP1 = GetImageMemoryStream("2_webp_a.webp"))
            {
                var imageWebP1 = ws.Drawings.AddPicture("webp1", msWebP1, OfficeOpenXml.Drawing.ePictureType.WebP);
                imageWebP1.SetPosition(0, 0, 0, 0);
            }

            //386*395
            using (var msWebP1 = GetImageMemoryStream("2_webp_ll.webp"))
            {
                var imageWebP1 = ws.Drawings.AddPicture("webp2", msWebP1, OfficeOpenXml.Drawing.ePictureType.WebP);
                imageWebP1.SetPosition(0, 0, 10, 0);
            }

            //400*400
            using (var msWebP1 = GetImageMemoryStream("animated.webp"))
            {
                var imageWebP1 = ws.Drawings.AddPicture("webp3", msWebP1, OfficeOpenXml.Drawing.ePictureType.WebP);
                imageWebP1.SetPosition(0, 0, 20, 0);
            }


            //320*214
            using (var msWebP1 = GetImageMemoryStream("1.sm.webp"))
            {
                var imageWebP1 = ws.Drawings.AddPicture("webp4-1", msWebP1, OfficeOpenXml.Drawing.ePictureType.WebP);
                imageWebP1.SetPosition(20, 0, 0, 0);
            }

            //320*214
            using (var msWebP1 = GetImageMemoryStream("2.sm.webp"))
            {
                var imageWebP1 = ws.Drawings.AddPicture("webp4-2", msWebP1, OfficeOpenXml.Drawing.ePictureType.WebP);
                imageWebP1.SetPosition(20, 0, 10, 0);
            }

            //320*214
            using (var msWebP1 = GetImageMemoryStream("3.sm.webp"))
            {
                var imageWebP1 = ws.Drawings.AddPicture("webp4-3", msWebP1, OfficeOpenXml.Drawing.ePictureType.WebP);
                imageWebP1.SetPosition(20, 0, 20, 0);
            }

            //320*214
            using (var msWebP1 = GetImageMemoryStream("4.sm.webp"))
            {
                var imageWebP1 = ws.Drawings.AddPicture("webp4-4", msWebP1, OfficeOpenXml.Drawing.ePictureType.WebP);
                imageWebP1.SetPosition(20, 0, 30, 0);
            }

            //320*214
            using (var msWebP1 = GetImageMemoryStream("5.sm.webp"))
            {
                var imageWebP1 = ws.Drawings.AddPicture("webp4-5", msWebP1, OfficeOpenXml.Drawing.ePictureType.WebP);
                imageWebP1.SetPosition(20, 0, 40, 0);
            }
        }
        [TestMethod]
        public void AddIconImages()
        {
            var ws = _pck.Workbook.Worksheets.Add("Icon");
            //32*32
            using (var msIco1 = GetImageMemoryStream("1_32x32.ico"))
            {
                var imageWebP1 = ws.Drawings.AddPicture("ico1", msIco1, OfficeOpenXml.Drawing.ePictureType.Ico);
                imageWebP1.SetPosition(40, 0, 0, 0);
            }
            //128*128
            using (var msIco2 = GetImageMemoryStream("1_128x128.ico"))
            {
                var imageWebP1 = ws.Drawings.AddPicture("ico2", msIco2, OfficeOpenXml.Drawing.ePictureType.Ico);
                imageWebP1.SetPosition(40, 0, 10, 0);
            }
        }
        [TestMethod]
        public void AddEmfImages()
        {
            var ws = _pck.Workbook.Worksheets.Add("Emz");
            //32*32
            using (var msIco1 = GetImageMemoryStream("example.emz"))
            {
                var imageWebP1 = ws.Drawings.AddPicture("Emf", msIco1, OfficeOpenXml.Drawing.ePictureType.Emz);
                imageWebP1.SetPosition(40, 0, 0, 0);
            }
        }
    }
}

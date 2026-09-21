/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using OfficeOpenXml.Drawing;

namespace OfficeOpenXml.Export.PdfExport.Data
{
    internal class PdfDrawing
    {
        public ExcelPicture Picture { get; }
        public byte[] ImageBytes => Picture.Image.ImageBytes;
        public ePictureType PictureType => Picture.Image.Type.Value;
        public PdfDrawing(ExcelPicture picture) { Picture = picture; }
    }
}

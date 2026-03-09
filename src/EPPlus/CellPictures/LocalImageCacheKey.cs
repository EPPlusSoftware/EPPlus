/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/11/2024         EPPlus Software AB       Initial release EPPlus 8
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.CellPictures
{
    internal class LocalImageCacheKey : PictureCacheKey
    {
        private readonly Uri imageUri;
        private readonly CalcOrigins calcOrigin;
        private readonly string altText;

        public LocalImageCacheKey(Uri imageUri, CalcOrigins calcOrigin, string altText)
        {
            this.imageUri = imageUri;
            this.calcOrigin = calcOrigin;
            this.altText = altText;
        }

        public LocalImageCacheKey(ExcelCellPicture picture) 
            : this(picture.ImageUri, picture.CalcOrigin, picture.AltText)
        {
        }

        protected override string Build()
        {
            var sb = new StringBuilder();
            sb.Append("LocalImage-");
            sb.Append(imageUri.OriginalString);
            sb.Append(calcOrigin.ToString());
            sb.Append(altText ?? string.Empty);
            return sb.ToString();
        }
    }
}

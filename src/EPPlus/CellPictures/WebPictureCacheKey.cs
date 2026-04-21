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
using OfficeOpenXml.RichData.RichValues.WebImages;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.CellPictures
{
    internal class WebPictureCacheKey : PictureCacheKey
    {
        private readonly Uri addressUri;
        private readonly string altText;
        private readonly CalcOrigins calcOrigin;
        private readonly WebImageSizing sizing;
        private readonly double? height;
        private readonly double? width;

        public WebPictureCacheKey(Uri addressUri, string altText, CalcOrigins calcOrigin, WebImageSizing sizing, double? height, double? width = null)
        {
            this.addressUri = addressUri;
            this.altText = altText;
            this.calcOrigin = calcOrigin;
            this.sizing = sizing;
            this.height = height;
            this.width = width;
        }

        public WebPictureCacheKey(ExcelCellPicture pic)
            : this(pic.ExternalAddress, pic.AltText, pic.CalcOrigin, pic.Sizing ?? WebImageSizing.FitToCellMaintainRatio, null)
        {
            
        }

        protected override string Build()
        {
            var sb = new StringBuilder();
            sb.Append("WebImage-");
            sb.Append(addressUri.OriginalString);
            sb.Append(altText ?? string.Empty);
            sb.Append(calcOrigin.ToString());
            sb.Append(sizing.ToString());
            sb.Append(height ?? -1);
            sb.Append(width ?? -1);
            return sb.ToString();
        }
    }
}

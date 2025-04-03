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
using OfficeOpenXml.CellPictures;
using OfficeOpenXml.RichData.Structures.Constants;
using System;

namespace OfficeOpenXml.RichData.RichValues.WebImages
{
    internal class WebImageRichValue : ExcelRichValue
    {
        public WebImageRichValue(RichDataDatabase richDataDb) : base(richDataDb, RichDataStructureTypes.WebImage)
        {
            _richDataDb = richDataDb;
        }

        private readonly RichDataDatabase _richDataDb;

        public Uri ImageUri
        {
            get
            {
                if(WebImageIdentifier.HasValue)
                {
                    var img = _richDataDb.WebImages.Get(WebImageIdentifier.Value);
                    return img.Blip;
                }
                return null;
            }
        }

        public Uri ExternalAddressUri
        {
            get
            {
                if (WebImageIdentifier.HasValue)
                {
                    var img = _richDataDb.WebImages.Get(WebImageIdentifier.Value);
                    return img.Address;
                }
                return null;
            }
        }

        internal uint? WebImageIdentifier
        {
            get
            {
                var id = GetValueInt(StructureKeyNames.WebImage.WebImageIdentifier);
                if(!id.HasValue)
                {
                    return null;
                }
                return (uint)id;
            }
            set
            {
                SetValue(StructureKeyNames.WebImage.WebImageIdentifier, value);
            }
        }

        internal override void PostProcessInitialRead()
        {
            base.PostProcessInitialRead();
            if(WebImageIdentifier.HasValue)
            {
                var imgId = _richDataDb.WebImages.GetIdByIndex((int)WebImageIdentifier.Value);
                WebImageIdentifier = imgId;
            }
        }

        /// <summary>
        /// String representation of the alt text associated with the image.
        /// </summary>
        public string Text 
        {
            get 
            {
                return GetValue(StructureKeyNames.WebImage.Text);
            } 
            set 
            {
                SetValue(StructureKeyNames.WebImage.Text, value);
            }
        }

        /// <summary>
        /// Boolean value that when true indicates the image can be generated on demand.
        /// </summary>
        public bool ComputedImage
        {
            get
            {
                return GetValueBool(StructureKeyNames.WebImage.ComputedImage) ?? false;
            }
            set
            {
                SetValue(StructureKeyNames.WebImage.ComputedImage, value);
            }
        }

        /// <summary>
        /// Integer value that indicates the dimension modes of the image. There are 4 supported values: 0, 1, 2, 3.
        /// At 0, image fits to the cell and maintains its aspect ratio.
        /// At 1, image fits to the cell and ignore its aspect ratio.
        /// At 2, image maintains its original size and MAY exceed the cell boundary.
        /// At 3, image size is customized by "ImageHeight" and "ImageWidth".
        /// </summary>
        public WebImageSizing ImageSizing
        {
            get
            {
                var v = GetValueInt(StructureKeyNames.WebImage.ImageSizing);
                return v.HasValue && v >= 0 && v <= 3 ?(WebImageSizing)v : WebImageSizing.FitToCellMaintainRatio;
            }
            set
            {
                SetValue(StructureKeyNames.WebImage.ImageSizing, (int?)value);
            }
        }

        /// <summary>
        /// Real number representation of the image height in pixels. This property SHOULD only be present when "ImageSizing" is set to 3.
        /// </summary>
        public double? ImageHeight
        {
            get
            {
                if (ImageSizing != WebImageSizing.CustomizeByHeightAndWidth) return null;
                return GetValueDouble(StructureKeyNames.WebImage.ImageHeight);
            }
            set
            {
                SetValue(StructureKeyNames.WebImage.ImageHeight, value);
            }
        }

        /// <summary>
        /// Real number representation of the image width in pixels. This property SHOULD only be present when "ImageSizing" is set to 3
        /// </summary>
        public double? ImageWidth
        {
            get
            {
                if (ImageSizing != WebImageSizing.CustomizeByHeightAndWidth) return null;
                return GetValueDouble(StructureKeyNames.WebImage.ImageWidth);
            }
            set
            {
                SetValue(StructureKeyNames.WebImage.ImageWidth, value);
            }
        }

        /// <summary>
        /// Integer value that indicates how the rich value was created.
        /// </summary>
        public CalcOrigins? CalcOrigin
        {
            get
            {
                var val = GetValueInt(StructureKeyNames.WebImage.CalcOrigin);
                if (val.HasValue)
                {
                    return (CalcOrigins)val;
                }
                return null;
            }
            set
            {
                SetValue(StructureKeyNames.WebImage.CalcOrigin, (int?)value);
            }
        }
    }
}
